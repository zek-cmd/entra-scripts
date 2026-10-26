#Requires -Version 5.1
<#
.SYNOPSIS
    Bulk delete Microsoft Entra ID (Azure AD) users listed in an Excel (.xlsx) file.

.DESCRIPTION
    Reads a column of user identifiers from an .xlsx file, resolves each one against
    your Entra ID tenant, and deletes (or disables) the matching user accounts.

    Safety features:
      * Group scope     -> a target group is ALWAYS required. Only users who are
                           current members of that group are deleted; every other
                           row is skipped and logged as NotInGroup. This bounds the
                           blast radius if the wrong spreadsheet is selected.
      * -WhatIf         -> dry run. Resolves and prints every target, deletes nothing.
      * -Confirm        -> interactive prompt before each delete.
      * -DisableOnly    -> disable accounts instead of deleting them (reversible).
      * License cleanup -> removes assigned licenses first, so deletes do not fail.
      * Audit log       -> appends a CSV log of every action attempted.

    Works with either the modern Microsoft.Graph module or the legacy AzureAD module.

.PARAMETER ExcelPath
    Path to the .xlsx file containing the list of users.
    If omitted, a graphical Open File Dialog will appear for selection.

.PARAMETER ColumnName
    Header name of the column holding the user identifier.
    Examples in this folder: 'FULLNAME' (sam.xlsx), 'ID' (SOFTNET.xlsx).

.PARAMETER IdentifierType
    How the column value is matched against Entra:
      UPN         - UserPrincipalName (exact match)
      DisplayName - displayName (exact match)
      Mail        - mail (exact match)
      ObjectId    - Entra object id

.PARAMETER DisableOnly
    Disable the accounts instead of deleting them. Reversible.

.PARAMETER SkipLicenseCleanup
    Do not remove licenses before deleting. Faster, but deletes may fail for licensed users.

.PARAMETER GroupName
    Display name of the group that scopes the delete. Only members of this group
    are eligible. Accepts an exact name or an unambiguous substring; if omitted,
    or if nothing matches, an interactive numbered group picker is shown.
    This scope is mandatory - there is no "delete regardless of group" mode.

.PARAMETER Source
    Where the list of users comes from:
      Excel - read identifiers from a spreadsheet (prompts for file + column)
      Group - use every member of a selected group
    If omitted, an interactive menu asks. Passing -ExcelPath implies Excel.

.PARAMETER GroupLimit
    Maximum number of groups to load into the picker. Default 500.

.PARAMETER LogPath
    Audit log CSV path. Defaults to 'BulkDelete_Entra_<timestamp>.csv' next to the script.

.EXAMPLE
    # Open File Dialog lets you pick the .xlsx interactively, then dry-run
    .\Bulk-Delete-Entra-Users.ps1 -ColumnName FULLNAME -IdentifierType DisplayName -WhatIf

.EXAMPLE
    # Specify the Excel path directly
    .\Bulk-Delete-Entra-Users.ps1 -ExcelPath .\sam.xlsx -ColumnName FULLNAME -IdentifierType DisplayName -WhatIf

.EXAMPLE
    # LIVE run - prompts for confirmation on each delete
    .\Bulk-Delete-Entra-Users.ps1 -ExcelPath .\sam.xlsx -ColumnName FULLNAME -IdentifierType DisplayName

.EXAMPLE
    # Safer: disable instead of delete
    .\Bulk-Delete-Entra-Users.ps1 -ExcelPath .\sam.xlsx -ColumnName FULLNAME -IdentifierType DisplayName -DisableOnly
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)]
    [string]$ExcelPath,

    [Parameter(Mandatory = $false)]
    [string]$ColumnName,

    [ValidateSet('UPN', 'DisplayName', 'Mail', 'ObjectId')]
    [string]$IdentifierType = 'UPN',

    [string]$NameSeparator = ' ',

    [switch]$DisableOnly,

    [switch]$SkipLicenseCleanup,

    [string]$GroupName,

    [ValidateSet('Excel', 'Group')]
    [string]$Source,

    [int]$GroupLimit = 500,

    [string]$LogPath
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Shared, portable Python resolver (dot-sourced so its function lands in this scope).
# Optional: the inline fallback in Get-PythonExecutable still works without it.
$script:PyResolverPath = Join-Path $scriptRoot 'Resolve-PythonWithOpenpyxl.ps1'
if (Test-Path -LiteralPath $script:PyResolverPath) { . $script:PyResolverPath }

# ExcelPath is resolved later (may come from the Open File Dialog)
if (-not [string]::IsNullOrWhiteSpace($ExcelPath)) {
    if (-not (Test-Path -LiteralPath $ExcelPath)) { throw "Excel file not found: $ExcelPath" }
    $ExcelPath = (Resolve-Path -LiteralPath $ExcelPath).Path
}

if (-not $LogPath) {
    $LogPath = Join-Path (Join-Path $scriptRoot 'Output Report') ("BulkDelete_Entra_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
$script:LogFile = $LogPath

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Write-Log {
    param(
        [string]$Action, [string]$Identifier, [string]$DisplayName,
        [string]$ObjectId, [string]$Result, [string]$Message, [string]$Group = ''
    )
    $row = [pscustomobject]@{
        Timestamp   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Action      = $Action
        Identifier  = $Identifier
        DisplayName = $DisplayName
        ObjectId    = $ObjectId
        Group       = if ($Group) { $Group } else { $script:ScopeGroupName }
        Result      = $Result
        Message     = $Message
    }
    # Written via .NET so -WhatIf does not suppress the audit trail.
    $fields = @(
        $row.Timestamp, $row.Action, $row.Identifier, $row.DisplayName,
        $row.ObjectId, $row.Group, $row.Result, ($row.Message -replace "`r|`n", ' ')
    ) | ForEach-Object { '"{0}"' -f ([string]$_ -replace '"', '""') }

    $line = ($fields -join ',') + [Environment]::NewLine
    $enc = New-Object System.Text.UTF8Encoding($true)

    $dir = Split-Path -Parent $script:LogFile
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }

    if (Test-Path -LiteralPath $script:LogFile) {
        [System.IO.File]::AppendAllText($script:LogFile, $line, $enc)
    } else {
        $header = '"Timestamp","Action","Identifier","DisplayName","ObjectId","Group","Result","Message"' + [Environment]::NewLine
        [System.IO.File]::WriteAllText($script:LogFile, $header + $line, $enc)
    }
}

function Get-PythonExecutable {
    <# Returns a Python interpreter that can import openpyxl, or $null.
       Uses the shared Resolve-PythonWithOpenpyxl helper when available, and falls
       back to the original inline probe if that file is missing. #>
    if (Get-Command Resolve-PythonWithOpenpyxl -ErrorAction SilentlyContinue) {
        $resolved = Resolve-PythonWithOpenpyxl
        if ($resolved) { return $resolved }
        return $null
    }

    # --- Fallback: original inline probe (kept for a standalone copy of this file) ---
    # 1) DSH bundled python, built from the profile rather than a hardcoded user name
    $candidates = @(
        (Join-Path $env:USERPROFILE '.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

    foreach ($c in $candidates) {
        & $c -c "import openpyxl" 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { return $c }
    }

    # 2) Any python on PATH that has openpyxl (skip 0-byte Store alias stubs)
    foreach ($cmd in @('python', 'python3', 'py')) {
        foreach ($found in @(Get-Command $cmd -All -ErrorAction SilentlyContinue)) {
            if (-not $found.Source -or -not (Test-Path -LiteralPath $found.Source)) { continue }
            $item = Get-Item -LiteralPath $found.Source -ErrorAction SilentlyContinue
            if ($item -and $item.Length -eq 0) { continue }
            & $found.Source -c "import openpyxl" 2>&1 | Out-Null
            if ($LASTEXITCODE -eq 0) { return $found.Source }
        }
    }
    return $null
}

function Read-ExcelColumn {
    <# Returns an array of non-empty values from $Column(s) in $File.
       Multiple columns are joined with $Separator. #>
    param([string]$File, [string[]]$Column, [string]$Separator = ' ')

    $cols = @($Column)
    if ($cols.Count -eq 0) { return @() }

    # --- Primary: Python + openpyxl ---
    $python = Get-PythonExecutable
    if ($python) {
        $pyScript = Join-Path ([System.IO.Path]::GetTempPath()) ("read_col_{0}.py" -f ([guid]::NewGuid().ToString('N')))
        $pyOut = Join-Path ([System.IO.Path]::GetTempPath()) ("results_{0}.txt" -f ([guid]::NewGuid().ToString('N')))
        $pyCode = @'
import openpyxl, sys, csv, os
path, sep = sys.argv[1], sys.argv[2]
cols = sys.argv[3:]
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb.active
headers = [str(c.value or "").strip() for c in ws[1]]
missing = [c for c in cols if c not in headers]
if missing:
    sys.stderr.write("COLUMN_NOT_FOUND: %s\n" % ", ".join(missing))
    sys.exit(2)
idxs = [headers.index(c) for c in cols]
out_path = os.environ.get('PYPYOUT', '')
if out_path:
    out_file = open(out_path, 'w', newline='', encoding='utf-8')
    out = csv.writer(out_file, lineterminator='\n')
else:
    out = csv.writer(sys.stdout, lineterminator='\n')
for row in ws.iter_rows(min_row=2, values_only=True):
    parts = []
    for i in idxs:
        v = row[i] if i < len(row) else None
        s = "" if v is None else str(v).strip()
        if s:
            parts.append(s)
    joined = sep.join(parts)
    if joined:
        out.writerow([joined])
'@
        # Use .NET IO so -WhatIf does not suppress the write
        [System.IO.File]::WriteAllText($pyScript, $pyCode)

        try {
            $env:PYPYOUT = $pyOut
            $null = & $python $pyScript $File $Separator @cols 2>&1
            $pyExit = $LASTEXITCODE
            if ($pyExit -eq 2) { throw "Column(s) not found in $File : $($cols -join ', ')" }
            if ($pyExit -eq 0) {
                $raw = @(Get-Content -LiteralPath $pyOut -Encoding UTF8 -ErrorAction SilentlyContinue |
                    Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace($_) })
                if ($raw.Count -gt 0) { return $raw }
            }
            Write-Verbose "openpyxl read returned exit $pyExit; falling back to ImportExcel."
        }
        finally {
            [System.IO.File]::Delete($pyScript)
            if (Test-Path -LiteralPath $pyOut) { Remove-Item -LiteralPath $pyOut -Force -ErrorAction SilentlyContinue }
        }
    }

    # --- Fallback: Import-Excel module ---
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        $data = Import-Excel -Path $File -ErrorAction Stop
        $i = 0
        return @($data | ForEach-Object {
            $i++
            $parts = foreach ($c in $cols) {
                $v = $_.$c
                if ($null -ne $v -and ([string]$v).Trim() -ne '') { ([string]$v).Trim() }
            }
            $joined = ($parts -join $Separator)
            if ($joined -ne '') { $joined }
        })
    }

    throw @"
Cannot read the Excel file. Install one of:
  python -m pip install openpyxl
  Install-Module ImportExcel -Scope CurrentUser
"@
}

function Show-OpenExcelDialog {
    <# Opens a Windows Forms file picker. Returns the selected path, or $null on cancel. #>
    Add-Type -AssemblyName System.Windows.Forms | Out-Null

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title      = 'Select Excel file with user list'
    $dialog.Filter     = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    $dialog.FilterIndex = 1
    $dialog.InitialDirectory = if (Test-Path -LiteralPath $scriptRoot) { $scriptRoot } else { $env:USERPROFILE }
    $dialog_multiselect = $false

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    return $null
}

function Get-ExcelColumnHeaders {
    <# Returns header names from the first sheet.
       Python + openpyxl first, ImportExcel as the fallback. #>
    param([string]$File)

    $python = Get-PythonExecutable
    if ($python) {
        $pyScript = Join-Path ([System.IO.Path]::GetTempPath()) ("get_headers_{0}.py" -f ([guid]::NewGuid().ToString('N')))
        $pyOut = Join-Path ([System.IO.Path]::GetTempPath()) ("results_hdr_{0}.txt" -f ([guid]::NewGuid().ToString('N')))
        $pyCode = @'
import openpyxl, sys, os
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
ws = wb.active
headers = [str(c.value or "").strip() for c in ws[1]]
out_path = os.environ.get('PYPYOUT', '')
if out_path:
    with open(out_path, 'w', encoding='utf-8') as f:
        for h in headers:
            if h:
                f.write(h + '\n')
else:
    for h in headers:
        if h:
            print(h)
'@
        # Use .NET IO so -WhatIf does not suppress the write
        [System.IO.File]::WriteAllText($pyScript, $pyCode)
        try {
            $env:PYPYOUT = $pyOut
            $null = & $python $pyScript $File 2>&1
            $pyExit = $LASTEXITCODE
            if ($pyExit -eq 0) {
                $hdrs = @(Get-Content -LiteralPath $pyOut -Encoding UTF8 -ErrorAction SilentlyContinue |
                    Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace($_) })
                if ($hdrs.Count -gt 0) { return $hdrs }
            }
            Write-Verbose "openpyxl header read returned exit $pyExit; falling back to ImportExcel."
        } finally {
            [System.IO.File]::Delete($pyScript)
            if (Test-Path -LiteralPath $pyOut) { [System.IO.File]::Delete($pyOut) }
        }
    }

    # Fallback: Import-Excel
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        $data = @(Import-Excel -Path $File -ErrorAction Stop)
        if ($data.Count -gt 0) { return @($data[0].PSObject.Properties.Name) }
    }
    return @()
}

function Get-GraphProvider {
    if (Get-Module -ListAvailable -Name Microsoft.Graph.Users) { return 'Graph' }
    if (Get-Module -ListAvailable -Name Microsoft.Graph)        { return 'Graph' }
    if (Get-Module -ListAvailable -Name AzureAD)                { return 'AzureAD' }
    return $null
}

function Resolve-EntraUser {
    <# Returns the matching user object(s), or $null if not found. #>
    param([string]$Value, [string]$Type, [string]$Provider)

    $esc = $Value.Replace("'", "''")

    if ($Provider -eq 'Graph') {
        switch ($Type) {
            'UPN'         { return @(Get-MgUser -Filter "userPrincipalName eq '$esc'" -ErrorAction SilentlyContinue) }
            'DisplayName' { return @(Get-MgUser -Filter "displayName eq '$esc'" -ErrorAction SilentlyContinue) }
            'Mail'        { return @(Get-MgUser -Filter "mail eq '$esc'" -ErrorAction SilentlyContinue) }
            'ObjectId'    { return @(Get-MgUser -UserId $Value -ErrorAction SilentlyContinue) }
        }
    }
    else {
        switch ($Type) {
            'UPN'         { return @(Get-AzureADUser -Filter "userPrincipalName eq '$esc'" -ErrorAction SilentlyContinue) }
            'DisplayName' { return @(Get-AzureADUser -Filter "displayName eq '$esc'" -ErrorAction SilentlyContinue) }
            'Mail'        { return @(Get-AzureADUser -Filter "mail eq '$esc'" -ErrorAction SilentlyContinue) }
            'ObjectId'    { return @(Get-AzureADUser -ObjectId $Value -ErrorAction SilentlyContinue) }
        }
    }
    return @()
}

function Remove-UserLicenses {
    param($User, [string]$Provider)
    try {
        if ($Provider -eq 'Graph') {
            $lic = @(Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction SilentlyContinue)
            foreach ($l in $lic) {
                if ($l.Id) { Remove-MgUserLicenseDetail -UserId $User.Id -LicenseDetailsId $l.Id -ErrorAction SilentlyContinue | Out-Null }
            }
        }
        else {
            Get-AzureADUserLicenseDetail -ObjectId $User.ObjectId -ErrorAction SilentlyContinue |
                Where-Object { $_.SkuId } |
                ForEach-Object {
                    Remove-AzureADUserLicense -ObjectId $User.ObjectId -SkuId $_.SkuId -ErrorAction SilentlyContinue
                }
        }
    } catch {
        Write-Verbose "License cleanup skipped: $_"
    }
}

function Get-UserId {
    param($User, [string]$Provider)
    if ($Provider -eq 'Graph') { return $User.Id }
    return $User.ObjectId
}

function Get-UserDisplay {
    param($User)
    if ($User.DisplayName) { return $User.DisplayName }
    if ($User.Mail) { return $User.Mail }
    if ($User.UserPrincipalName) { return $User.UserPrincipalName }
    return ''
}

# ---------------------------------------------------------------------------
# Group scoping helpers (delete safety boundary)
# ---------------------------------------------------------------------------
function Get-EntraGroups {
    <# Returns all groups (security + M365) with Id and DisplayName. #>
    param([string]$Provider, [int]$Limit)
    if ($Provider -eq 'Graph') {
        $g = @(Get-MgGroup -All -ErrorAction SilentlyContinue | Select-Object Id, DisplayName, GroupTypes, Mail)
    } else {
        $g = @(Get-AzureADGroup -ErrorAction SilentlyContinue | Select-Object ObjectId, DisplayName, Mail)
    }
    $g = @($g | Where-Object { $_.DisplayName })
    if ($Limit -gt 0 -and $g.Count -gt $Limit) { $g = $g | Select-Object -First $Limit }
    return $g
}

function Get-GroupId {
    param($Group, [string]$Provider)
    if ($Provider -eq 'Graph') { return $Group.Id }
    return $Group.ObjectId
}

function Get-EntraGroupByName {
    <# Exact name first, then an unambiguous substring. Returns $null if not found
       or if a substring matches more than one group. #>
    param([string]$Name, $Groups)

    $exact = @($Groups | Where-Object { $_.DisplayName -eq $Name })
    if ($exact.Count -eq 1) { return $exact[0] }
    if ($exact.Count -gt 1) { return $null }

    $partial = @($Groups | Where-Object { $_.DisplayName -like "*$Name*" })
    if ($partial.Count -eq 1) { return $partial[0] }
    return $null
}

function Select-GroupInteractive {
    <# Shows a numbered menu of groups and returns the selected group object. #>
    param(
        [string]$Prompt,
        $Groups
    )

    Write-Host ''
    Write-Host ("=== {0} ===" -f $Prompt) -ForegroundColor Cyan
    Write-Host ' Available groups:' -ForegroundColor DarkGray

    $i = 0
    foreach ($g in $Groups) {
        $i++
        $type = ''
        if ($g.PSObject.Properties['GroupTypes'] -and $g.GroupTypes) {
            if ($g.GroupTypes -contains 'Unified') { $type = ' [M365]' }
            elseif ($g.GroupTypes -contains 'DynamicMembership') { $type = ' [Dynamic]' }
            else { $type = ' [Security]' }
        }
        Write-Host ("  [{0}] {1}{2}" -f $i, $g.DisplayName, $type) -ForegroundColor Gray
    }

    while ($true) {
        $choice = (Read-Host ("Enter number for {0} (or type part of a name to filter)" -f $Prompt)).Trim()
        if ($choice -eq '') { continue }

        if ($choice -match '^\d+$') {
            $idx = [int]$choice
            if ($idx -ge 1 -and $idx -le $Groups.Count) { return $Groups[$idx - 1] }
            Write-Host '  Invalid number.' -ForegroundColor Red
            continue
        }

        $matches = @($Groups | Where-Object { $_.DisplayName -like "*$choice*" })
        if ($matches.Count -eq 1) { return $matches[0] }
        if ($matches.Count -eq 0) {
            Write-Host ("  No group matches '{0}'." -f $choice) -ForegroundColor Red
            continue
        }

        Write-Host ("  {0} matches:" -f $matches.Count) -ForegroundColor Yellow
        $j = 0
        foreach ($m in $matches) {
            $j++
            Write-Host ("   [{0}] {1}" -f $j, $m.DisplayName) -ForegroundColor Gray
        }
        $pick = (Read-Host '   Enter number').Trim()
        if ($pick -match '^\d+$' -and [int]$pick -ge 1 -and [int]$pick -le $matches.Count) {
            return $matches[[int]$pick - 1]
        }
        Write-Host '  Invalid selection.' -ForegroundColor Red
    }
}

function Get-GroupMemberIds {
    <# Returns the object ids of every member of the group. #>
    param([string]$GroupId, [string]$Provider)
    if ($Provider -eq 'Graph') {
        return @(Get-MgGroupMember -GroupId $GroupId -All -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    }
    return @(Get-AzureADGroupMember -ObjectId $GroupId -ErrorAction SilentlyContinue | ForEach-Object { $_.ObjectId })
}

function Get-GroupMemberUsers {
    <# Returns group members that are actually user accounts.
       A group can contain nested groups, devices and service principals; those are
       not users and must never be fed to a delete loop. Returns objects with
       Id / DisplayName / UserPrincipalName.

       IMPORTANT: Get-MgGroupMember returns IMicrosoftGraphDirectoryObject, whose
       only real property is Id - it has NO DisplayName or UserPrincipalName. That
       is why members used to print as empty strings. Get-MgGroupMemberAsUser asks
       Graph for the user-typed cast instead: it returns full User objects AND
       excludes nested groups / devices / service principals server-side. #>
    param([string]$GroupId, [string]$Provider)

    $users = @()
    $skipped = 0

    if ($Provider -eq 'Graph') {
        if (Get-Command Get-MgGroupMemberAsUser -ErrorAction SilentlyContinue) {
            $rows = @(Get-MgGroupMemberAsUser -GroupId $GroupId -All -ErrorAction SilentlyContinue)
            foreach ($r in $rows) {
                if (-not $r.Id) { $skipped++; continue }
                $users += [pscustomobject]@{
                    Id                = $r.Id
                    DisplayName       = $r.DisplayName
                    UserPrincipalName = $r.UserPrincipalName
                }
            }
        } else {
            # Older Microsoft.Graph without the *AsUser cast: fall back to the base
            # DirectoryObject list and resolve each id to a real user.
            $rows = @(Get-MgGroupMember -GroupId $GroupId -All -ErrorAction SilentlyContinue)
            foreach ($r in $rows) {
                if (-not $r.Id) { $skipped++; continue }
                $u = $null
                try { $u = Get-MgUser -UserId $r.Id -ErrorAction SilentlyContinue } catch { $u = $null }
                if (-not $u) { $skipped++; continue }   # nested group / device / SP
                $users += [pscustomobject]@{
                    Id                = $u.Id
                    DisplayName       = $u.DisplayName
                    UserPrincipalName = $u.UserPrincipalName
                }
            }
        }
    } else {
        # AzureAD already returns typed objects, so DisplayName/UPN are populated.
        $rows = @(Get-AzureADGroupMember -ObjectId $GroupId -ErrorAction SilentlyContinue)
        foreach ($r in $rows) {
            if ($r.ObjectType -and [string]$r.ObjectType -ne 'User') { $skipped++; continue }
            $id = if ($r.ObjectId) { $r.ObjectId } else { $r.Id }
            if (-not $id) { $skipped++; continue }
            $dn = $r.DisplayName
            $upn = $r.UserPrincipalName
            if (-not $dn -or -not $upn) {
                # Some tenants return only the id; fill in from the user object.
                $u = $null
                try { $u = Get-AzureADUser -ObjectId $id -ErrorAction SilentlyContinue } catch { $u = $null }
                if ($u) {
                    if (-not $dn) { $dn = $u.DisplayName }
                    if (-not $upn) { $upn = $u.UserPrincipalName }
                }
            }
            $users += [pscustomobject]@{
                Id                = $id
                DisplayName       = $dn
                UserPrincipalName = $upn
            }
        }
    }

    if ($skipped -gt 0) {
        Write-Host (" Skipped {0} non-user member(s) (nested groups / devices / service principals)." -f $skipped) -ForegroundColor DarkYellow
    }
    return @($users)
}

function Resolve-ScopeGroup {
    <# Resolves the mandatory scope group to a verified object id.
       Accepts an interactive picker, an exact/substring name, or a pre-selected
       group object. Throws rather than guessing. #>
    param(
        [string]$Provider,
        [string]$Name,
        $Groups,
        $Preselected
    )

    if ($Preselected) {
        $id = Get-GroupId -Group $Preselected -Provider $Provider
        if ($id) { return $Preselected }
    }

    $group = $null
    if ($Name) {
        if ($Groups.Count -gt 0) {
            $group = Get-EntraGroupByName -Name $Name -Groups $Groups
            if (-not $group) {
                Write-Host (" No unique group matches '{0}' - choose interactively." -f $Name) -ForegroundColor Yellow
            }
        } else {
            # Cannot enumerate (permissions). Accept the name, resolve it directly.
            $group = [pscustomobject]@{ DisplayName = $Name }
        }
    }

    if (-not $group) {
        if ($Groups.Count -eq 0) {
            throw @"
A group scope is required, but no groups could be listed.
Grant Group.Read.All / Directory.Read.All, or pass -GroupName '<exact group name>'.
There is no unscoped delete mode in this script.
"@
        }
        $group = Select-GroupInteractive -Prompt 'scope group' -Groups $Groups
    }

    # Verify the group resolves to a real object id
    $gid = Get-GroupId -Group $group -Provider $Provider
    if (-not $gid -and $group.DisplayName) {
        $esc = $group.DisplayName.Replace("'", "''")
        if ($Provider -eq 'Graph') {
            $found = @(Get-MgGroup -Filter "displayName eq '$esc'" -ErrorAction SilentlyContinue)
            if ($found.Count -ge 1) { $gid = $found[0].Id }
        } else {
            $found = @(Get-AzureADGroup -Filter "displayName eq '$esc'" -ErrorAction SilentlyContinue)
            if ($found.Count -ge 1) { $gid = $found[0].ObjectId }
        }
    }

    if (-not $gid) {
        throw ("Could not resolve an object id for group '{0}'. Refusing to run without a verified scope." -f $group.DisplayName)
    }
    return $group
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Source selection: Excel file, or members of a group
# ---------------------------------------------------------------------------
# -ExcelPath implies Excel, so existing command lines keep working unchanged.
if (-not $Source) {
    if (-not [string]::IsNullOrWhiteSpace($ExcelPath)) {
        $Source = 'Excel'
    } else {
        Write-Host ''
        Write-Host '=== Select user source ===' -ForegroundColor Cyan
        Write-Host '  [1] Excel file - read users from a spreadsheet' -ForegroundColor Gray
        Write-Host '  [2] Group     - use all members of a selected group' -ForegroundColor Gray
        $srcMode = ''
        while ($srcMode -notmatch '^[12]$') {
            $srcMode = (Read-Host 'Enter 1 or 2').Trim()
            if ($srcMode -notmatch '^[12]$') { Write-Host '  Please enter 1 or 2.' -ForegroundColor Red }
        }
        $Source = if ($srcMode -eq '1') { 'Excel' } else { 'Group' }
    }
}
$useExcel = ($Source -eq 'Excel')
Write-Host ''
Write-Host (" Source selected : {0}" -f $(if ($useExcel) { 'Excel file' } else { 'Group members' })) -ForegroundColor Cyan

if ($useExcel) {
    # Use Open File Dialog if ExcelPath was not provided
    if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
        $ExcelPath = Show-OpenExcelDialog
        if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
            Write-Host 'No Excel file selected. Exiting.' -ForegroundColor Red
            exit 1
        }
    } elseif (-not (Test-Path -LiteralPath $ExcelPath)) {
        Write-Host ("File not found: {0}" -f $ExcelPath) -ForegroundColor Red
        $ExcelPath = Show-OpenExcelDialog
        if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
            Write-Host 'No Excel file selected. Exiting.' -ForegroundColor Red
            exit 1
        }
    }

    $ExcelPath = (Resolve-Path -LiteralPath $ExcelPath).Path
    Write-Host (" Selected file : {0}" -f $ExcelPath) -ForegroundColor DarkYellow

    # Auto-detect column if not specified
    if (-not $ColumnName -or [string]::IsNullOrWhiteSpace($ColumnName)) {
        $headers = @(Get-ExcelColumnHeaders -File $ExcelPath)
        if ($headers.Count -eq 0) { throw "No columns detected in $ExcelPath." }
        Write-Host 'Detected columns:' -ForegroundColor DarkGray
        for ($i = 0; $i -lt $headers.Count; $i++) { Write-Host ("  [{0}] {1}" -f $i, ([string]$headers[$i])) -ForegroundColor DarkGray }
        $choice = Read-Host "Select column index, name, or combo (e.g. FIRST NAME + LASTNAME)"
        $choice = $choice.Trim().Trim('"').Trim("'")
        if ($choice -match '^\d+$') {
            $ColumnName = [string]$headers[[int]$choice]
        } else {
            $ColumnName = [string]$choice
        }
        Write-Host (" Using column : {0}" -f $ColumnName) -ForegroundColor Cyan
    }

    # Auto-set IdentifierType based on column name unless user explicitly specified it
    if (-not $PSBoundParameters.ContainsKey('IdentifierType')) {
        $colLower = $ColumnName.ToLower()
        if ($colLower -match 'fullname|name' -and -not ($colLower -match 'id')) { $IdentifierType = 'DisplayName'; Write-Host " Auto-detected identifier type: DisplayName" -ForegroundColor DarkGray }
        elseif ($colLower -match 'mail|email') { $IdentifierType = 'Mail'; Write-Host " Auto-detected identifier type: Mail" -ForegroundColor DarkGray }
        elseif ($colLower -match 'id' -and -not ($colLower -match 'name')) { $IdentifierType = 'ObjectId'; Write-Host " Auto-detected identifier type: ObjectId" -ForegroundColor DarkGray }
        else { $IdentifierType = 'UPN'; Write-Host " Auto-detected identifier type: UPN" -ForegroundColor DarkGray }
    }
} else {
    # Group mode: the group is the source, so no spreadsheet or column is involved.
    if ($ColumnName) {
        Write-Host ' -ColumnName is ignored when the source is a group.' -ForegroundColor DarkYellow
        $ColumnName = $null
    }
}

$provider = Get-GraphProvider
if (-not $provider) {
    throw @"
No Entra ID module found. Install one of:
  Install-Module Microsoft.Graph -Scope CurrentUser     (preferred)
  Install-Module AzureAD -Scope CurrentUser
"@
}

if ($provider -eq 'Graph') {
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -Scopes 'User.ReadWrite.All', 'User.Read.All', 'Directory.Read.All', 'GroupMember.Read.All', 'Group.Read.All'
    } else {
        Write-Host ("Already connected to Graph as: {0}" -f $ctx.Account) -ForegroundColor DarkGray
    }
} else {
    if (-not (Get-AzureADTenant -ErrorAction SilentlyContinue)) {
        Write-Host "Connecting to Azure AD..." -ForegroundColor Cyan
        Connect-AzureAD
    }
}

# ---------------------------------------------------------------------------
# Target group selection
#
# Excel mode : this group is a SCOPE FILTER. Users from the sheet that are not
#              members are skipped and logged as NotInGroup.
# Group mode : this group IS the user list, so the membership filter would be
#              circular. Instead every member is a target, and the scope check
#              is reported as not-applicable rather than silently tautological.
# ---------------------------------------------------------------------------
Write-Host ''
if ($useExcel) {
    Write-Host '=== Select scope group (required) ===' -ForegroundColor Cyan
    Write-Host ' Only members of this group can be deleted. Everyone else is skipped.' -ForegroundColor DarkGray
} else {
    Write-Host '=== Select group whose members will be deleted ===' -ForegroundColor Cyan
    Write-Host ' Every user account in this group becomes a delete target.' -ForegroundColor DarkGray
}

$scopeGroup = $null
$allGroups = @()
try {
    $allGroups = @(Get-EntraGroups -Provider $provider -Limit $GroupLimit)
} catch {
    Write-Host (" Could not enumerate groups: {0}" -f $_.Exception.Message) -ForegroundColor Red
}

$scopeGroup = Resolve-ScopeGroup -Provider $provider -Name $GroupName -Groups $allGroups
if ($scopeGroup) { Write-Host (" Using group : {0}" -f $scopeGroup.DisplayName) -ForegroundColor Cyan }

$scopeGroupId = Get-GroupId -Group $scopeGroup -Provider $provider
$script:ScopeGroupName = $scopeGroup.DisplayName

if (-not $scopeGroupId) {
    throw ("Could not resolve an object id for group '{0}'. Refusing to run without a verified scope." -f $scopeGroup.DisplayName)
}

# Membership is required in both modes: as a filter (Excel) or as the work list (Group).
$scopeMemberIds = @(Get-GroupMemberIds -GroupId $scopeGroupId -Provider $provider)
if ($scopeMemberIds.Count -eq 0) {
    throw ("Group '{0}' has no members (or membership could not be read). Nothing would be eligible - refusing to run." -f $scopeGroup.DisplayName)
}

$scopeLookup = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
foreach ($mid in $scopeMemberIds) { if ($mid) { [void]$scopeLookup.Add([string]$mid) } }

Write-Host (" Scope group    : {0}" -f $scopeGroup.DisplayName) -ForegroundColor Green
Write-Host (" Group members  : {0}" -f $scopeLookup.Count) -ForegroundColor Green

Write-Host ''
Write-Host '==================================================' -ForegroundColor Cyan
Write-Host ' Bulk Delete - Microsoft Entra ID Users' -ForegroundColor Cyan
Write-Host '==================================================' -ForegroundColor Cyan
if ($useExcel) {
    Write-Host (" Source         : Excel file")
    Write-Host (" Excel file     : {0}" -f $ExcelPath)
    Write-Host (" Column         : {0}" -f $ColumnName)
} else {
    Write-Host (" Source         : Group members")
}
Write-Host (" Identifier     : {0}" -f $(if ($useExcel) { $IdentifierType } else { 'ObjectId (group members)' }))
Write-Host (" Scope group    : {0}" -f $scopeGroup.DisplayName)
Write-Host (" Action         : {0}" -f $(if ($DisableOnly) { 'DISABLE (reversible)' } else { 'DELETE (permanent)' }))
Write-Host (" Log            : {0}" -f $LogPath)
Write-Host (" WhatIf active  : {0}" -f $WhatIfPreference)
Write-Host '--------------------------------------------------'

if ($useExcel) {
    # Handle composite column selection (e.g. "FIRST NAME + LASTNAME")
    $columns = @($ColumnName)
    if ($ColumnName -match ' \+ ') {
        $columns = $ColumnName.Split(@(' + '), [System.StringSplitOptions]::None) | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        Write-Host (" Composite columns: {0}" -f ($columns -join ' | ')) -ForegroundColor DarkGray
    }

    $values = @(Read-ExcelColumn -File $ExcelPath -Column $columns -Separator $NameSeparator |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique)

    if ($values.Count -eq 0) {
        Write-Warning "No non-empty values found in column '$ColumnName'. Nothing to do."
        exit 0
    }
} else {
    # Group mode: members are the targets. Only real user objects are eligible.
    Write-Host (" Fetching members of '{0}'..." -f $scopeGroup.DisplayName) -ForegroundColor DarkGray
    $memberUsers = @(Get-GroupMemberUsers -GroupId $scopeGroupId -Provider $provider)

    if ($memberUsers.Count -eq 0) {
        Write-Warning "Group '$($scopeGroup.DisplayName)' has no user members. Nothing to do."
        exit 0
    }

    $values = @($memberUsers | ForEach-Object {
        if ($_.UserPrincipalName) { $_.UserPrincipalName }
        elseif ($_.DisplayName)   { $_.DisplayName }
        else                      { $_.Id }
    } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)

    # In group mode identifiers are resolved by object id, so pre-seed a map to
    # avoid a directory lookup per member.
    $memberByKey = @{}
    foreach ($m in $memberUsers) {
        $k = if ($m.UserPrincipalName) { $m.UserPrincipalName }
             elseif ($m.DisplayName)   { $m.DisplayName }
             else                      { $m.Id }
        if ($k -and -not $memberByKey.ContainsKey([string]$k)) { $memberByKey[[string]$k] = $m }
    }
    $IdentifierType = 'ObjectId'
}

Write-Host (" Rows to process : {0}" -f $values.Count)
Write-Host '--------------------------------------------------'

$stats = [ordered]@{ Resolved = 0; Deleted = 0; Disabled = 0; NotFound = 0; Ambiguous = 0; NotInGroup = 0; Failed = 0; Skipped = 0 }
$index = 0

foreach ($value in $values) {
    $index++

    if ($useExcel) {
        $users = @(Resolve-EntraUser -Value $value -Type $IdentifierType -Provider $provider)
    } else {
        # Group mode: the member object is already known - no directory lookup.
        $known = $null
        if ($memberByKey.ContainsKey([string]$value)) { $known = $memberByKey[[string]$value] }
        $users = if ($known) { @($known) } else { @() }
    }

    if ($users.Count -eq 0) {
        $stats.NotFound++
        Write-Host ("[{0,3}/{1}] NOT FOUND  : {2}" -f $index, $values.Count, $value) -ForegroundColor DarkYellow
        Write-Log -Action 'Resolve' -Identifier $value -DisplayName '' -ObjectId '' -Result 'NotFound' -Message 'No matching Entra user'
        continue
    }

    if ($users.Count -gt 1) {
        $stats.Ambiguous++
        Write-Host ("[{0,3}/{1}] AMBIGUOUS  : '{2}' matched {3} users - skipped" -f $index, $values.Count, $value, $users.Count) -ForegroundColor Yellow
        Write-Log -Action 'Resolve' -Identifier $value -DisplayName '' -ObjectId '' -Result 'Ambiguous' -Message "Matched $($users.Count) users"
        continue
    }

    $user = $users[0]
    $oid = Get-UserId -User $user -Provider $provider
    $disp = Get-UserDisplay -User $user
    # Never leave the audit line blank: fall back to the identifier, then the id.
    if ([string]::IsNullOrWhiteSpace($disp)) {
        $disp = if (-not [string]::IsNullOrWhiteSpace($value)) { $value } else { [string]$oid }
    }
    $stats.Resolved++

    # --- Group scope enforcement (Excel source only) ------------------------
    # A user that is not a current member of the scope group is never touched.
    # In group mode every target already came from that group, so this check
    # would be circular and is not applied.
    if ($useExcel -and -not $scopeLookup.Contains([string]$oid)) {
        $stats.NotInGroup++
        Write-Host ("[{0,3}/{1}] NOT IN GROUP : {2}" -f $index, $values.Count, $disp) -ForegroundColor DarkYellow
        Write-Log -Action 'Resolve' -Identifier $value -DisplayName $disp -ObjectId $oid -Result 'NotInGroup' -Message ("Not a member of '{0}'" -f $script:ScopeGroupName)
        continue
    }

    if ($DisableOnly) {
        if ($PSCmdlet.ShouldProcess($disp, "Disable account (UPN: $($user.UserPrincipalName))")) {
            try {
                if ($provider -eq 'Graph') { Update-MgUser -UserId $oid -AccountEnabled:$false -ErrorAction Stop }
                else { Set-AzureADUser -ObjectId $oid -AccountEnabled $false -ErrorAction Stop }
                $stats.Disabled++
                Write-Host ("[{0,3}/{1}] DISABLED  : {2}" -f $index, $values.Count, $disp) -ForegroundColor Yellow
                Write-Log -Action 'Disable' -Identifier $value -DisplayName $disp -ObjectId $oid -Result 'Success' -Message ''
            } catch {
                $stats.Failed++
                Write-Host ("[{0,3}/{1}] FAILED    : {2} -> {3}" -f $index, $values.Count, $disp, $_.Exception.Message) -ForegroundColor Red
                Write-Log -Action 'Disable' -Identifier $value -DisplayName $disp -ObjectId $oid -Result 'Failed' -Message $_.Exception.Message
            }
        } else {
            $stats.Skipped++
            Write-Host ("[{0,3}/{1}] WHATIF    : would disable {2}" -f $index, $values.Count, $disp) -ForegroundColor DarkCyan
            Write-Log -Action 'Disable' -Identifier $value -DisplayName $disp -ObjectId $oid -Result 'WhatIf' -Message 'Not executed'
        }
        continue
    }

    # DELETE path
    if ($PSCmdlet.ShouldProcess($disp, "DELETE account (UPN: $($user.UserPrincipalName))")) {
        try {
            if (-not $SkipLicenseCleanup) { Remove-UserLicenses -User $user -Provider $provider }
            if ($provider -eq 'Graph') { Remove-MgUser -UserId $oid -ErrorAction Stop }
            else { Remove-AzureADUser -ObjectId $oid -ErrorAction Stop }
            $stats.Deleted++
            Write-Host ("[{0,3}/{1}] DELETED   : {2}" -f $index, $values.Count, $disp) -ForegroundColor Red
            Write-Log -Action 'Delete' -Identifier $value -DisplayName $disp -ObjectId $oid -Result 'Success' -Message ''
        } catch {
            $stats.Failed++
            Write-Host ("[{0,3}/{1}] FAILED    : {2} -> {3}" -f $index, $values.Count, $disp, $_.Exception.Message) -ForegroundColor Red
            Write-Log -Action 'Delete' -Identifier $value -DisplayName $disp -ObjectId $oid -Result 'Failed' -Message $_.Exception.Message
        }
    } else {
        $stats.Skipped++
        Write-Host ("[{0,3}/{1}] WHATIF    : would delete {2}" -f $index, $values.Count, $disp) -ForegroundColor DarkCyan
        Write-Log -Action 'Delete' -Identifier $value -DisplayName $disp -ObjectId $oid -Result 'WhatIf' -Message 'Not executed'
    }
}

Write-Host ''
Write-Host '--------------------------------------------------'
Write-Host 'Results:' -ForegroundColor Cyan
Write-Host (" Scope group : {0}" -f $script:ScopeGroupName)
Write-Host (" Resolved  : {0}" -f $stats.Resolved)
Write-Host (" Deleted   : {0}" -f $stats.Deleted)
Write-Host (" Disabled  : {0}" -f $stats.Disabled)
Write-Host (" NotFound  : {0}" -f $stats.NotFound)
Write-Host (" Ambiguous : {0}" -f $stats.Ambiguous)
Write-Host (" NotInGroup: {0}" -f $stats.NotInGroup)
Write-Host (" Failed    : {0}" -f $stats.Failed)
Write-Host (" WhatIf    : {0}" -f $stats.Skipped)
Write-Host (" Audit log : {0}" -f $LogPath)
if ($WhatIfPreference) {
    Write-Host 'WhatIf mode was ACTIVE - nothing was deleted.' -ForegroundColor Cyan
} else {
    Write-Host 'Live execution completed.' -ForegroundColor Cyan
}