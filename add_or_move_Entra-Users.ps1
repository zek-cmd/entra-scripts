#Requires -Version 5.1
<#
.SYNOPSIS
    Transfer users between Microsoft Entra ID (Azure AD) groups using a list from an Excel file.

.DESCRIPTION
    Reads user identifiers from an .xlsx file, then moves each matching user from one
    group ("From" group) to another group ("To" group).

    Both groups are chosen interactively from a menu of your tenant's groups,
    matched by display name. Group names do NOT come from the spreadsheet.

    Safety features:
      * -WhatIf       -> dry run. Lists every action, changes nothing.
      * -Confirm      -> prompts before each change.
      * -Force        -> add to the "To" group even if the user is not in the "From" group.
      * Audit log     -> CSV log of every action attempted.

    Works with either the modern Microsoft.Graph module or the legacy AzureAD module.

.PARAMETER ExcelPath
    Path to the .xlsx file containing the list of users.
    If omitted, a graphical Open File Dialog appears for selection.

.PARAMETER ColumnName
    Header name of the column holding the user identifier.
    Omit to auto-detect and choose interactively.
    Multiple columns can be combined with ' + ', e.g. 'FIRST NAME + LASTNAME'.

.PARAMETER IdentifierType
    How the column value is matched against Entra:
      UPN | DisplayName | Mail | ObjectId
    Auto-detected from the column name when not specified.

.PARAMETER NameSeparator
    Separator used when joining multiple columns. Default is a single space.

.PARAMETER Force
    Add users to the "To" group even when they are not currently members of the "From" group.

.PARAMETER GroupLimit
    Maximum number of groups listed in the selection menu. Default 500.

.PARAMETER LogPath
    Audit log CSV path. Defaults to 'GroupMove_<timestamp>.csv' next to the script.

.EXAMPLE
    # Dry run - safe, changes nothing
    .\Move-Entra-GroupMembers.ps1 -WhatIf

.EXAMPLE
    # Live run with confirmation prompts
    .\Move-Entra-GroupMembers.ps1 -Confirm

.EXAMPLE
    # Explicit columns
    .\Move-Entra-GroupMembers.ps1 -ExcelPath .\sam.xlsx -ColumnName FULLNAME -WhatIf
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

    [switch]$Force,

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

if (-not [string]::IsNullOrWhiteSpace($ExcelPath)) {
    if (-not (Test-Path -LiteralPath $ExcelPath)) { throw "Excel file not found: $ExcelPath" }
    $ExcelPath = (Resolve-Path -LiteralPath $ExcelPath).Path
}

if (-not $LogPath) {
    $LogPath = Join-Path (Join-Path $scriptRoot 'Output Report') ("GroupMove_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
$script:LogFile = $LogPath

# ---------------------------------------------------------------------------
# Audit log (written via .NET so -WhatIf does not suppress it)
# ---------------------------------------------------------------------------
function Write-Log {
    param(
        [string]$Action, [string]$Identifier, [string]$DisplayName,
        [string]$UserId, [string]$FromGroup, [string]$ToGroup,
        [string]$Result, [string]$Message
    )
    $line = @(
        $Action, $Identifier, $DisplayName, $UserId, $FromGroup, $ToGroup, $Result,
        ($Message -replace "`r|`n", ' '), (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    ) -join '","'
    $line = '"' + $line + '"'
    $enc = New-Object System.Text.UTF8Encoding($true)
    # Create the target folder first: the default path points at 'Output Report',
    # which does not exist on a fresh checkout (bulk delete already does this).
    $dir = Split-Path -Parent $script:LogFile
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
    [System.IO.File]::AppendAllText($script:LogFile, $line + [Environment]::NewLine, $enc)
}

# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------
function Get-PythonExecutable {
    <# Returns a Python interpreter that can import openpyxl, or $null.
       Uses the shared Resolve-PythonWithOpenpyxl helper when available, and falls
       back to the original inline probe if that file is missing. #>
    if (Get-Command Resolve-PythonWithOpenpyxl -ErrorAction SilentlyContinue) {
        $resolved = Resolve-PythonWithOpenpyxl
        if ($resolved) { return $resolved }
        # Helper present but found nothing - do not silently retry the old paths,
        # or a broken install would appear to work only sometimes.
        return $null
    }

    # --- Fallback: original inline probe (kept for a standalone copy of this file) ---
    $candidates = @(
        (Join-Path $env:USERPROFILE '.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

    foreach ($c in $candidates) {
        & $c -c "import openpyxl" 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { return $c }
    }
    foreach ($cmd in @('python', 'python3', 'py')) {
        foreach ($found in @(Get-Command $cmd -All -ErrorAction SilentlyContinue)) {
            if (-not $found.Source -or -not (Test-Path -LiteralPath $found.Source)) { continue }
            # Skip 0-byte Microsoft Store alias stubs (they exit 9009)
            $item = Get-Item -LiteralPath $found.Source -ErrorAction SilentlyContinue
            if ($item -and $item.Length -eq 0) { continue }
            & $found.Source -c "import openpyxl" 2>&1 | Out-Null
            if ($LASTEXITCODE -eq 0) { return $found.Source }
        }
    }
    return $null
}

function Show-OpenExcelDialog {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Select Excel file with user list'
    $dialog.Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    $dialog.FilterIndex = 1
    $dialog.InitialDirectory = if (Test-Path -LiteralPath $scriptRoot) { $scriptRoot } else { $env:USERPROFILE }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.FileName }
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
        if ($data.Count -gt 0) {
            return @($data[0].PSObject.Properties.Name)
        }
    }
    return @()
}

function Read-ExcelColumn {
    <# Returns non-empty values from $Column(s) in $File, joining multiple columns.
       Python + openpyxl is tried FIRST; ImportExcel is the fallback.
       openpyxl is the primary path because it is the one actually present on the
       target machines, and it reads cell values directly without module quirks. #>
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
        [System.IO.File]::WriteAllText($pyScript, $pyCode)
        try {
            $env:PYPYOUT = $pyOut
            $null = & $python $pyScript $File $Separator @cols 2>&1
            $pyExit = $LASTEXITCODE
            # Exit 2 = the column genuinely does not exist. That is a definitive
            # answer, so surface it instead of silently trying another reader.
            if ($pyExit -eq 2) { throw "Column(s) not found in $File : $($cols -join ', ')" }
            if ($pyExit -eq 0) {
                $raw = @(Get-Content -LiteralPath $pyOut -Encoding UTF8 -ErrorAction SilentlyContinue |
                    Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace($_) })
                if ($raw.Count -gt 0) { return $raw }
            }
            Write-Verbose "openpyxl read returned exit $pyExit; falling back to ImportExcel."
        } finally {
            [System.IO.File]::Delete($pyScript)
            if (Test-Path -LiteralPath $pyOut) { [System.IO.File]::Delete($pyOut) }
        }
    }

    # --- Fallback: Import-Excel module ---
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        $data = Import-Excel -Path $File -ErrorAction Stop
        return @($data | ForEach-Object {
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

# ---------------------------------------------------------------------------
# Entra helpers
# ---------------------------------------------------------------------------
function Get-GraphProvider {
    if (Get-Module -ListAvailable -Name Microsoft.Graph.Users) { return 'Graph' }
    if (Get-Module -ListAvailable -Name Microsoft.Graph) { return 'Graph' }
    if (Get-Module -ListAvailable -Name AzureAD) { return 'AzureAD' }
    return $null
}

function Resolve-EntraUser {
    param([string]$Value, [string]$Type, [string]$Provider)
    $esc = $Value.Replace("'", "''")
    if ($Provider -eq 'Graph') {
        switch ($Type) {
            'UPN' { return @(Get-MgUser -Filter "userPrincipalName eq '$esc'" -ErrorAction SilentlyContinue) }
            'DisplayName' { return @(Get-MgUser -Filter "displayName eq '$esc'" -ErrorAction SilentlyContinue) }
            'Mail' { return @(Get-MgUser -Filter "mail eq '$esc'" -ErrorAction SilentlyContinue) }
            'ObjectId' { return @(Get-MgUser -UserId $Value -ErrorAction SilentlyContinue) }
        }
    } else {
        switch ($Type) {
            'UPN' { return @(Get-AzureADUser -Filter "userPrincipalName eq '$esc'" -ErrorAction SilentlyContinue) }
            'DisplayName' { return @(Get-AzureADUser -Filter "displayName eq '$esc'" -ErrorAction SilentlyContinue) }
            'Mail' { return @(Get-AzureADUser -Filter "mail eq '$esc'" -ErrorAction SilentlyContinue) }
            'ObjectId' { return @(Get-AzureADUser -ObjectId $Value -ErrorAction SilentlyContinue) }
        }
    }
    return @()
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

function Get-EntraGroups {
    <# Returns all groups (security + M365) with Id and DisplayName. #>
    param([string]$Provider, [int]$Limit)
    if ($Provider -eq 'Graph') {
        $g = @(Get-MgGroup -All -ErrorAction SilentlyContinue | Select-Object Id, DisplayName, GroupTypes, Mail)
    } else {
        $g = @(Get-AzureADGroup -ErrorAction SilentlyContinue | Select-Object ObjectId, DisplayName, Mail)
    }
    if ($Limit -gt 0 -and $g.Count -gt $Limit) { $g = $g | Select-Object -First $Limit }
    return $g
}

function Get-GroupId {
    param($Group, [string]$Provider)
    if ($Provider -eq 'Graph') { return $Group.Id }
    return $Group.ObjectId
}

function Select-GroupInteractive {
    <# Shows a numbered menu of groups and returns the selected group object. #>
    param(
        [string]$Prompt,
        [string]$Provider,
        [int]$Limit,
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
    param([string]$GroupId, [string]$Provider)
    if ($Provider -eq 'Graph') { return @(Get-MgGroupMember -GroupId $GroupId -All -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }
    return @(Get-AzureADGroupMember -ObjectId $GroupId -ErrorAction SilentlyContinue | ForEach-Object { $_.ObjectId })
}

function Get-GroupMemberUsers {
    <# Returns group members that are actually user accounts, with names populated.
       A group can also contain nested groups, devices and service principals; those
       are not users and are excluded.

       IMPORTANT: Get-MgGroupMember returns IMicrosoftGraphDirectoryObject, whose only
       real property is Id - it has NO DisplayName or UserPrincipalName. Reading those
       yields $null, which is why members used to print as blank. Get-MgGroupMemberAsUser
       asks Graph for the user-typed cast instead: it returns full User objects AND
       excludes non-user members server-side. #>
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
                    Mail              = $r.Mail
                    UserPrincipalName = $r.UserPrincipalName
                }
            }
        } else {
            # Older Microsoft.Graph without the *AsUser cast: resolve each id instead.
            $rows = @(Get-MgGroupMember -GroupId $GroupId -All -ErrorAction SilentlyContinue)
            foreach ($r in $rows) {
                if (-not $r.Id) { $skipped++; continue }
                $u = $null
                try { $u = Get-MgUser -UserId $r.Id -ErrorAction SilentlyContinue } catch { $u = $null }
                if (-not $u) { $skipped++; continue }   # nested group / device / SP
                $users += [pscustomobject]@{
                    Id                = $u.Id
                    DisplayName       = $u.DisplayName
                    Mail              = $u.Mail
                    UserPrincipalName = $u.UserPrincipalName
                }
            }
        }
    } else {
        $rows = @(Get-AzureADGroupMember -ObjectId $GroupId -ErrorAction SilentlyContinue)
        foreach ($r in $rows) {
            if ($r.ObjectType -and [string]$r.ObjectType -ne 'User') { $skipped++; continue }
            $id = if ($r.ObjectId) { $r.ObjectId } else { $r.Id }
            if (-not $id) { $skipped++; continue }
            $dn = $r.DisplayName
            $upn = $r.UserPrincipalName
            $mail = $r.Mail
            if (-not $dn -or -not $upn) {
                $u = $null
                try { $u = Get-AzureADUser -ObjectId $id -ErrorAction SilentlyContinue } catch { $u = $null }
                if ($u) {
                    if (-not $dn) { $dn = $u.DisplayName }
                    if (-not $upn) { $upn = $u.UserPrincipalName }
                    if (-not $mail) { $mail = $u.Mail }
                }
            }
            $users += [pscustomobject]@{
                Id                = $id
                DisplayName       = $dn
                Mail              = $mail
                UserPrincipalName = $upn
            }
        }
    }

    if ($skipped -gt 0) {
        Write-Host (" Skipped {0} non-user member(s) (nested groups / devices / service principals)." -f $skipped) -ForegroundColor DarkYellow
    }
    return @($users)
}
function Add-GroupMember {
    param([string]$GroupId, [string]$MemberId, [string]$Provider)
    if ($Provider -eq 'Graph') { New-MgGroupMember -GroupId $GroupId -DirectoryObjectId $MemberId -ErrorAction Stop | Out-Null }
    else { Add-AzureADGroupMember -ObjectId $GroupId -RefObjectId $MemberId -ErrorAction Stop | Out-Null }
}
function Remove-GroupMember {
    param([string]$GroupId, [string]$MemberId, [string]$Provider)
    if ($Provider -eq 'Graph') { Remove-MgGroupMemberByRef -GroupId $GroupId -DirectoryObjectId $MemberId -ErrorAction Stop | Out-Null }
    else { Remove-AzureADGroupMember -ObjectId $GroupId -MemberId $MemberId -ErrorAction Stop | Out-Null }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
# Provider
$provider = Get-GraphProvider
if (-not $provider) { throw "No Entra module found." }
if ($provider -eq 'Graph') {
    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) { Connect-MgGraph -Scopes 'User.Read.All','GroupMember.ReadWrite.All','Directory.Read.All' }
} else {
    if (-not (Get-AzureADTenant -ErrorAction SilentlyContinue)) { Connect-AzureAD }
}

# Group list
$allGroups = @(Get-EntraGroups -Provider $provider -Limit $GroupLimit)
if ($allGroups.Count -eq 0) { throw "No groups found or module error." }

# --- Source selection: Excel or Group ---
Write-Host ''
Write-Host '=== Select user source ===' -ForegroundColor Cyan
Write-Host '  [1] Excel file - read users from a spreadsheet' -ForegroundColor Gray
Write-Host '  [2] Group     - use all members of a selected group' -ForegroundColor Gray
$srcMode = ''
while ($srcMode -notmatch '^[12]$') {
    $srcMode = (Read-Host 'Enter 1 or 2').Trim()
    if ($srcMode -notmatch '^[12]$') { Write-Host '  Please enter 1 or 2.' -ForegroundColor Red }
}
$useExcel = ($srcMode -eq '1')

# --- Operation mode ---
Write-Host ''
Write-Host '=== Select operation ===' -ForegroundColor Cyan
if ($useExcel) {
    Write-Host '  [1] Move  - remove users from a source group, then add to a target group' -ForegroundColor Gray
    Write-Host '  [2] Add   - add users to a target group (no source group)' -ForegroundColor Gray
    $mode = ''
    while ($mode -notmatch '^[12]$') {
        $mode = (Read-Host 'Enter 1 or 2').Trim()
        if ($mode -notmatch '^[12]$') { Write-Host '  Please enter 1 or 2.' -ForegroundColor Red }
    }
} else {
    Write-Host '  Group source supports MOVE only (members of the picked group are moved to the target).' -ForegroundColor DarkGray
    $mode = '1'
}
$isMove = ($mode -eq '1')
Write-Host (" Operation   : {0}" -f $(if ($isMove) { 'MOVE (from source group to target group)' } else { 'ADD (add to target group only)' })) -ForegroundColor Cyan

# --- Build the user list ---
$userList = @()      # each: @{ Identifier=..; UserObj=.. }
$groupSource = $null

if ($useExcel) {
    # Resolve Excel
    if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
        $ExcelPath = Show-OpenExcelDialog
        if ([string]::IsNullOrWhiteSpace($ExcelPath)) { Write-Host 'No file selected.' -ForegroundColor Red; exit 1 }
    } elseif (-not (Test-Path -LiteralPath $ExcelPath)) {
        Write-Host ("File not found: {0}" -f $ExcelPath) -ForegroundColor Red
        $ExcelPath = Show-OpenExcelDialog
        if ([string]::IsNullOrWhiteSpace($ExcelPath)) { Write-Host 'No file selected.' -ForegroundColor Red; exit 1 }
    }
    $ExcelPath = (Resolve-Path -LiteralPath $ExcelPath).Path

    # Column detection / selection
    if ([string]::IsNullOrWhiteSpace($ColumnName)) {
        $headers = @(Get-ExcelColumnHeaders -File $ExcelPath)
        if ($headers.Count -eq 0) { throw "No columns detected." }
        Write-Host ' Detected columns:' -ForegroundColor DarkGray
        for ($i = 0; $i -lt $headers.Count; $i++) { Write-Host ("  [{0}] {1}" -f $i, $headers[$i]) -ForegroundColor DarkGray }
        $choice = Read-Host "Select index, name, or combo (e.g. FIRST NAME + LASTNAME)"
        $choice = $choice.Trim().Trim('"').Trim("'")
        if ($choice -match '^\d+$') { $ColumnName = [string]$headers[[int]$choice] } else { $ColumnName = [string]$choice }
        Write-Host (" Using column: {0}" -f $ColumnName) -ForegroundColor Cyan
    }

    # Auto-detect identifier
    if (-not $PSBoundParameters.ContainsKey('IdentifierType')) {
        $cl = $ColumnName.ToLower()
        if ($cl -match 'fullname|name' -and -not ($cl -match 'id')) { $IdentifierType = 'DisplayName' }
        elseif ($cl -match 'mail|email') { $IdentifierType = 'Mail' }
        elseif ($cl -match 'id' -and -not ($cl -match 'name')) { $IdentifierType = 'ObjectId' }
        else { $IdentifierType = 'UPN' }
        Write-Host (" Auto-detected identifier: {0}" -f $IdentifierType) -ForegroundColor DarkGray
    }

    # Handle composite
    $columns = @($ColumnName)
    if ($ColumnName -match ' \+ ') {
        $columns = $ColumnName.Split(@(' + '), [System.StringSplitOptions]::None) | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        Write-Host (" Composite columns: {0}" -f ($columns -join ' | ')) -ForegroundColor DarkGray
    }

    $values = @(Read-ExcelColumn -File $ExcelPath -Column $columns -Separator $NameSeparator |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($values.Count -eq 0) { Write-Warning 'No users found.'; exit 0 }
    $userList = @($values | ForEach-Object { @{ Identifier = $_; UserObj = $null } })
    Write-Host (" Users to process: {0} (from Excel)" -f $userList.Count) -ForegroundColor Cyan
} else {
    # Group source: the selected group's members are the users.
    # Only real user accounts are pulled in (nested groups / devices / service
    # principals are excluded), and names come from typed User objects so the
    # identifier is never blank.
    Write-Host ''
    $groupSource = Select-GroupInteractive -Prompt 'Select group whose members will be used' -Provider $provider -Limit $GroupLimit -Groups $allGroups
    $gsId = Get-GroupId -Group $groupSource -Provider $provider
    Write-Host " Fetching members of '$($groupSource.DisplayName)'..." -ForegroundColor DarkGray
    $members = @(Get-GroupMemberUsers -GroupId $gsId -Provider $provider)
    $userList = @($members | ForEach-Object {
        $d = if ($_.UserPrincipalName) { $_.UserPrincipalName }
             elseif ($_.DisplayName)   { $_.DisplayName }
             elseif ($_.Mail)          { $_.Mail }
             else                      { $_.Id }
        @{ Identifier = $d; UserObj = $_ }
    })
    if ($userList.Count -eq 0) { Write-Warning 'Group has no user members.'; exit 0 }
    Write-Host (" Users to process: {0} (members of '{1}')" -f $userList.Count, $groupSource.DisplayName) -ForegroundColor Cyan
}

# --- Group menus ---
$fromGroup = $null
if ($isMove) {
    Write-Host ''
    Write-Host '--- Select FROM group (current membership) ---' -ForegroundColor Cyan
    if ($groupSource) {
        $fromGroup = $groupSource
        Write-Host (" FROM group  : {0} (the source group you picked)" -f $fromGroup.DisplayName) -ForegroundColor Cyan
    } else {
        $fromGroup = Select-GroupInteractive -Prompt 'From (source) group' -Provider $provider -Limit $GroupLimit -Groups $allGroups
        Write-Host (" FROM group  : {0}" -f $fromGroup.DisplayName) -ForegroundColor Cyan
    }
}

Write-Host ''
Write-Host '--- Select TO group (target) ---' -ForegroundColor Cyan
$toGroup = Select-GroupInteractive -Prompt 'To (target) group' -Provider $provider -Limit $GroupLimit -Groups $allGroups
Write-Host (" TO group    : {0}" -f $toGroup.DisplayName) -ForegroundColor Cyan
# Pre-fetch group members
$toId = Get-GroupId -Group $toGroup -Provider $provider
if ($isMove) {
    $fromId = Get-GroupId -Group $fromGroup -Provider $provider
    Write-Host "Fetching members of FROM group..." -ForegroundColor DarkGray
    $fromMembers = @(Get-GroupMemberIds -GroupId $fromId -Provider $provider)
}
Write-Host "Fetching members of TO group..." -ForegroundColor DarkGray
$toMembers = @(Get-GroupMemberIds -GroupId $toId -Provider $provider)

# Process
$fromGroupName = if ($fromGroup) { $fromGroup.DisplayName } else { '(none)' }
$stats = [ordered]@{ Resolved=0; Moved=0; AddedOnly=0; Skipped=0; NotFound=0; Failed=0; WhatIf=0 }

foreach ($item in $userList) {
    $value = $item.Identifier
    $userObj = $item.UserObj
    $users = @()
    if ($userObj) {
        $users = @($userObj)
    } else {
        $users = @(Resolve-EntraUser -Value $value -Type $IdentifierType -Provider $provider)
    }
    if ($users.Count -eq 0) { $stats.NotFound++; Write-Host (" NOT FOUND  : {0}" -f $value) -ForegroundColor DarkYellow; Write-Log -Action 'Resolve' -Identifier $value -DisplayName '' -UserId '' -FromGroup $fromGroupName -ToGroup $toGroup.DisplayName -Result 'NotFound' -Message 'No matching Entra user'; continue }
    if ($users.Count -gt 1) { Write-Warning (" Ambiguous: '{0}' matched {1} users. Skipping." -f $value, $users.Count); continue }

    $user = $users[0]
    $oid = Get-UserId -User $user -Provider $provider
    $disp = Get-UserDisplay -User $user
    # Never leave the audit line blank: fall back to the identifier, then the id.
    if ([string]::IsNullOrWhiteSpace($disp)) {
        $disp = if (-not [string]::IsNullOrWhiteSpace($value)) { $value } else { [string]$oid }
    }
    $stats.Resolved++

    # Check group membership (cached lists, updated as we go)
    $inFrom = $isMove -and ($fromMembers -contains $oid)
    $inTo = $toMembers -contains $oid

    $actionDesc = if ($isMove) { "Move from '$($fromGroup.DisplayName)' to '$($toGroup.DisplayName)'" } else { "Add to '$($toGroup.DisplayName)'" }
    if ($PSCmdlet.ShouldProcess($disp, $actionDesc)) {
        try {
            if ($isMove -and $inFrom) {
                Remove-GroupMember -GroupId $fromId -MemberId $oid -Provider $provider
                $fromMembers = @($fromMembers | Where-Object { $_ -ne $oid })
            } elseif ($isMove -and -not $Force -and -not $inFrom) {
                $stats.Skipped++
                Write-Host (" SKIPPED   : {0} (not in FROM group; use -Force to add anyway)" -f $disp) -ForegroundColor Yellow
                Write-Log -Action 'Move' -Identifier $value -DisplayName $disp -UserId $oid -FromGroup $fromGroupName -ToGroup $toGroup.DisplayName -Result 'SkippedNotInSource' -Message ''
                continue
            }
            if (-not $inTo) {
                Add-GroupMember -GroupId $toId -MemberId $oid -Provider $provider
                $toMembers += $oid
                $stats.Moved++
                $msg = if ($isMove) { 'MOVED' } else { 'ADDED' }
                Write-Host (" {0,-10}: {1}" -f $msg, $disp) -ForegroundColor Green
            } else {
                $stats.AddedOnly++
                Write-Host (" ALREADY   : {0} (already in TO group)" -f $disp) -ForegroundColor DarkGray
            }
            Write-Log -Action $(if ($isMove) { 'Move' } else { 'Add' }) -Identifier $value -DisplayName $disp -UserId $oid -FromGroup $(if ($fromGroup) { $fromGroup.DisplayName } else { '(none)' }) -ToGroup $toGroup.DisplayName -Result 'Success' -Message ''
        } catch {
            $stats.Failed++
            Write-Host (" FAILED    : {0} -> {1}" -f $disp, $_.Exception.Message) -ForegroundColor Red
            Write-Log -Action 'Move' -Identifier $value -DisplayName $disp -UserId $oid -FromGroup $fromGroupName -ToGroup $toGroup.DisplayName -Result 'Failed' -Message $_.Exception.Message
        }
    } else {
        $stats.WhatIf++
        Write-Host (" WHATIF    : would move {0}" -f $disp) -ForegroundColor DarkCyan
        Write-Log -Action 'Move' -Identifier $value -DisplayName $disp -UserId $oid -FromGroup $fromGroupName -ToGroup $toGroup.DisplayName -Result 'WhatIf' -Message 'Not executed'
    }
}

Write-Host ''
Write-Host '-------------------------------------------'
Write-Host 'Results:' -ForegroundColor Cyan
Write-Host (" Resolved : {0}" -f $stats.Resolved)
Write-Host (" Moved    : {0}" -f $stats.Moved)
Write-Host (" AddedOnly: {0}" -f $stats.AddedOnly)
Write-Host (" Skipped  : {0}" -f $stats.Skipped)
Write-Host (" NotFound : {0}" -f $stats.NotFound)
Write-Host (" Failed   : {0}" -f $stats.Failed)
Write-Host (" WhatIf   : {0}" -f $stats.WhatIf)
Write-Host (" Audit log: {0}" -f $LogPath)
if ($WhatIfPreference) { Write-Host 'WhatIf mode active - nothing changed.' -ForegroundColor Cyan }
