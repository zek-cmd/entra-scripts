#Requires -Version 5.1
<#
.SYNOPSIS
    Remove licenses from Entra users — by Excel file or by group membership.
.DESCRIPTION
    Reads users from an .xlsx file OR selects an interactive group, then removes
    either ALL licenses or a specific SKU (e.g., M365_E3) from each user.

    Supports both Microsoft.Graph and legacy AzureAD modules.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)][string]$ExcelPath,
    [Parameter(Mandatory = $false)][string]$ColumnName,
    [ValidateSet('DisplayName', 'UPN', 'Mail', 'ObjectId')][string]$IdentifierType = 'DisplayName',
    [string]$NameSeparator = ' ',
    [switch]$Force,
    [int]$GroupLimit = 500,
    [switch]$RemoveAll,
    [string]$SkuId,
    [string]$LogPath
)

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
    $LogPath = Join-Path (Join-Path $scriptRoot 'Output Report') ("LicenseRemoval_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
$script:LogFile = $LogPath

# ---------------------------------------------------------------------------
# Audit log (.NET IO so -WhatIf does not suppress it)
# ---------------------------------------------------------------------------
function Write-Log {
    param([string]$Action, [string]$Identifier, [string]$DisplayName, [string]$UserId,
          [string]$Sku, [string]$Result, [string]$Message)
    $line = @($Action, $Identifier, $DisplayName, $UserId, $Sku, $Result,
        ($Message -replace "`r|`n", ' '), (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) -join '","'
    $line = '"' + $line + '"'
    $enc = New-Object System.Text.UTF8Encoding($true)
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
    $d = New-Object System.Windows.Forms.OpenFileDialog
    $d.Title = 'Select Excel file with user list'
    $d.Filter = 'Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    $d.FilterIndex = 1
    $d.InitialDirectory = if (Test-Path -LiteralPath $scriptRoot) { $scriptRoot } else { $env:USERPROFILE }
    if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $d.FileName }
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

    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        $data = @(Import-Excel -Path $File -ErrorAction Stop)
        if ($data.Count -gt 0) { return @($data[0].PSObject.Properties.Name) }
    }
    return @()
}

function Read-ExcelColumn {
    <# Python + openpyxl first, ImportExcel as the fallback. #>
    param([string]$File, [string[]]$Column, [string]$Separator = ' ')
    $cols = @($Column)
    if ($cols.Count -eq 0) { return @() }

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

    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        $data = Import-Excel -Path $File -ErrorAction Stop
        return @($data | ForEach-Object {
            $parts = foreach ($c in $cols) {
                $v = $_.$c
                if ($null -ne $v -and ([string]$v).Trim() -ne '') { ([string]$v).Trim() }
            }
            $j = ($parts -join $Separator)
            if ($j -ne '') { $j }
        })
    }

    throw "Cannot read Excel. Install one of: python -m pip install openpyxl / Install-Module ImportExcel -Scope CurrentUser"
}

function Get-EntraGroups {
    param([string]$Provider, [int]$Limit)
    if ($Provider -eq 'Graph') { $g = @(Get-MgGroup -All -ErrorAction SilentlyContinue | Select-Object Id, DisplayName, GroupTypes, Mail) }
    else { $g = @(Get-AzureADGroup -ErrorAction SilentlyContinue | Select-Object ObjectId, DisplayName, Mail) }
    if ($Limit -gt 0 -and $g.Count -gt $Limit) { $g = $g | Select-Object -First $Limit }
    return $g
}

function Select-GroupInteractive {
    param([string]$Prompt, [string]$Provider, [int]$Limit, $Groups)
    Write-Host ("=== {0} ===" -f $Prompt) -ForegroundColor Cyan
    Write-Host ' Groups (numbered):' -ForegroundColor DarkGray
    $i = 0; foreach ($g in $Groups) { $i++; $type = ''; if ($g.PSObject.Properties['GroupTypes'] -and $g.GroupTypes) { if ($g.GroupTypes -contains 'Unified') { $type = ' [M365]' } elseif ($g.GroupTypes -contains 'DynamicMembership') { $type = ' [Dynamic]' } else { $type = ' [Sec]' } } Write-Host ("  [{0}] {1}{2}" -f $i, $g.DisplayName, $type) -ForegroundColor Gray }
    while ($true) {
        $choice = (Read-Host "Select group number or type part of name").Trim()
        if ($choice -match '^\d+$') { $idx = [int]$choice; if ($idx -ge 1 -and $idx -le $Groups.Count) { return $Groups[$idx - 1] } }
        $matches = @($Groups | Where-Object { $_.DisplayName -like "*$choice*" })
        if ($matches.Count -eq 1) { return $matches[0] }
        if ($matches.Count -gt 1) { Write-Host ("  {0} matches:" -f $matches.Count) -ForegroundColor Yellow; $j = 0; foreach ($m in $matches) { $j++; Write-Host ("    [{0}] {1}" -f $j, $m.DisplayName) -ForegroundColor Gray } $pick = Read-Host '  Enter number'; if ($pick -match '^\d+$' -and [int]$pick -ge 1 -and [int]$pick -le $matches.Count) { return $matches[[int]$pick - 1] } }
        else { Write-Host ("  No match for '{0}'." -f $choice) -ForegroundColor Red }
    }
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

function Get-GraphProvider {
    if (Get-Module -ListAvailable -Name Microsoft.Graph.Users) { return 'Graph' }
    if (Get-Module -ListAvailable -Name Microsoft.Graph) { return 'Graph' }
    if (Get-Module -ListAvailable -Name AzureAD) { return 'AzureAD' }
    return $null
}

# ---------------------------------------------------------------------------
# License functions
# ---------------------------------------------------------------------------
function Get-UserLicenses {
    param([string]$UserId, [string]$Provider)
    if ($Provider -eq 'Graph') {
        return @(Get-MgUserLicenseDetail -UserId $UserId -ErrorAction SilentlyContinue)
    } else {
        return @(Get-AzureADUserLicenseDetail -ObjectId $UserId -ErrorAction SilentlyContinue)
    }
}

function Get-UserSkuIds {
    param([string]$UserId, [string]$Provider)
    return @(Get-UserLicenses -UserId $UserId -Provider $Provider | ForEach-Object { $_.SkuPartNumber })
}

function Remove-UserSku {
    <# Removes license(s) by LicenseDetailsId (Graph) or SkuId (AzureAD). #>
    param([string]$UserId, [string]$SkuId, [string]$Provider)
    $all = @(Get-UserLicenses -UserId $UserId -Provider $Provider)
    if ($Provider -eq 'Graph') {
        foreach ($l in $all) {
            if ([string]::IsNullOrWhiteSpace($SkuId) -or $SkuId -eq '*' -or $l.SkuPartNumber -eq $SkuId) {
                if ($l.Id) { Remove-MgUserLicenseDetail -UserId $UserId -LicenseDetailsId $l.Id -ErrorAction SilentlyContinue | Out-Null }
            }
        }
    } else {
        foreach ($l in $all) {
            if ([string]::IsNullOrWhiteSpace($SkuId) -or $SkuId -eq '*' -or $l.SkuPartNumber -eq $SkuId) {
                if ($l.SkuPartNumber) { Remove-AzureADUserLicense -ObjectId $UserId -SkuId $l.SkuPartNumber -ErrorAction SilentlyContinue | Out-Null }
            }
        }
    }
}

function Get-SkuNameMap {
    <# Builds a lookup: SkuPartNumber -> friendly license name (e.g. ENTERPRISEPACK -> "Office 365 (Enterprise)"). #>
    param([string]$Provider)
    $map = @{}
    $raw = @()
    if ($Provider -eq 'Graph') {
        if (-not (Get-Command Get-MgSubscribedSku -ErrorAction SilentlyContinue)) { return $map }
        $raw = @(Get-MgSubscribedSku -ErrorAction SilentlyContinue)
    } else {
        if (-not (Get-Command Get-AzureADSubscribedSku -ErrorAction SilentlyContinue)) { return $map }
        $raw = @(Get-AzureADSubscribedSku -ErrorAction SilentlyContinue)
    }
    foreach ($s in $raw) {
        $part = $s.SkuPartNumber
        if (-not $part) { continue }
        $name = ''
        $plans = $null
        if ($s.PSObject.Properties['ServicePlans']) { $plans = $s.ServicePlans }
        if ($plans) {
            try { $first = $plans[0]; if ($first -and $first.ServicePlanName) { $name = $first.ServicePlanName } } catch { }
            if (-not $name) {
                try { if ($plans.ServicePlanName) { $name = $plans.ServicePlanName } } catch { }
            }
        }
        if (-not $name) { $name = $part }
        $map[$part] = $name
    }
    return $map
}

function Get-SkuLabel {
    <# Display label for a SKU part number. #>
    param([string]$PartNumber, $Map)
    if (-not $PartNumber) { return '' }
    if ($PartNumber -eq '*') { return 'ALL licenses' }
    return $PartNumber
}

function Get-SkuPicker {
    param([string]$UserId, [string]$Provider)
    $lic = @(Get-UserLicenses -UserId $UserId -Provider $Provider)
    if ($lic.Count -eq 0) { return '*' }
    Write-Host '  Assigned licenses:' -ForegroundColor DarkGray
    $i = 0
    foreach ($l in $lic) {
        $i++
        $friendly = Get-SkuLabel -PartNumber $l.SkuPartNumber -Map $script:SkuNameMap
        Write-Host ("    [{0}] {1}" -f $i, $friendly) -ForegroundColor Gray
    }
    Write-Host '    [A] ALL assigned licenses' -ForegroundColor Gray
    $choice = (Read-Host '    Select number, A (all), or type part of the name').Trim()
    if ($choice -eq 'A' -or $choice -eq 'a') { return '*' }
    if ($choice -match '^\d+$') {
        $idx = [int]$choice
        if ($idx -ge 1 -and $idx -le $lic.Count) { return $lic[$idx - 1].SkuPartNumber }
    }
    foreach ($l in $lic) { if ($l.SkuPartNumber -like "*$choice*") { return $l.SkuPartNumber } }
    foreach ($l in $lic) {
        $friendly = Get-SkuLabel -PartNumber $l.SkuPartNumber -Map $script:SkuNameMap
        if ($friendly -like "*$choice*") { return $l.SkuPartNumber }
    }
    return $choice
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$provider = Get-GraphProvider
if (-not $provider) { throw "No Entra module found. Install Microsoft.Graph or AzureAD." }
if ($provider -eq 'Graph') {
    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        Connect-MgGraph -Scopes 'User.Read.All','User.ReadWrite.All','Directory.Read.All','Organization.Read.All'
    } else { Write-Host (" Connected as: {0}" -f (Get-MgContext).Account) -ForegroundColor DarkGray }
} else {
    if (-not (Get-AzureADTenant -ErrorAction SilentlyContinue)) { Connect-AzureAD }
}

# --- Source mode: Excel or Group ---
Write-Host ''
Write-Host '=== Select source ===' -ForegroundColor Cyan
Write-Host '  [1] Excel file - read users from a spreadsheet' -ForegroundColor Gray
Write-Host '  [2] Group     - use all members of a selected group' -ForegroundColor Gray
$srcMode = ''
while ($srcMode -notmatch '^[12]$') {
    $srcMode = (Read-Host 'Enter 1 or 2').Trim()
    if ($srcMode -notmatch '^[12]$') { Write-Host '  Please enter 1 or 2.' -ForegroundColor Red }
}
$useExcel = ($srcMode -eq '1')

$targets = @()   # list of @{ Id=...; Display=...; Identifier=... }
$sourceLabel = ''

if ($useExcel) {
    if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
        $ExcelPath = Show-OpenExcelDialog
        if ([string]::IsNullOrWhiteSpace($ExcelPath)) { Write-Host 'No file selected.' -ForegroundColor Red; exit 1 }
    } elseif (-not (Test-Path -LiteralPath $ExcelPath)) {
        Write-Host ("File not found: {0}" -f $ExcelPath) -ForegroundColor Red
        $ExcelPath = Show-OpenExcelDialog
        if ([string]::IsNullOrWhiteSpace($ExcelPath)) { Write-Host 'No file selected.' -ForegroundColor Red; exit 1 }
    }
    $ExcelPath = (Resolve-Path -LiteralPath $ExcelPath).Path
    Write-Host (" Excel file  : {0}" -f $ExcelPath) -ForegroundColor DarkGray

    # Column selection
    if ([string]::IsNullOrWhiteSpace($ColumnName)) {
        $headers = @(Get-ExcelColumnHeaders -File $ExcelPath)
        if ($headers.Count -eq 0) { throw "No columns detected." }
        Write-Host ' Detected columns:' -ForegroundColor DarkGray
        for ($i = 0; $i -lt $headers.Count; $i++) { Write-Host ("   [{0}] {1}" -f $i, $headers[$i]) -ForegroundColor DarkGray }
        $choice = Read-Host ' Select index, name, or combo (e.g. FIRST NAME + LASTNAME)'
        $choice = $choice.Trim().Trim('"').Trim("'")
        if ($choice -match '^\d+$') { $ColumnName = [string]$headers[[int]$choice] } else { $ColumnName = [string]$choice }
    }
    if (-not $PSBoundParameters.ContainsKey('IdentifierType')) {
        $cl = $ColumnName.ToLower()
        if ($cl -match 'fullname|name' -and -not ($cl -match 'id')) { $IdentifierType = 'DisplayName' }
        elseif ($cl -match 'mail|email') { $IdentifierType = 'Mail' }
        elseif ($cl -match 'id' -and -not ($cl -match 'name')) { $IdentifierType = 'ObjectId' }
        else { $IdentifierType = 'DisplayName' }
        Write-Host (" Identifier   : {0}" -f $IdentifierType) -ForegroundColor DarkGray
    }
    $columns = @($ColumnName)
    if ($ColumnName -match ' \+ ') {
        $columns = $ColumnName.Split(@(' + '), [System.StringSplitOptions]::None) | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        Write-Host (" Composite    : {0}" -f ($columns -join ' | ')) -ForegroundColor DarkGray
    }
    $values = @(Read-ExcelColumn -File $ExcelPath -Column $columns -Separator $NameSeparator |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($values.Count -eq 0) { Write-Warning 'No users found.'; exit 0 }
    $sourceLabel = "Excel: $(Split-Path -Leaf $ExcelPath) / $ColumnName"
    $targets = @($values | ForEach-Object { @{ Identifier = $_; Id = $null; Display = $_ } })
} else {
    $allGroups = @(Get-EntraGroups -Provider $provider -Limit $GroupLimit)
    if ($allGroups.Count -eq 0) { throw "No groups found." }
    Write-Host ''
    $grp = Select-GroupInteractive -Prompt 'Select group to process' -Provider $provider -Limit $GroupLimit -Groups $allGroups
    $grpId = if ($provider -eq 'Graph') { $grp.Id } else { $grp.ObjectId }
    Write-Host ''
    Write-Host " Fetching members of '$($grp.DisplayName)'..." -ForegroundColor DarkGray
    # Typed member fetch: Get-MgGroupMember returns IMicrosoftGraphDirectoryObject,
    # which has only Id - no DisplayName/UserPrincipalName. Get-MgGroupMemberAsUser
    # returns full User objects and excludes nested groups / devices / service
    # principals server-side. Falls back to per-id resolution on older modules.
    $members = @()
    if ($provider -eq 'Graph') {
        if (Get-Command Get-MgGroupMemberAsUser -ErrorAction SilentlyContinue) {
            $members = @(Get-MgGroupMemberAsUser -GroupId $grpId -All -ErrorAction SilentlyContinue)
        } else {
            foreach ($m in @(Get-MgGroupMember -GroupId $grpId -All -ErrorAction SilentlyContinue)) {
                if (-not $m.Id) { continue }
                $u = $null
                try { $u = Get-MgUser -UserId $m.Id -ErrorAction SilentlyContinue } catch { $u = $null }
                if ($u) { $members += $u }
            }
        }
    } else {
        $members = @(Get-AzureADGroupMember -ObjectId $grpId -ErrorAction SilentlyContinue |
            Where-Object { -not $_.ObjectType -or [string]$_.ObjectType -eq 'User' })
    }
    $targets = @($members | ForEach-Object {
        $mid = if ($_.Id) { $_.Id } else { $_.ObjectId }
        $d = if ($_.UserPrincipalName) { $_.UserPrincipalName }
             elseif ($_.DisplayName)   { $_.DisplayName }
             elseif ($_.Mail)          { $_.Mail }
             else                      { $mid }
        @{ Identifier = $d; Id = $mid; Display = $d }
    })
    if ($targets.Count -eq 0) { Write-Warning 'Group has no user members.'; exit 0 }
    $sourceLabel = "Group: $($grp.DisplayName)"
}

Write-Host ''
Write-Host (" Users to process: {0}" -f $targets.Count) -ForegroundColor Cyan
Write-Host (" Source           : {0}" -f $sourceLabel) -ForegroundColor Cyan

# --- License scope selection ---
Write-Host ''
Write-Host '=== Select license scope ===' -ForegroundColor Cyan
Write-Host '  [1] ALL      - remove every license from each user' -ForegroundColor Gray
Write-Host '  [2] SPECIFIC - choose one SKU (per user, from their assigned licenses)' -ForegroundColor Gray
Write-Host '  [3] FIXED    - remove one specific SKU id from every user' -ForegroundColor Gray
$scope = ''
while ($scope -notmatch '^[123]$') {
    $scope = (Read-Host 'Enter 1, 2 or 3').Trim()
    if ($scope -notmatch '^[123]$') { Write-Host '  Please enter 1, 2 or 3.' -ForegroundColor Red }
}
if ($scope -eq '3') {
    # Auto-scan assigned SKUs from all target users (resolve once, cache on target)
    $allSkus = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($t in $targets) {
        if (-not $t.Id) {
            $tmpUser = @(Resolve-EntraUser -Value $t.Identifier -Type $IdentifierType -Provider $provider)
            if ($tmpUser.Count -gt 0) {
                $t.Id = Get-UserId -User $tmpUser[0] -Provider $provider
                $t.Display = Get-UserDisplay -User $tmpUser[0]
            }
        }
        if ($t.Id) {
            $tmpSkus = @(Get-UserSkuIds -UserId $t.Id -Provider $provider)
            $t.Skus = $tmpSkus
            foreach ($s in $tmpSkus) { [void]$allSkus.Add($s) }
        }
    }
    $skus = @($allSkus | Sort-Object)
    if ($skus.Count -eq 0) { Write-Warning 'No licenses found across users.'; exit 1 }
    Write-Host ' Found SKUs in dataset:' -ForegroundColor Cyan
    $i = 0; foreach ($s in $skus) { $i++; Write-Host ("   [{0}] {1}" -f $i, (Get-SkuLabel -PartNumber $s -Map $script:SkuNameMap)) -ForegroundColor Gray }
    $pick = (Read-Host ' Select SKU number (or type part of SKU id)').Trim()
    if ($pick -match '^\d+$' -and [int]$pick -ge 1 -and [int]$pick -le $skus.Count) {
        $SkuId = $skus[[int]$pick - 1]
    } else {
        $match = $skus | Where-Object { $_ -like "*$pick*" } | Select-Object -First 1
        if ($match) { $SkuId = $match } else { $SkuId = $pick }
    }
    Write-Host (" Fixed SKU set to: {0}" -f $SkuId) -ForegroundColor Cyan
}

$stats = [ordered]@{ Processed=0; Removed=0; NoLicense=0; NotFound=0; Failed=0; WhatIf=0 }

foreach ($t in $targets) {
    # Resolve user object (reuse cached Id when available)
    $userObj = @()
    if ($t.Id) {
        $userObj = @(Get-MgUser -UserId $t.Id -ErrorAction SilentlyContinue)
        if ($userObj.Count -eq 0) { $userObj = @(Get-AzureADUser -ObjectId $t.Id -ErrorAction SilentlyContinue) }
    } else {
        $userObj = @(Resolve-EntraUser -Value $t.Identifier -Type $IdentifierType -Provider $provider)
    }
    $u = if ($userObj.Count -gt 0) { $userObj[0] } else { $null }
    $oid = $null; $disp = ''
    if ($u) {
        $oid = Get-UserId -User $u -Provider $provider
        $disp = Get-UserDisplay -User $u
        $t.Id = $oid; $t.Display = $disp
    } elseif ($t.Id) {
        $oid = $t.Id
        $disp = if ($t.Display) { $t.Display } else { $t.Identifier }
    } else {
        $stats.NotFound++
        Write-Host (" NOT FOUND : {0}" -f $t.Identifier) -ForegroundColor DarkYellow
        Write-Log -Action 'Resolve' -Identifier $t.Identifier -DisplayName '' -UserId '' -Sku '' -Result 'NotFound' -Message 'No matching Entra user'
        continue
    }
    $stats.Processed++
    $skus = if ($t.Skus) { $t.Skus } else { @(Get-UserSkuIds -UserId $oid -Provider $provider) }
    if ($skus.Count -eq 0) {
        $stats.NoLicense++
        Write-Host (" NO LICENSE: {0}" -f $disp) -ForegroundColor DarkGray
        Write-Log -Action 'RemoveLicense' -Identifier $t.Identifier -DisplayName $disp -UserId $oid -Sku '' -Result 'NoLicense' -Message ''
        continue
    }

    # Determine SKU(s) to remove
    $targetSku = '*'
    if ($scope -eq '2') { $targetSku = Get-SkuPicker -UserId $oid -Provider $provider }
    elseif ($scope -eq '3') { $targetSku = $SkuId }

    $label = if ($targetSku -eq '*') { 'ALL licenses' } else { Get-SkuLabel -PartNumber $targetSku -Map $script:SkuNameMap }
    if ($PSCmdlet.ShouldProcess($disp, "Remove $label")) {
        try {
            Remove-UserSku -UserId $oid -SkuId $targetSku -Provider $provider
            $stats.Removed++
            Write-Host (" REMOVED   : {0} ({1})" -f $disp, $label) -ForegroundColor Green
            Write-Log -Action 'RemoveLicense' -Identifier $t.Identifier -DisplayName $disp -UserId $oid -Sku $targetSku -Result 'Success' -Message ''
        } catch {
            $stats.Failed++
            Write-Host (" FAILED    : {0} -> {1}" -f $disp, $_.Exception.Message) -ForegroundColor Red
            Write-Log -Action 'RemoveLicense' -Identifier $t.Identifier -DisplayName $disp -UserId $oid -Sku $targetSku -Result 'Failed' -Message $_.Exception.Message
        }
    } else {
        $stats.WhatIf++
        Write-Host (" WHATIF    : would remove {0} from {1}" -f $label, $disp) -ForegroundColor DarkCyan
        Write-Log -Action 'RemoveLicense' -Identifier $t.Identifier -DisplayName $disp -UserId $oid -Sku $targetSku -Result 'WhatIf' -Message 'Not executed'
    }
}

Write-Host ''
Write-Host '-------------------------------------------'
Write-Host 'Results:' -ForegroundColor Cyan
Write-Host (" Processed : {0}" -f $stats.Processed)
Write-Host (" Removed   : {0}" -f $stats.Removed)
Write-Host (" NoLicense : {0}" -f $stats.NoLicense)
Write-Host (" NotFound  : {0}" -f $stats.NotFound)
Write-Host (" Failed    : {0}" -f $stats.Failed)
Write-Host (" WhatIf    : {0}" -f $stats.WhatIf)
Write-Host (" Audit log : {0}" -f $LogPath)
if ($WhatIfPreference) { Write-Host 'WhatIf mode active - nothing was changed.' -ForegroundColor Cyan }
