<#
.SYNOPSIS
  Bulk-create M365 accounts from SOFTNET.xlsx using Microsoft Graph.
.DESCRIPTION
  Reads first/last names, builds usernames (first word of first name only,
  lowercase; last name lowercase + accents stripped), creates users with
  P@ssw0rd + force-change, prompts interactively for groups and licenses,
  skips existing accounts, and writes an account-creation CSV report.
.NOTES
  Prerequisites: Microsoft.Graph module + Connect-MgGraph with
  User.ReadWrite.All, Group.ReadWrite.All, Directory.Read.All,
  UserManagementPolicy.ReadWrite.All, Organization.Read.All scopes.
  Set $Domain to your tenant (e.g., softnetitc.edu.ph) before running.

  UsageLocation defaults to PH (Philippines, ISO 3166-1 alpha-2). Licenses
  cannot be assigned until UsageLocation is set to a country the tenant is
  actually licensed for, so override with -UsageLocation if needed.
#>
param(
    [string]$XlsxPath       = "",
    [string]$CsvReport     = "",
    [string]$Domain        = "",

    # ISO 3166-1 alpha-2 country code. PH = Philippines.
    [string]$UsageLocation = "PH",

    # Default password for new accounts; every account must change it at first sign-in.
    [string]$DefaultPassword = "P@ssw0rd",

    [switch]$WhatIf,
    [switch]$SkipExisting  = $true
)

# Report path: resolve next to the script so the run does not depend on the
# caller's current directory (menu.ps1 runs from the script folder either way).
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($CsvReport)) {
    $CsvReport = Join-Path (Join-Path $scriptDir 'Output Report') 'Account-Creation-Report.csv'
}

if ($UsageLocation -notmatch '^[A-Za-z]{2}$') {
    Write-Error "UsageLocation must be a 2-letter ISO country code (e.g. PH). Got: '$UsageLocation'"
    exit 1
}
$UsageLocation = $UsageLocation.ToUpper()

# --- DOMAIN INPUT ---
if ([string]::IsNullOrWhiteSpace($Domain)) {
    $Domain = Read-Host "Enter M365 domain (e.g. contoso.onmicrosoft.com)"
}
if ([string]::IsNullOrWhiteSpace($Domain) -or $Domain -match "YOURDOMAIN") {
    Write-Error "Domain is required. Pass -Domain or enter it at the prompt. Example: contoso.onmicrosoft.com"; exit 1
}

function Select-ExcelFile {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Select Excel File"
    $dialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
    $dialog.InitialDirectory = (Get-Location).Path
    $dialog.Multiselect = $false
    $dialog.CheckFileExists = $true
    $dialog.CheckPathExists = $true
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    return $null
}

# --- EXCEL FILE PICKER ---
if ([string]::IsNullOrWhiteSpace($XlsxPath)) {
    $XlsxPath = Select-ExcelFile
    if ([string]::IsNullOrWhiteSpace($XlsxPath)) { Write-Error "No file selected."; exit 1 }
}
if (-not (Test-Path $XlsxPath)) { Write-Error "File not found: $XlsxPath"; exit 1 }

# --- EXCEL READER SETUP ---
# Primary path: Python + openpyxl, invoked with -c so NO external .py file is
# written to %TEMP%. (The old approach wrote a script to TEMP then ran it, which
# broke under -WhatIf: Set-Content was suppressed, so Python was handed a path
# that did not exist -> "can't open file ...read_softnet_for_ps.py".)
# Fallback: the ImportExcel module, if Python/openpyxl is unavailable.
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:PyResolverPath = Join-Path $scriptRoot 'Resolve-PythonWithOpenpyxl.ps1'
if (Test-Path -LiteralPath $script:PyResolverPath) { . $script:PyResolverPath }

function Get-PythonForOpenpyxl {
    <# Returns an interpreter that can import openpyxl, or $null. #>
    if (Get-Command Resolve-PythonWithOpenpyxl -ErrorAction SilentlyContinue) {
        $found = Resolve-PythonWithOpenpyxl
        if ($found) { return $found }
    }
    $candidate = Join-Path $env:USERPROFILE '.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe'
    if (Test-Path -LiteralPath $candidate) {
        & $candidate -c "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('openpyxl') else 1)" 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { return $candidate }
    }
    return $null
}

function Split-FullNameParts {
    <# Mirror of the Python splitter:
       'JUAN FLAVIO C. DELA CRUZ' -> first=JUAN, mname=FLAVIO C., lname=DELA CRUZ
       (last two tokens are the surname; anything between first and surname is middle) #>
    param([string]$Raw)
    $parts = @((([string]$Raw) -replace ',', ' ') -split '\s+' | Where-Object { $_ })
    if ($parts.Count -eq 0) { return [pscustomobject]@{ fname = ''; mname = ''; lname = '' } }
    if ($parts.Count -eq 1) { return [pscustomobject]@{ fname = $parts[0]; mname = ''; lname = '' } }
    if ($parts.Count -eq 2) { return [pscustomobject]@{ fname = $parts[0]; mname = ''; lname = $parts[1] } }
    $mname = if ($parts.Count -ge 4) { ($parts[1..($parts.Count - 3)] -join ' ') } else { '' }
    $lname = ($parts[($parts.Count - 2)..($parts.Count - 1)] -join ' ')
    return [pscustomobject]@{ fname = $parts[0]; mname = $mname; lname = $lname }
}

function Get-HeaderColumnMap {
    <# Maps header names to logical columns using the same rules as the Python
       reader (case-insensitive; non [a-z. ] characters stripped). #>
    param([string[]]$Headers)
    $map = @{}
    $fullCol = $null
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $h = ([string]$Headers[$i]).ToLower() -replace '[^a-z. ]', ''
        $h = $h.Trim()
        if (@('full name', 'fullname', 'complete name', 'completename', 'name') -contains $h) {
            if ($null -eq $fullCol) { $fullCol = $i }
        }
        elseif ($h -match 'id')   { if (-not $map.ContainsKey('id'))   { $map['id'] = $i } }
        elseif ($h -match 'first' -or $h -eq 'fname') { if (-not $map.ContainsKey('fname')) { $map['fname'] = $i } }
        elseif (($h -match 'm\.i\.') -or (@('mi', 'm.i', 'middle name', 'middlename', 'middle initial') -contains $h)) {
            if (-not $map.ContainsKey('mname')) { $map['mname'] = $i }
        }
        elseif ($h -match 'last' -or $h -eq 'lname') { if (-not $map.ContainsKey('lname')) { $map['lname'] = $i } }
        elseif ($h -match 'position') { if (-not $map.ContainsKey('position')) { $map['position'] = $i } }
        elseif ($h -match 'class')    { if (-not $map.ContainsKey('class'))    { $map['class'] = $i } }
    }
    return [pscustomobject]@{ Map = $map; FullCol = $fullCol }
}

function Read-UsersWithImportExcel {
    <# Fallback reader. Returns rows shaped like the Python path. #>
    param([string]$Path)
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) { return $null }
    Import-Module ImportExcel -ErrorAction Stop
    $data = @(Import-Excel -Path $Path -ErrorAction Stop)
    if ($data.Count -eq 0) { return @() }

    $headers = @($data[0].PSObject.Properties.Name)
    $info = Get-HeaderColumnMap -Headers $headers
    $rows = @()
    foreach ($d in $data) {
        $vals = @($headers | ForEach-Object { $d.$_ })

        $getVal = {
            param($key, $defaultIdx)
            if ($info.Map.ContainsKey($key)) {
                $idx = $info.Map[$key]
                if ($idx -lt $vals.Count) { return $vals[$idx] }
                return $null
            }
            if ($defaultIdx -lt $vals.Count) { return $vals[$defaultIdx] }
            return $null
        }

        $srcFull = $null
        if ($null -ne $info.FullCol -and $info.FullCol -lt $vals.Count) { $srcFull = $vals[$info.FullCol] }

        if ($null -ne $srcFull -and -not [string]::IsNullOrWhiteSpace([string]$srcFull)) {
            $sp = Split-FullNameParts -Raw ([string]$srcFull)
            $f = $sp.fname; $m = $sp.mname; $l = $sp.lname
        } else {
            $f = & $getVal 'fname' 1
            $m = & $getVal 'mname' 2
            $l = & $getVal 'lname' 3
        }

        $rows += [pscustomobject]@{
            id       = & $getVal 'id' 0
            fname    = $f
            mname    = $m
            lname    = $l
            position = & $getVal 'position' 4
            class    = & $getVal 'class' 5
            fullname = $srcFull
        }
    }
    return $rows
}

# --- READ EXCEL ---
# Python + openpyxl first (inline, no temp .py file), ImportExcel as fallback.
Write-Host "Reading $XlsxPath ..." -ForegroundColor Cyan

$readPyCode = @'
import openpyxl, sys, json, re

wb = openpyxl.load_workbook(sys.argv[1], data_only=False)
ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb[wb.sheetnames[0]]

header_map = {}
full_col = None
headers = [str(c.value or '').strip().lower() for c in ws[1]]
for idx, h in enumerate(headers):
    h = re.sub(r'[^a-z. ]', '', h).strip()
    if h in ('full name', 'fullname', 'complete name', 'completename', 'name'):
        if full_col is None:
            full_col = idx
    elif 'id' in h:
        header_map.setdefault('id', idx)
    elif 'first' in h or h == 'fname':
        header_map.setdefault('fname', idx)
    elif 'm.i.' in h or h in ('mi', 'm.i', 'middle name', 'middlename', 'middle initial'):
        header_map.setdefault('mname', idx)
    elif 'last' in h or h == 'lname':
        header_map.setdefault('lname', idx)
    elif 'position' in h:
        header_map.setdefault('position', idx)
    elif 'class' in h:
        header_map.setdefault('class', idx)

def split_full_name(raw):
    # 'JUAN FLAVIO C. DELA CRUZ' -> first=JUAN, mname=FLAVIO C., lname=DELA CRUZ
    parts = [p for p in str(raw or '').replace(',', ' ').split() if p]
    if not parts:
        return '', '', ''
    if len(parts) == 1:
        return parts[0], '', ''
    if len(parts) == 2:
        return parts[0], '', parts[1]
    return parts[0], ' '.join(parts[1:-2]), ' '.join(parts[-2:])

rows = []
for r in ws.iter_rows(min_row=2):
    vals = [c.value for c in r]
    def get(col_key, default_idx):
        if col_key in header_map:
            idx = header_map[col_key]
            return vals[idx] if idx < len(vals) else None
        return vals[default_idx] if default_idx < len(vals) else None

    if full_col is not None and full_col < len(vals) and vals[full_col]:
        f, m, l = split_full_name(vals[full_col])
    else:
        f = get('fname', 1)
        m = get('mname', 2)
        l = get('lname', 3)

    rows.append({
        'id': get('id', 0),
        'fname': f,
        'mname': m,
        'lname': l,
        'position': get('position', 4),
        'class': get('class', 5),
        'fullname': vals[full_col] if (full_col is not None and full_col < len(vals)) else None
    })

sys.stdout.write(json.dumps(rows, ensure_ascii=True))
'@

$users = $null
$Py = Get-PythonForOpenpyxl

if ($Py) {
    Write-Host "Using Python (openpyxl): $Py" -ForegroundColor DarkGray
    # Encoding note: Python emits the JSON with ensure_ascii=True, so the byte
    # stream is pure 7-bit ASCII ('MU\u00d1OZ', not 'MUÑOZ'). That makes the read
    # immune to the console code page - without it, UTF-8 bytes were being
    # decoded as CP437 and 'MUÑOZ' arrived as 'MU├æOZ', which then failed the
    # accent-strip and produced a broken UPN (akina.mu├æoz@...).
    # ConvertFrom-Json turns the \u escapes back into real characters.
    # PYTHONIOENCODING is still set so any non-ASCII *stderr* text does not
    # raise UnicodeEncodeError inside Python.
    $prevEnc = $env:PYTHONIOENCODING
    $env:PYTHONIOENCODING = 'utf-8'
    try {
        $raw = & $Py -c $readPyCode $XlsxPath 2>&1
        $pyExit = $LASTEXITCODE
    } finally {
        if ($null -eq $prevEnc) { Remove-Item Env:PYTHONIOENCODING -ErrorAction SilentlyContinue }
        else { $env:PYTHONIOENCODING = $prevEnc }
    }

    if ($pyExit -eq 0 -and $raw) {
        $jsonText = (@($raw) | Where-Object { $_ -is [string] }) -join ''
        if (-not [string]::IsNullOrWhiteSpace($jsonText)) {
            # NOTE: do not wrap in @() - ConvertFrom-Json already returns an array
            # for a top-level JSON array, and @() would re-wrap it as one element.
            try {
                $parsed = $jsonText | ConvertFrom-Json
                if ($null -ne $parsed) { $users = @($parsed | ForEach-Object { $_ }) }
            } catch { $users = $null }
        }
    }
    if ($null -eq $users) {
        Write-Warning "Python/Openpyxl read failed (exit $pyExit); falling back to ImportExcel."
    }
}

if ($null -eq $users) {
    $users = Read-UsersWithImportExcel -Path $XlsxPath
    if ($null -ne $users) { Write-Host "Using ImportExcel module." -ForegroundColor DarkGray }
}

if ($null -eq $users) {
    Write-Error "Could not read '$XlsxPath'. Install one of: python -m pip install openpyxl  OR  Install-Module ImportExcel -Scope CurrentUser"
    exit 1
}

Write-Host "Found $($users.Count) rows." -ForegroundColor Cyan

# --- HELPER: ACCENT STRIP ---
function Strip-Accents($text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return $text }
    $norm = $text.Normalize([System.Text.NormalizationForm]::FormD)
    $stripped = $norm -replace '\p{Mn}', ''
    return $stripped.ToLower()
}

function To-TitleCase($text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return $text }
    $parts = $text.Split(' ')
    $out = @()
    foreach ($p in $parts) {
        $trim = $p.Trim()
        if ($trim.Length -gt 0) {
            $out += ($trim.Substring(0,1).ToUpper() + $trim.Substring(1).ToLower())
        }
    }
    return ($out -join ' ')
}

# --- HELPER: USERNAME ---
function Get-Username($fname, $lname) {
    $firstWord = if ($fname) { (($fname -split ' ')[0]).Trim() } else { "" }
    $cleanFirst = Strip-Accents $firstWord
    $cleanLast  = if ($lname) { (Strip-Accents $lname).Trim() -replace '\s+', '' } else { "" }
    return "$cleanFirst.$cleanLast"
}

# --- CONNECT CHECK ---
try {
    $ctx = Get-MgContext -ErrorAction Stop
    if (-not $ctx) { throw "No Graph context." }
} catch {
    Write-Warning "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes 'User.ReadWrite.All,Group.ReadWrite.All,Directory.Read.All'"
    Write-Warning "(optional: UserManagementPolicy.ReadWrite.All,Organization.Read.All for licenses)"
    # Continue; creation calls will fail with clear errors if missing.
}

# --- INTERACTIVE GROUP SELECTION ---
Write-Host "`n--- GROUP SELECTION ---" -ForegroundColor Yellow
$groupsRaw = $null; $groupError = $null
try {
    $groupsRaw = Get-MgGroup -Property "displayName,id" -All -ErrorAction Stop
} catch {
    $groupError = $_.Exception.Message
    try {
        $groupsRaw = Get-MgGroup -ErrorAction Stop
        $groupError = $null
    } catch {
        $groupError = $_.Exception.Message
    }
}
if (-not $groupsRaw) {
    Write-Warning "Could not retrieve groups. Error: $groupError - Check Directory.Read.All / Group.Read.All. Enter group IDs manually or press Enter to skip groups."
    $groupIds = @()
    $groupMap = @{}
    $groupNames = @{}
} else {
    $groupIds = @()
    $groupMap = @{}
    $groupNames = @{}
    $i = 1
    $groupsRaw | Sort-Object DisplayName | ForEach-Object {
        $groupMap[$i] = $_.Id
        $groupNames[$_.Id] = $_.DisplayName
        Write-Host ("{0,2}. {1}" -f $i, $_.DisplayName) -ForegroundColor Green
        $i++
    }
    Write-Host "Enter group numbers (comma-separated) or press Enter for none: " -NoNewline -ForegroundColor Yellow
    $sel = Read-Host
    if ($sel -match '\d+') {
        $sel.Split(',') | ForEach-Object { $n = [int]$_; if ($groupMap.ContainsKey($n)) { $groupIds += $groupMap[$n] } }
    }
    Write-Host "Selected $($groupIds.Count) group(s)." -ForegroundColor Green
}

# --- INTERACTIVE LICENSE SELECTION ---
Write-Host "`n--- LICENSE SELECTION ---" -ForegroundColor Yellow
$skusRaw = $null; $skuError = $null
try {
    # Plain call first: most compatible across Graph module versions.
    $skusRaw = Get-MgSubscribedSku -ErrorAction Stop
} catch {
    $skuError = $_.Exception.Message
    try {
        # Fallback: request only the properties the SDK is known to expose.
        $skusRaw = Get-MgSubscribedSku -Property "skuPartNumber,skuId" -ErrorAction Stop
        $skuError = $null
    } catch {
        $skuError = $_.Exception.Message
    }
}
if (-not $skusRaw) {
    Write-Warning "Could not retrieve SKUs. Error: $skuError - Check Organization.Read.All scope + admin consent. Enter SKU IDs manually or press Enter to skip licenses."
    $licenseIds = @()
    $skuMap = @{}
    $licenseNames = @{}
} else {
    $licenseIds = @()
    $skuMap = @{}
    $licenseNames = @{}
    $i = 1
    $skusRaw | Sort-Object SkuPartNumber | ForEach-Object {
        $skuMap[$i] = $_.SkuId
        $licenseNames[$i] = $_.SkuPartNumber
        Write-Host ("{0,2}. {1}" -f $i, $_.SkuPartNumber) -ForegroundColor Green
        $i++
    }
    Write-Host "Enter license numbers (comma-separated) or press Enter for none: " -NoNewline -ForegroundColor Yellow
    $sel = Read-Host
    if ($sel -match '\d+') {
        $sel.Split(',') | ForEach-Object { $n = [int]$_; if ($skuMap.ContainsKey($n)) { $licenseIds += $skuMap[$n] } }
    }
    Write-Host "Selected $($licenseIds.Count) license(s)." -ForegroundColor Green
}

# --- REPORT INIT ---
$report = @()

# --- PROCESS LOOP ---
foreach ($row in $users) {
    $fname = $row.fname
    $lname = $row.lname
    $mname = $row.mname
    $id = $row.id
    $srcFull = $row.fullname

    $username = Get-Username $fname $lname
    $upn = "$username@$Domain"

    $status = "Pending"
    $note = ""

    # Skip rows with blank names
    $tFname = if ($fname) { $fname.Trim() } else { "" }
    $tLname = if ($lname) { $lname.Trim() } else { "" }
    $tMname = if ($mname) { $mname.Trim() } else { "" }
    $displayName = (To-TitleCase "$tFname $tLname")
    # Combined full name: first + middle initial(s) + last (e.g. "Juan Flavio C. Dela Cruz")
    $fullName = (To-TitleCase "$tFname $tMname $tLname").Trim()

    # Check: does the reconstructed full name match the source full-name column?
    $nameCheck = ""
    if (-not [string]::IsNullOrWhiteSpace($srcFull)) {
        $srcNorm = (To-TitleCase ($srcFull -replace '\s+', ' ')).Trim()
        $builtNorm = ($fullName -replace '\s+', ' ').Trim()
        if ($srcNorm -eq $builtNorm) {
            $nameCheck = "OK"
        } else {
            $nameCheck = "MISMATCH (sheet: '$srcFull')"
            Write-Warning "Full name mismatch for ${upn}: sheet='$srcFull' built='$fullName'"
        }
    }
    if ([string]::IsNullOrWhiteSpace($tFname) -or [string]::IsNullOrWhiteSpace($tLname)) {
        $status = "Skipped - Blank name"
        $note = "fname=$fname lname=$lname"
        $report += [PSCustomObject]@{ ID=$id; Username=$username; UPN=$upn; Status=$status; Note=$note; DisplayName=$displayName; FullName=$fullName; NameCheck=$nameCheck }
        Write-Host "SKIP (blank) | $fname $lname" -ForegroundColor Gray
        continue
    }

    # Check existing
    $existing = $null
    try {
        $existing = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue -Property "id"
    } catch { }
    if ($existing -and $SkipExisting) {
        $status = "Skipped - Exists"
        $note = "Already present"
        $report += [PSCustomObject]@{ ID=$id; Username=$username; UPN=$upn; Status=$status; Note=$note; DisplayName=$displayName; FullName=$fullName; NameCheck=$nameCheck }
        Write-Host "SKIP (exists) | $upn" -ForegroundColor Yellow
        continue
    }

    if ($WhatIf) {
        $status = "WhatIf - Created"
        $note = "Would create + groups/licenses"
        $report += [PSCustomObject]@{ ID=$id; Username=$username; UPN=$upn; Status=$status; Note=$note; DisplayName=$displayName; FullName=$fullName; NameCheck=$nameCheck }
        Write-Host "WHATIF | $upn" -ForegroundColor Cyan
        continue
    }

    # Create user
    try {
        $given = if ($fname) { (($fname -split ' ')[0]).Trim() } else { "" }
        $fullFirst = if ($fname) { ($fname).Trim() } else { "" }
        $surname = if ($lname) { ($lname).Trim() } else { "" }
        $displayName = (To-TitleCase "$fullFirst $surname")
        $mailNickname = $username

        $newUser = New-MgUser -DisplayName $displayName `
            -GivenName $given `
            -Surname $surname `
            -UserPrincipalName $upn `
            -MailNickname $mailNickname `
            -AccountEnabled `
            -PasswordProfile @{ ForceChangePasswordNextSignIn = $true; ForceChangePasswordNextSignInWithMfa = $false; Password = $DefaultPassword } `
            -PasswordPolicies "DisablePasswordExpiration" `
            -UsageLocation $UsageLocation `
            -ErrorAction Stop

        $status = "Created"
        $note = "Account created"
        Write-Host "CREATED | $upn" -ForegroundColor Green
    } catch {
        $status = "Failed - Creation"
        $note = $_.Exception.Message
        $report += [PSCustomObject]@{ ID=$id; Username=$username; UPN=$upn; Status=$status; Note=$note; DisplayName=$displayName; FullName=$fullName; NameCheck=$nameCheck }
        Write-Host "FAIL | $upn : $note" -ForegroundColor Red
        continue
    }

    # Group joins
    if ($groupIds.Count -gt 0 -and $newUser.Id) {
        foreach ($gid in $groupIds) {
            $gName = if ($groupNames.ContainsKey($gid)) { $groupNames[$gid] } else { $gid }
            $gOk = $false
            $gErr = ""
            foreach ($attempt in 1..4) {
                try {
                    New-MgGroupMember -GroupId $gid -DirectoryObjectId $newUser.Id -ErrorAction Stop
                    $gOk = $true
                    break
                } catch {
                    $gErr = $_.Exception.Message
                    # "already exist" means the user is already in the group -> treat as success
                    if ($gErr -match "already exist") { $gOk = $true; break }
                    Start-Sleep -Seconds (2 * $attempt)
                }
            }
            if ($gOk) {
                $note += " +group($gName)"
            } else {
                $note += " !group($gName)"
                Write-Warning "Group join failed for ${upn} -> ${gName}: $gErr"
            }
        }
    }

    # License assignment
    if ($licenseIds.Count -gt 0 -and $newUser.Id) {
        # A newly created account is not immediately licenseable in Graph.
        # Ensure UsageLocation matches the tenant's country (PH), then retry.
        try {
            Update-MgUser -UserId $newUser.Id -UsageLocation $UsageLocation -ErrorAction SilentlyContinue
        } catch { }

        $licenseAssignments = @()
        foreach ($lid in $licenseIds) {
            $licenseAssignments += @{ SkuId = $lid; DisabledPlans = @() }
        }
        $licOk = $false
        $licErr = ""
        foreach ($attempt in 1..4) {
            try {
                Set-MgUserLicense -UserId $newUser.Id -AddLicenses $licenseAssignments -RemoveLicenses @() -ErrorAction Stop
                $licOk = $true
                break
            } catch {
                $licErr = $_.Exception.Message
                # Country/region mismatch or unprovisioned SKU will never succeed
                # on retry - fail fast instead of sleeping through four attempts.
                if ($licErr -match 'country|region|UsageLocation|not valid|does not exist|proxy address') { break }
                Start-Sleep -Seconds (3 * $attempt)
            }
        }
        if ($licOk) {
            $note += " +licenses($($licenseIds.Count))"
            Write-Host ("        licensed: {0} SKU(s) [UsageLocation {1}]" -f $licenseIds.Count, $UsageLocation) -ForegroundColor DarkGreen
        } else {
            $note += " !license"
            Write-Warning "License assignment failed for ${upn} (UsageLocation=${UsageLocation}): $licErr"
            Write-Warning "Check that the tenant is licensed for these SKUs in your country, and that the account's usage location is ${UsageLocation}."
        }
    }

    $report += [PSCustomObject]@{ ID=$id; Username=$username; UPN=$upn; Status=$status; Note=$note; DisplayName=$displayName; FullName=$fullName; NameCheck=$nameCheck }
}

# --- CSV EXPORT ---
# Create the target folder first; 'Output Report' may not exist on a fresh checkout.
$csvDir = Split-Path -Parent $CsvReport
if ($csvDir -and -not (Test-Path -LiteralPath $csvDir)) {
    [void][System.IO.Directory]::CreateDirectory($csvDir)
}
$report | Export-Csv -Path $CsvReport -NoTypeInformation -Encoding UTF8 -Force
Write-Host "`nReport exported to: $CsvReport ($($report.Count) rows)" -ForegroundColor Magenta

# --- SUMMARY ---
$created = ($report | Where-Object { $_.Status -match "Created" }).Count
$skipped = ($report | Where-Object { $_.Status -match "Skipped" }).Count
$failed  = ($report | Where-Object { $_.Status -match "Failed" }).Count
Write-Host "Summary: Created=$created | Skipped=$skipped | Failed=$failed | WhatIf=$($WhatIf.IsPresent)" -ForegroundColor Yellow
