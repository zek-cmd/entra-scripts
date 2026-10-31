#Requires -Version 5.1
<#
.SYNOPSIS
    Read-only scan of all Microsoft Entra ID (Azure AD) users for duplicate or
    near-duplicate display names. Changes nothing in Entra.

.DESCRIPTION
    Loads every user from the tenant (display name, sign-in name, mail),
    then reports:
      * Exact  - two or more users whose names are the same after tidying
                 (extra spaces, capital letters and punctuation are ignored).
      * Near   - names that look alike, such as "Jon Smith" and "John Smith",
                 or "Smith John" and "John Smith" (word order ignored).

    The result is shown on screen and saved as a CSV file you can open in Excel.
    The CSV is saved in the "Output Report" folder next to this script.

    Only read permission is requested (User.Read.All).

.PARAMETER Threshold
    How alike two names must be to count as a near duplicate, from 0 to 1.
    1 means identical. Default 0.85.

.PARAMETER OutputPath
    CSV report path. Default: DuplicateNames_<timestamp>.csv in the Output Report folder.

.EXAMPLE
    .\Find-Duplicate-Entra-Names.ps1

.EXAMPLE
    .\Find-Duplicate-Entra-Names.ps1 -Threshold 0.9
#>

[CmdletBinding()]
param(
    [ValidateRange(0.0, 1.0)]
    [double]$Threshold = 0.85,

    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

if (-not $OutputPath) {
    $folder = $PSScriptRoot
    if (-not $folder) { $folder = (Get-Location).Path }
    $folder = Join-Path $folder 'Output Report'
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $OutputPath = Join-Path $folder ('DuplicateNames_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

# ---------- helpers ----------

function Get-NameKey {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $k = $Name -replace '[^\p{L}\p{N} ]', ' '
    $k = ($k -replace '\s+', ' ').Trim().ToLowerInvariant()
    return $k
}

function Get-SortedWordKey {
    param([string]$Key)
    if ($Key -eq '') { return '' }
    $words = $Key.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
    [array]::Sort($words)
    return ($words -join ' ')
}

function Get-EditDistance {
    param([string]$A, [string]$B)
    $n = $A.Length
    $m = $B.Length
    if ($n -eq 0) { return $m }
    if ($m -eq 0) { return $n }
    $prev = New-Object 'int[]' ($m + 1)
    $cur  = New-Object 'int[]' ($m + 1)
    for ($j = 0; $j -le $m; $j++) { $prev[$j] = $j }
    for ($i = 1; $i -le $n; $i++) {
        $cur[0] = $i
        for ($j = 1; $j -le $m; $j++) {
            $cost = 1
            if ($A[$i - 1] -eq $B[$j - 1]) { $cost = 0 }
            $del = $prev[$j] + 1
            $ins = $cur[$j - 1] + 1
            $sub = $prev[$j - 1] + $cost
            $cur[$j] = [Math]::Min($del, [Math]::Min($ins, $sub))
        }
        $tmp = $prev; $prev = $cur; $cur = $tmp
    }
    return $prev[$m]
}

function Get-Similarity {
    param([string]$A, [string]$B)
    $max = [Math]::Max($A.Length, $B.Length)
    if ($max -eq 0) { return 1.0 }
    $d = Get-EditDistance -A $A -B $B
    return [Math]::Round(1.0 - ($d / $max), 3)
}

# ---------- connect (read-only) ----------

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    throw 'Microsoft.Graph module is missing. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser'
}
Import-Module Microsoft.Graph.Users -ErrorAction Stop
if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
    Connect-MgGraph -Scopes 'User.Read.All' | Out-Null
}

# ---------- load users ----------

Write-Host 'Loading users from Entra ID...' -ForegroundColor Cyan
$users = Get-MgUser -All -Property 'id,displayName,userPrincipalName,mail' |
    Select-Object Id, DisplayName, UserPrincipalName, Mail
Write-Host ("Loaded {0} users." -f @($users).Count) -ForegroundColor Cyan

$items = foreach ($u in $users) {
    $key = Get-NameKey -Name $u.DisplayName
    if ($key -eq '') { continue }
    [pscustomobject]@{
        Name    = $u.DisplayName
        Key     = $key
        Sorted  = Get-SortedWordKey -Key $key
        UPN     = $u.UserPrincipalName
        Mail    = $u.Mail
        Id      = $u.Id
    }
}
$items = @($items)

$report = New-Object System.Collections.Generic.List[object]

# ---------- exact duplicates ----------

$exactGroups = $items | Group-Object -Property Key | Where-Object { $_.Count -gt 1 }
foreach ($g in $exactGroups) {
    $list = @($g.Group)
    $report.Add([pscustomobject]@{
        Type       = 'Exact'
        Similarity = 1.0
        NameA      = $list[0].Name
        UPNA       = (($list | ForEach-Object { $_.UPN }) -join '; ')
        NameB      = ''
        UPNB       = ''
        Count      = $list.Count
    })
}

# ---------- near duplicates ----------
# Compare names only inside groups that share the same first letter, to keep it fast.
# Names with a different first letter are not compared (a typo in the first letter is missed).

$uniqueItems = @($items | Group-Object -Property Key | ForEach-Object { $_.Group[0] })
$blocks = $uniqueItems | Group-Object -Property { $_.Key.Substring(0, 1) }

foreach ($block in $blocks) {
    $list = @($block.Group)
    for ($i = 0; $i -lt $list.Count; $i++) {
        for ($j = $i + 1; $j -lt $list.Count; $j++) {
            $a = $list[$i]
            $b = $list[$j]
            if ([Math]::Abs($a.Key.Length - $b.Key.Length) -gt 4) { continue }

            $s1 = Get-Similarity -A $a.Key -B $b.Key
            $s2 = Get-Similarity -A $a.Sorted -B $b.Sorted
            $best = [Math]::Max($s1, $s2)

            if ($best -ge $Threshold -and $best -lt 1.0) {
                $report.Add([pscustomobject]@{
                    Type       = 'Near'
                    Similarity = $best
                    NameA      = $a.Name
                    UPNA       = $a.UPN
                    NameB      = $b.Name
                    UPNB       = $b.UPN
                    Count      = 2
                })
            }
        }
    }
}

# ---------- output ----------

$sorted = @($report | Sort-Object -Property @{Expression = 'Type'; Descending = $false}, @{Expression = 'Similarity'; Descending = $true})

if ($sorted.Count -eq 0) {
    Write-Host ''
    Write-Host 'No duplicate or near-duplicate names found.' -ForegroundColor Green
} else {
    $sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    $exactCount = @($sorted | Where-Object { $_.Type -eq 'Exact' }).Count
    $nearCount  = @($sorted | Where-Object { $_.Type -eq 'Near' }).Count
    Write-Host ''
    Write-Host ("Exact duplicate groups : {0}" -f $exactCount) -ForegroundColor Yellow
    Write-Host ("Near-duplicate pairs   : {0}" -f $nearCount) -ForegroundColor Yellow
    Write-Host ("Report saved to: {0}" -f $OutputPath) -ForegroundColor Cyan
    Write-Host ''
    $sorted | Select-Object -First 25 | Format-Table Type, Similarity, NameA, UPNA, NameB, UPNB -AutoSize
    if ($sorted.Count -gt 25) { Write-Host '(First 25 shown - see the CSV for all rows.)' -ForegroundColor DarkGray }
}

Write-Host 'Read-only scan finished. No users were changed.' -ForegroundColor Green
