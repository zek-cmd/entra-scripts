#Requires -Version 5.1
<#
.SYNOPSIS
    Manually type a user's name and set the account's Office / Department to
    "Provision Fix". Works on one user, or on a list you paste in.

.DESCRIPTION
    You type a name (or part of a name, or a sign-in name / email). The script
    searches Entra ID and shows the closest matches so you can pick the right
    person. Once you confirm, it runs the same update as:

        Update-MgUser -UserId "username@yourdomain.com" -Department "Provision Fix"

    Safety features:
      * -WhatIf     -> dry run. Shows what would change, changes nothing.
      * -Confirm    -> asks before each account is changed.
      * Audit log   -> CSV of every account touched, in the "Output Report" folder.
      * Id check    -> users who already say "Provision Fix" are left alone.

    Close matches work like this:
      * Words are compared, not exact spelling, so "Jon Smith" finds "John Smith".
      * Word order does not matter, so "Smith Jon" also finds "John Smith".
      * Spelling differences are scored with letter edits, so small typos still show.
      * Results are ranked 0-100 and shown as a numbered list you pick from.

    If nothing matches, you can pick any of the suggestions, or the script asks
    again - it never guesses by itself.

.PARAMETER User
    A name, sign-in name (UPN) or email. If omitted, the script asks for it.

.PARAMETER Department
    The department value written to the account. Default: "Provision Fix".

.PARAMETER MatchCount
    How many close matches to show for each name. Default 8.

.PARAMETER MinScore
    Lowest score a suggestion may have (0-100). Default 40. Lower it to see
    more names, raise it to see only strong matches.

.PARAMETER LogPath
    Audit log CSV path. Defaults to 'ProvisionFix_<timestamp>.csv' in the
    "Output Report" folder next to this script.

.EXAMPLE
    # Asks for the name, uses "Provision Fix", dry run first
    .\Set-Entra-UserDepartment.ps1 -WhatIf

.EXAMPLE
    # Straight to one account, live
    .\Set-Entra-UserDepartment.ps1 -User "john.smith@contoso.com"

.EXAMPLE
    # Name typed directly, no menu
    .\Set-Entra-UserDepartment.ps1 -User "John Smith" -Confirm
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false)]
    [string]$User,

    [string]$Department = 'Provision Fix',

    [ValidateRange(1, 30)]
    [int]$MatchCount = 8,

    [ValidateRange(0, 100)]
    [int]$MinScore = 40,

    [string]$LogPath
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------
$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) { $scriptRoot = (Get-Location).Path }

if (-not $LogPath) {
    $LogPath = Join-Path (Join-Path $scriptRoot 'Output Report') ('ProvisionFix_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
$script:LogFile = $LogPath

# ---------------------------------------------------------------------------
# Audit log (written through .NET so -WhatIf cannot suppress it)
# ---------------------------------------------------------------------------
function Write-Log {
    param(
        [string]$Typed, [string]$DisplayName, [string]$UserId,
        [string]$Upn, [string]$OldDepartment, [string]$NewDepartment,
        [string]$Result, [string]$Message, [string]$Score
    )
    $line = @(
        $Typed, $DisplayName, $UserId, $Upn, $OldDepartment, $NewDepartment,
        $Result, $Score, (($Message -replace "`r|`n", ' ')),
        (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    ) -join '","'
    $line = '"' + $line + '"'
    $enc = New-Object System.Text.UTF8Encoding($true)
    $dir = Split-Path -Parent $script:LogFile
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { [void][System.IO.Directory]::CreateDirectory($dir) }
    if (-not (Test-Path -LiteralPath $script:LogFile)) {
        $header = '"TypedInput","DisplayName","UserId","UserPrincipalName","OldDepartment","NewDepartment","Result","Score","Message","Timestamp"'
        [System.IO.File]::AppendAllText($script:LogFile, $header + [Environment]::NewLine, $enc)
    }
    [System.IO.File]::AppendAllText($script:LogFile, $line + [Environment]::NewLine, $enc)
}

# ---------------------------------------------------------------------------
# Name matching helpers (close matches)
# ---------------------------------------------------------------------------
function ConvertTo-NameTokens {
    param([string]$Text)
    return @((([string]$Text).ToLowerInvariant() -replace '[^\p{L}\p{N}]+', ' ') -split ' ' | Where-Object { $_ })
}

function Get-EditDistance {
    <# Number of single-letter edits that turns A into B. #>
    param([string]$A, [string]$B)
    $n = $A.Length
    $m = $B.Length
    if ($n -eq 0) { return $m }
    if ($m -eq 0) { return $n }
    $prev = New-Object 'int[]' ($m + 1)
    $cur = New-Object 'int[]' ($m + 1)
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
    <# 0..1 how alike two tidy strings are, based on letter edits. #>
    param([string]$A, [string]$B)
    $max = [Math]::Max($A.Length, $B.Length)
    if ($max -eq 0) { return 1.0 }
    $d = Get-EditDistance -A $A -B $B
    return [Math]::Round(1.0 - ($d / $max), 3)
}

function Get-NameKey {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return ((@(ConvertTo-NameTokens -Text $Text)) -join ' ')
}

function Get-SortedWordKey {
    param($Tokens)
    if (@($Tokens).Count -eq 0) { return '' }
    $words = @($Tokens)
    [array]::Sort($words)
    return ($words -join ' ')
}

function Get-AllUsers {
    <# Loads every user once and caches it. Only the three fields we need. #>
    if ($script:AllUsersCache) { return $script:AllUsersCache }
    Write-Host ' Loading users from Entra ID...' -ForegroundColor DarkGray
    $script:AllUsersCache = @(Get-MgUser -All -Property @('Id', 'DisplayName', 'UserPrincipalName', 'Mail', 'Department') -ErrorAction Stop |
        Select-Object Id, DisplayName, UserPrincipalName, Mail, Department)
    Write-Host (' Loaded {0} users.' -f $script:AllUsersCache.Count) -ForegroundColor DarkGray
    return $script:AllUsersCache
}

function Find-UserCandidates {
    <# Returns the best matching users for a typed name, scored 0-100.
       Exact sign-in name / email / display name always scores 100 and comes first. #>
    param(
        [string]$Value,
        [object[]]$Users,
        [int]$Max = 8,
        [int]$Floor = 40
    )

    $typed = ([string]$Value).Trim()
    if ($typed -eq '') { return @() }
    $typedKey = Get-NameKey -Text $typed
    $typedTokens = @(ConvertTo-NameTokens -Text $typed)
    $typedSorted = Get-SortedWordKey -Tokens $typedTokens
    $typedHasSpace = $typedTokens.Count -gt 1

    $results = foreach ($u in $Users) {
        $name = [string]$u.DisplayName
        $upn = [string]$u.UserPrincipalName
        $mail = [string]$u.Mail

        $score = 0
        $why = ''

        # --- exact hits first ---
        if ([string]::Equals($upn, $typed, [StringComparison]::OrdinalIgnoreCase)) {
            $score = 100; $why = 'exact sign-in name'
        } elseif ($mail -and [string]::Equals($mail, $typed, [StringComparison]::OrdinalIgnoreCase)) {
            $score = 100; $why = 'exact email'
        } elseif ([string]::Equals($name, $typed, [StringComparison]::OrdinalIgnoreCase)) {
            $score = 100; $why = 'exact display name'
        } elseif ([string]::Equals((Get-NameKey -Text $name), $typedKey, [StringComparison]::OrdinalIgnoreCase)) {
            $score = 98; $why = 'same name, different spacing or punctuation'
        } else {
            # --- word overlap: word order ignored, partial words allowed ---
            $candTokens = @(ConvertTo-NameTokens -Text $name)
            if ($candTokens.Count -gt 0 -and $typedTokens.Count -gt 0) {
                $shared = 0
                foreach ($t in $typedTokens) {
                    $hit = $candTokens | Where-Object {
                        $_ -eq $t -or ($t.Length -ge 3 -and $_.StartsWith($t)) -or ($_.Length -ge 3 -and $t.StartsWith($_))
                    } | Select-Object -First 1
                    if ($hit) { $shared++ }
                }
                if ($shared -gt 0) {
                    $union = @(@($typedTokens) + @($candTokens) | Select-Object -Unique).Count
                    $score = [int][Math]::Round(100 * $shared / $union)
                    $why = ('{0} of {1} words match' -f $shared, $typedTokens.Count)
                }
            }

            # --- spelling closeness for typos, and for word-order swaps ---
            $nameKey = Get-NameKey -Text $name
            if ($nameKey -ne '' -and $typedKey -ne '') {
                $sim = Get-Similarity -A $typedKey -B $nameKey
                $simSorted = Get-Similarity -A $typedSorted -B (Get-SortedWordKey -Tokens @(ConvertTo-NameTokens -Text $name))
                $best = [Math]::Max($sim, $simSorted)
                # Whole-string likeness is only trusted for full-name or single-name typing,
                # otherwise one shared letter would score far too high.
                if (-not $typedHasSpace -or $typedTokens.Count -le 3) {
                    $spellScore = [int][Math]::Round(100 * $best)
                    if ($spellScore -gt $score) {
                        $score = $spellScore
                        $why = if ($simSorted -ge $sim) { 'words swapped, spelled close' } else { 'spelling is close' }
                    }
                }
            }
        }

        if ($score -lt $Floor) { continue }
        [pscustomobject]@{
            Id          = $u.Id
            Name        = if ($name) { $name } else { [string]$upn }
            Upn         = $upn
            Mail        = $mail
            Department  = $u.Department
            Score       = [int]$score
            Why         = $why
        }
    }

    return @($results | Sort-Object -Property @{Expression = 'Score'; Descending = $true}, @{Expression = 'Name'; Descending = $false} | Select-Object -First $Max)
}

# ---------------------------------------------------------------------------
# Entra connection
# ---------------------------------------------------------------------------
function Connect-EntraForUsers {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        throw 'Microsoft.Graph module is missing. Install it with: Install-Module Microsoft.Graph -Scope CurrentUser'
    }
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
        # User.ReadWrite.All is needed because this script changes the account.
        Connect-MgGraph -Scopes 'User.Read.All', 'User.ReadWrite.All' | Out-Null
    }
}

# ---------------------------------------------------------------------------
# Pick a user for one typed name
# ---------------------------------------------------------------------------
function Select-UserInteractive {
    <# Shows the close matches and returns the chosen user object, or $null. #>
    param([string]$Value, [object[]]$Users)

    $candidates = @(Find-UserCandidates -Value $Value -Users $Users -Max $MatchCount -Floor $MinScore)

    # One perfect hit and the name is unique -> no menu, just go.
    $perfect = @($candidates | Where-Object { $_.Score -eq 100 })
    if ($perfect.Count -eq 1) {
        Write-Host (' MATCH 100 : {0}  ({1})  {2}' -f $perfect[0].Name, $perfect[0].Upn, $perfect[0].Why) -ForegroundColor DarkGray
        return $perfect[0]
    }

    if ($candidates.Count -eq 0) {
        Write-Host (' NO MATCH  : {0}' -f $Value) -ForegroundColor DarkYellow
        $retry = (Read-Host ' Nothing close found. Type another name, or press Enter to skip').Trim()
        if ($retry -eq '') { return $null }
        return (Select-UserInteractive -Value $retry -Users $Users)
    }

    Write-Host (' CLOSEST MATCHES for "{0}":' -f $Value) -ForegroundColor Cyan
    $n = 0
    foreach ($c in $candidates) {
        $n++
        $deptText = if ([string]::IsNullOrWhiteSpace([string]$c.Department)) { '(no dept)' } else { $c.Department }
        Write-Host ('   [{0}] {1,3}%  {2}  <{3}>  dept: {4}  - {5}' -f $n, $c.Score, $c.Name, $c.Upn, $deptText, $c.Why) -ForegroundColor Gray
    }
    Write-Host '   [S] search again with a different name' -ForegroundColor DarkGray
    Write-Host '   [Q] skip this name' -ForegroundColor DarkGray

    while ($true) {
        $pick = (Read-Host (' Pick a number 1-{0}, S to search again, Q to skip' -f $candidates.Count)).Trim()
        if ($pick -eq '') { continue }
        if ($pick -match '^\d+$') {
            $idx = [int]$pick
            if ($idx -ge 1 -and $idx -le $candidates.Count) { return $candidates[$idx - 1] }
            Write-Host '  Number is not in the list.' -ForegroundColor Red
            continue
        }
        if ($pick -match '^[Ss]$') {
            $again = (Read-Host ' Type the name again').Trim()
            if ($again -eq '') { continue }
            return (Select-UserInteractive -Value $again -Users $Users)
        }
        if ($pick -match '^[Qq]$') { return $null }
        Write-Host '  Enter a number, S or Q.' -ForegroundColor Red
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
Connect-EntraForUsers
$allUsers = @(Get-AllUsers)

$department = ([string]$Department)
Write-Host ''
Write-Host '========================================' -ForegroundColor Cyan
Write-Host '  Set Department: Provision Fix' -ForegroundColor Cyan
Write-Host '========================================' -ForegroundColor Cyan
Write-Host (' Department value : {0}' -f $department) -ForegroundColor Cyan
Write-Host (' Command it runs  : Update-MgUser -UserId "<sign-in name>" -Department "{0}"' -f $department) -ForegroundColor DarkGray

$stats = [ordered]@{ Changed = 0; Already = 0; Skipped = 0; Failed = 0; WhatIf = 0 }

# One name from the command line, otherwise ask in a loop so several users can be done.
$queue = @()
if (-not [string]::IsNullOrWhiteSpace($User)) {
    $queue = @($User)
} else {
    Write-Host ''
    Write-Host ' Type a name, part of a name, or a sign-in name / email.' -ForegroundColor Gray
    Write-Host ' Leave it blank and press Enter when you are finished.' -ForegroundColor DarkGray
    while ($true) {
        $typed = (Read-Host ' User name').Trim()
        if ($typed -eq '') { break }
        $queue += $typed
        Write-Host (' {0} name(s) queued.' -f $queue.Count) -ForegroundColor DarkGray
    }
    if ($queue.Count -eq 0) {
        Write-Host ' Nothing entered. Nothing to do.' -ForegroundColor Yellow
        return
    }
}

foreach ($typedName in $queue) {
    Write-Host ''
    $chosen = Select-UserInteractive -Value $typedName -Users $allUsers
    if (-not $chosen) {
        $stats.Skipped++
        Write-Host (' SKIPPED   : {0} (nothing picked)' -f $typedName) -ForegroundColor Yellow
        Write-Log -Typed $typedName -DisplayName '' -UserId '' -Upn '' -OldDepartment '' -NewDepartment $department -Result 'SkippedByOperator' -Message 'No user picked' -Score ''
        continue
    }

    $oldDept = [string]$chosen.Department
    if ($oldDept -eq $department) {
        $stats.Already++
        Write-Host (' ALREADY   : {0} already says "{1}"' -f $chosen.Name, $department) -ForegroundColor DarkGray
        Write-Log -Typed $typedName -DisplayName $chosen.Name -UserId $chosen.Id -Upn $chosen.Upn -OldDepartment $oldDept -NewDepartment $department -Result 'AlreadySet' -Message 'No change needed' -Score $chosen.Score
        continue
    }

    $actionDesc = ('Set department from "{0}" to "{1}"' -f $(if ($oldDept) { $oldDept } else { '(none)' }), $department)
    if (-not $PSCmdlet.ShouldProcess($chosen.Name, $actionDesc)) {
        $stats.WhatIf++
        Write-Host (' WHATIF    : would set {0} ({1}) to "{2}"' -f $chosen.Name, $chosen.Upn, $department) -ForegroundColor DarkCyan
        Write-Log -Typed $typedName -DisplayName $chosen.Name -UserId $chosen.Id -Upn $chosen.Upn -OldDepartment $oldDept -NewDepartment $department -Result 'WhatIf' -Message 'Not executed' -Score $chosen.Score
        continue
    }

    try {
        # The exact call requested, made by name so it is easy to read.
        Update-MgUser -UserId $chosen.Upn -Department $department -ErrorAction Stop
        $stats.Changed++
        Write-Host (' CHANGED   : {0} ({1}) -> department "{2}"' -f $chosen.Name, $chosen.Upn, $department) -ForegroundColor Green
        Write-Log -Typed $typedName -DisplayName $chosen.Name -UserId $chosen.Id -Upn $chosen.Upn -OldDepartment $oldDept -NewDepartment $department -Result 'Success' -Message '' -Score $chosen.Score
    } catch {
        $stats.Failed++
        Write-Host (' FAILED    : {0} -> {1}' -f $chosen.Name, $_.Exception.Message) -ForegroundColor Red
        Write-Host '             You need the User.ReadWrite.All permission and an admin consent.' -ForegroundColor DarkGray
        Write-Log -Typed $typedName -DisplayName $chosen.Name -UserId $chosen.Id -Upn $chosen.Upn -OldDepartment $oldDept -NewDepartment $department -Result 'Failed' -Message $_.Exception.Message -Score $chosen.Score
    }
}

Write-Host ''
Write-Host '----------------------------------------'
Write-Host 'Results:' -ForegroundColor Cyan
Write-Host (' Changed      : {0}' -f $stats.Changed)
Write-Host (' Already set  : {0}' -f $stats.Already)
Write-Host (' Skipped      : {0}' -f $stats.Skipped)
Write-Host (' Failed       : {0}' -f $stats.Failed)
Write-Host (' WhatIf       : {0}' -f $stats.WhatIf)
Write-Host (' Audit log    : {0}' -f $LogPath)
if ($WhatIfPreference) { Write-Host 'WhatIf mode active - nothing changed.' -ForegroundColor Cyan }
