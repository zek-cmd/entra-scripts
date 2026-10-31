# Interactive launcher for Entra scripts
Clear-Host
Write-Host '========================================' -ForegroundColor Cyan
Write-Host '  Entra ID Bulk Operations Menu' -ForegroundColor Cyan
Write-Host '========================================' -ForegroundColor Cyan
Write-Host ''
Write-Host '[1] Move / Add users between groups (add_or_move_Entra-Users.ps1)' -ForegroundColor Gray
Write-Host '[2] Bulk delete users (Bulk-Delete-Entra-Users.ps1, Excel or group)' -ForegroundColor Gray
Write-Host '[3] Bulk create users (bulk_create_user.ps1)' -ForegroundColor Gray
Write-Host '[4] Remove licenses (Remove-Entra-Licenses.ps1)' -ForegroundColor Gray
Write-Host '[5] Find duplicate or near-duplicate names (Find-Duplicate-Entra-Names.ps1, read-only)' -ForegroundColor Gray
Write-Host '[6] Exit' -ForegroundColor DarkGray
Write-Host ''

$choice = (Read-Host 'Select number (1-6)').Trim()
if ($choice -eq '6' -or $choice -eq '') { Write-Host 'Exiting.'; exit 0 }
if ($choice -notin @('1','2','3','4','5')) { Write-Host 'Invalid selection. Exiting.' -ForegroundColor Red; exit 1 }

# Option 5 only reads data, so it skips the dry-run question
if ($choice -eq '5') {
    Write-Host ''
    & (Join-Path $PSScriptRoot 'Find-Duplicate-Entra-Names.ps1')
    exit 0
}

# Dry-run preference: default safe (dry run / WhatIf)
$dry = (Read-Host 'Dry run only? (no changes will be made) [Y/n]').Trim().ToLower()
if ($dry -eq '' -or $dry -eq 'y' -or $dry -eq 'yes') {
    $WhatIfPreference = $true
    Write-Host 'Mode: DRY RUN ($WhatIfPreference = $true) - nothing will change.' -ForegroundColor Cyan
} else {
    $WhatIfPreference = $false
    Write-Host 'Mode: LIVE ($WhatIfPreference = $false) - changes will apply.' -ForegroundColor Yellow
}

Write-Host ''
switch ($choice) {
    # These use CmdletBinding(SupportsShouldProcess) -> honour $WhatIfPreference
    '1' { & (Join-Path $PSScriptRoot 'add_or_move_Entra-Users.ps1') }
    '2' { & (Join-Path $PSScriptRoot 'Bulk-Delete-Entra-Users.ps1') }
    '4' { & (Join-Path $PSScriptRoot 'Remove-Entra-Licenses.ps1') }
    # This one declares an explicit [switch]$WhatIf -> must pass the argument
    '3' {
        if ($WhatIfPreference) { & (Join-Path $PSScriptRoot 'bulk_create_user.ps1') -WhatIf }
        else { & (Join-Path $PSScriptRoot 'bulk_create_user.ps1') }
    }
}
