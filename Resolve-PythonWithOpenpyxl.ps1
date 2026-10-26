function Resolve-PythonWithOpenpyxl {
    <#
    .SYNOPSIS
        Dynamically locates a Python interpreter that has openpyxl available.

    .DESCRIPTION
        Replaces the hardcoded 'C:\Users\<name>\.dsh\...\python.exe' path used by
        add_or_move_Entra-Users.ps1, Bulk-Delete-Entra-Users.ps1 and
        Remove-Entra-Licenses.ps1.

        Searches, in order:
          1. $env:ENTRA_PYTHON        - explicit override
          2. DSH bundled runtime      - built from $env:USERPROFILE, not a literal user name
          3. py launcher registry     - 'py -0p' lists every registered install
          4. Registry PythonCore      - HKLM, HKLM WOW6432Node, HKCU
          5. python / python3 on PATH - Store 0-byte alias stubs rejected

        Only an interpreter that can actually resolve openpyxl is returned.

    .PARAMETER SkipBundled
        Ignore the DSH bundled interpreter. Useful for testing what a machine
        without the DSH runtime would do.

    .PARAMETER NoCache
        Force a fresh search even if a previous call already resolved an
        interpreter. Without this, a successful result is cached for the session
        so repeated calls are free.

    .OUTPUTS
        System.String - full path to a usable interpreter, or $null.

    .NOTES
        Never throws. Callers decide what to do with $null.
        Probe uses importlib.util.find_spec: ~80 ms vs ~1400 ms for a real
        'import openpyxl', because it avoids compiling the package.

        Caching matters here: the Excel helpers are called once per run for
        headers and again for data, so an uncached resolver costs ~260 ms each
        time against ~0 ms when cached. The cache lives in the caller's session
        (script scope of this dotted file), so it disappears with the session.
    #>
    param(
        [switch]$SkipBundled,
        [switch]$NoCache
    )

    # Session cache: only successful, non-SkipBundled lookups are stored, so a
    # -SkipBundled probe never poisons the normal path.
    if (-not $NoCache -and -not $SkipBundled -and $script:ResolvedPythonPath) {
        if (Test-Path -LiteralPath $script:ResolvedPythonPath) { return $script:ResolvedPythonPath }
        $script:ResolvedPythonPath = $null   # went stale, fall through and re-scan
    }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $candidates = New-Object 'System.Collections.Generic.List[string]'

    # 1) Explicit override
    if ($env:ENTRA_PYTHON) { $candidates.Add($env:ENTRA_PYTHON) }

    # 2) DSH bundled runtime - derived from the profile, so it survives a rename
    if (-not $SkipBundled -and $env:USERPROFILE) {
        $candidates.Add((Join-Path $env:USERPROFILE '.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe'))
    }

    # 3) py launcher - covers pyenv, Store installs and custom locations
    $pyLauncher = Get-Command py -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        foreach ($line in (& $pyLauncher.Source -0p 2>$null)) {
            # format: " -V:3.13 *        C:\path\to\python.exe"
            if ($line -match '([A-Za-z]:\\.*?python(?:w)?\.exe)\s*$') { $candidates.Add($Matches[1]) }
        }
    }

    # 4) Registry - note 3.12-style entries can lack InstallPath, hence the guard
    foreach ($root in @('HKLM:\SOFTWARE\Python\PythonCore',
                        'HKLM:\SOFTWARE\WOW6432Node\Python\PythonCore',
                        'HKCU:\SOFTWARE\Python\PythonCore')) {
        if (-not (Test-Path $root)) { continue }
        foreach ($ver in (Get-ChildItem $root -ErrorAction SilentlyContinue)) {
            $ip = Join-Path $ver.PSPath 'InstallPath'
            if (-not (Test-Path $ip)) { continue }
            $props = Get-ItemProperty -Path $ip -ErrorAction SilentlyContinue
            foreach ($val in @('ExecutablePath', '(default)')) {
                $p = $props.$val
                if (-not $p) { continue }
                if ($p -notmatch '\.exe$') { $p = Join-Path $p 'python.exe' }
                $candidates.Add($p)
            }
        }
    }

    # 5) PATH - Get-Command -All, so a second entry is not hidden by the first
    foreach ($name in @('python', 'python3')) {
        foreach ($cmd in (Get-Command $name -All -ErrorAction SilentlyContinue)) {
            $src = $cmd.Source
            if (-not $src -or -not (Test-Path -LiteralPath $src)) { continue }
            # The Microsoft Store alias is a 0-byte reparse point that exits 9009.
            # Test-Path reports it as present, so size is the only reliable filter.
            $item = Get-Item -LiteralPath $src -ErrorAction SilentlyContinue
            if ($item -and $item.Length -eq 0) { continue }
            $candidates.Add($src)
        }
    }

    $probe = "import importlib.util,sys; sys.exit(0 if importlib.util.find_spec('openpyxl') else 1)"
    foreach ($c in $candidates) {
        if (-not $c -or -not (Test-Path -LiteralPath $c)) { continue }
        $full = (Resolve-Path -LiteralPath $c).Path
        if (-not $seen.Add($full)) { continue }
        $null = & $full -c $probe 2>&1
        if ($LASTEXITCODE -eq 0) {
            if (-not $SkipBundled) { $script:ResolvedPythonPath = $full }
            return $full
        }
    }
    return $null
}
