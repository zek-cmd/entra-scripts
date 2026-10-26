# Required PowerShell Modules

Reference for the five Entra ID scripts in this folder. Every module below was
verified against what the scripts actually call — not guessed.

Verified environment: **Windows PowerShell 5.1.19041** with
**Microsoft.Graph 2.38.1** installed.

---

## 1. At a glance

| Module | Required? | Why |
| --- | --- | --- |
| `Microsoft.Graph` (meta) | **Yes** — or the 5 sub-modules in §3 | All cloud operations |
| `Microsoft.Graph.Authentication` | **Yes** | `Connect-MgGraph`, `Get-MgContext` |
| `Microsoft.Graph.Users` | **Yes** | User get/create/update/delete, license details |
| `Microsoft.Graph.Users.Actions` | **Yes** | `Set-MgUserLicense` |
| `Microsoft.Graph.Groups` | **Yes** | Group list + group membership |
| `Microsoft.Graph.Identity.DirectoryManagement` | **Yes** | `Get-MgSubscribedSku` (license SKU list) |
| `ImportExcel` | Optional | Faster `.xlsx` reading; auto-skipped if absent |
| `AzureAD` | Optional (legacy) | Falls back to this only if Graph is missing |
| Python + `openpyxl` | **Yes, in practice** | The `.xlsx` reader while `ImportExcel` is absent — see §10 |
| `Microsoft.PowerShell.Management` etc. | Built in | `Test-Path`, `Get-Content`, `Set-Content`, `Join-Path` … |

> **Note:** `AzureAD` / `AzureADPreview` is **not installed** on this machine, and
> that is fine. It is deprecated by Microsoft and every script prefers Graph when
> both are present. Install it only if a tenant blocks Graph access.

---

## 2. Framework prerequisites

Nothing to install, but the scripts depend on these:

| Requirement | Used by | Notes |
| --- | --- | --- |
| Windows PowerShell **5.1+** | all | `add_or_move_Entra-Users.ps1` has `#Requires -Version 5.1` |
| `System.Windows.Forms` | Excel picker dialogs | Loaded at runtime via `Add-Type`; ships with Windows PowerShell |
| PowerShell Gallery access | installation | Needs TLS 1.2 + NuGet provider (see §5) |

PowerShell 7 is not required, and the dialog code (`System.Windows.Forms`) is
happiest under 5.1.

---

## 3. Microsoft Graph — cmdlet to sub-module map

Each row verified with `Get-Command <cmdlet> | Select ModuleName`.

### Authentication
| Cmdlet | Module |
| --- | --- |
| `Connect-MgGraph` | `Microsoft.Graph.Authentication` |
| `Get-MgContext` | `Microsoft.Graph.Authentication` |

### Users
| Cmdlet | Module |
| --- | --- |
| `Get-MgUser` | `Microsoft.Graph.Users` |
| `New-MgUser` | `Microsoft.Graph.Users` |
| `Update-MgUser` | `Microsoft.Graph.Users` |
| `Remove-MgUser` | `Microsoft.Graph.Users` |
| `Get-MgUserLicenseDetail` | `Microsoft.Graph.Users` |
| `Remove-MgUserLicenseDetail` | `Microsoft.Graph.Users` |
| `Set-MgUserLicense` | `Microsoft.Graph.Users.Actions` |

### Groups
| Cmdlet | Module |
| --- | --- |
| `Get-MgGroup` | `Microsoft.Graph.Groups` |
| `Get-MgGroupMember` | `Microsoft.Graph.Groups` |
| `New-MgGroupMember` | `Microsoft.Graph.Groups` |
| `Remove-MgGroupMemberByRef` | `Microsoft.Graph.Groups` |

### Licenses / directory
| Cmdlet | Module |
| --- | --- |
| `Get-MgSubscribedSku` | `Microsoft.Graph.Identity.DirectoryManagement` |

> `Microsoft.Graph.Users.Actions` and `Microsoft.Graph.Identity.DirectoryManagement`
> are the two that are easy to forget. Installing `Microsoft.Graph.Users` alone
> leaves `Set-MgUserLicense` and `Get-MgSubscribedSku` unavailable.

---

## 4. Per-script requirements

| Script | Graph sub-modules needed | Excel module | Notes |
| --- | --- | --- | --- |
| `add_or_move_Entra-Users.ps1` | Authentication, Users, Groups | Optional | Group move only — no license cmdlets |
| `Bulk-Delete-Entra-Users.ps1` | Authentication, Users, Groups, **Users.Actions**, **Identity.DirectoryManagement** | Optional | Deletes + strips licenses; **group-scoped** |
| `Remove-Entra-Licenses.ps1` | Authentication, Users, Groups, **Users.Actions**, **Identity.DirectoryManagement** | Optional | Most demanding script |
| `bulk_create_user.ps1` | Authentication, Users, Groups, **Users.Actions**, **Identity.DirectoryManagement** | Not used | Inline Python reader + `ImportExcel` fallback; **no module guard — fails at runtime if Graph is absent** |
| `menu.ps1` | none | none | Launcher only; inherits the target script's needs |

**Simplest safe answer:** install the `Microsoft.Graph` meta-module once and all
five scripts are covered.

### 4.1 Bulk-Delete: two user sources, always group-scoped

`Bulk-Delete-Entra-Users.ps1` asks **where the users come from**:

```
=== Select user source ===
  [1] Excel file - read users from a spreadsheet
  [2] Group     - use all members of a selected group
```

`-Source Excel|Group` skips the menu. `-ExcelPath` implies `-Source Excel`, so
existing command lines keep working unchanged.

A group is required in **both** modes, but it does different work in each:

| Source | What the group does | Non-members |
| --- | --- | --- |
| `Excel` | **Scope filter** — bounds the blast radius of a wrong spreadsheet | Skipped, logged `NotInGroup` |
| `Group` | **Is the user list** — every member becomes a delete target | n/a (circular, so the filter is not applied) |

Other behaviour:

- Interactive numbered picker by default (same UX as `add_or_move_Entra-Users.ps1`).
- `-GroupName '<exact name or unique substring>'` skips the picker; an ambiguous
  or unknown name falls back to the picker rather than guessing.
- The script **refuses to run** if the group cannot be resolved to an object id,
  or if it has no readable members. There is no unscoped mode.
- In `Group` mode only **user** accounts are targeted. Nested groups, devices and
  service principals are counted and skipped, never passed to the delete loop.
- In `Group` mode `-ColumnName` is ignored (there is no spreadsheet).
- Requires `Group.Read.All` + `GroupMember.Read.All`, which is why those scopes
  were added to the `Connect-MgGraph` call.

### 4.2 Account creation: Philippines usage location

`bulk_create_user.ps1` sets `UsageLocation` on every new account and again
immediately before licence assignment.

- Default is **`PH`** (Philippines, ISO 3166-1 alpha-2). Override with
  `-UsageLocation <code>`; the value is validated as two letters and upper-cased.
- A licence **cannot** be assigned while the account's usage location is a
  country the tenant is not licensed for. This was previously hardcoded to `US`,
  which is a likely cause of silent licence failures on this tenant.
- Country/region errors now **fail fast** instead of sleeping through four
  retries (~30 s), and the warning names the `UsageLocation` in use. Transient
  errors (timeouts) still retry with backoff.
- `-DefaultPassword` exposes the initial password (default `P@ssw0rd`, with
  force-change at first sign-in).

---

## 5. Install

### Full install (recommended — covers every script)

```powershell
# One-time gallery prep (Windows PowerShell 5.1)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser

# Everything, all five scripts
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

### Minimal install (smaller download, same coverage)

```powershell
Install-Module Microsoft.Graph.Authentication,
               Microsoft.Graph.Users,
               Microsoft.Graph.Users.Actions,
               Microsoft.Graph.Groups,
               Microsoft.Graph.Identity.DirectoryManagement `
               -Scope CurrentUser -Force
```

> Keep the versions aligned. Mixing a 1.x `Users` with a 2.x `Authentication`
> produces `Assembly with same name is already loaded` errors.

### Optional extras

```powershell
Install-Module ImportExcel -Scope CurrentUser -Force   # faster .xlsx reading
Install-Module AzureAD     -Scope CurrentUser -Force   # legacy fallback only
```

---

## 6. Verify

Run this in the folder before using the scripts:

```powershell
$need = @(
  'Microsoft.Graph.Authentication',
  'Microsoft.Graph.Users',
  'Microsoft.Graph.Users.Actions',
  'Microsoft.Graph.Groups',
  'Microsoft.Graph.Identity.DirectoryManagement'
)
$need | ForEach-Object {
  $m = Get-Module -ListAvailable -Name $_
  if ($m) { "{0,-48} {1}" -f $_, ($m | Sort-Object Version -Descending | Select-Object -First 1).Version }
  else    { "{0,-48} MISSING" -f $_ }
}

# Cmdlet-level spot check
'Get-MgContext','Connect-MgGraph','Get-MgUser','New-MgUser','Set-MgUserLicense',
'Get-MgGroup','New-MgGroupMember','Get-MgSubscribedSku' | ForEach-Object {
  if (Get-Command $_ -ErrorAction SilentlyContinue) { "OK      $_" } else { "MISSING $_" }
}
```

The five `Get-Module -ListAvailable` calls the scripts themselves run can be
checked at the same time:

```powershell
Get-Module -ListAvailable -Name Microsoft.Graph.Users, Microsoft.Graph, AzureAD
```

---

## 7. Permissions (not installable — grant in Entra)

Modules only provide the cmdlets; the scopes differ per script. Grant admin
consent for the union if you run all four operations:

| Script | Scopes requested in code |
| --- | --- |
| `add_or_move_Entra-Users.ps1` | `User.Read.All`, `GroupMember.ReadWrite.All`, `Directory.Read.All` |
| `Bulk-Delete-Entra-Users.ps1` | `User.ReadWrite.All`, `User.Read.All`, `Directory.Read.All`, `Group.Read.All`, `GroupMember.Read.All` |
| `Remove-Entra-Licenses.ps1` | `User.Read.All`, `User.ReadWrite.All`, `Directory.Read.All`, `Organization.Read.All` |
| `bulk_create_user.ps1` | Not requested in code — run `Connect-MgGraph` manually with `User.ReadWrite.All`, `Group.ReadWrite.All`, `Directory.Read.All`, `UserManagementPolicy.ReadWrite.All`, `Organization.Read.All` |

---

## 8. Troubleshooting

| Symptom | Cause / fix |
| --- | --- |
| `Install-Module` prompts about an untrusted repository | Add `-Force`, or `Set-PSRepository -Name PSGallery -InstallationPolicy Trusted` |
| `Unable to find repository 'PSGallery'` | `Register-PSRepository -Default` |
| NuGet provider prompt blocks automation | Run the `Install-PackageProvider` line from §5 first |
| `Assembly with same name is already loaded` | Version mismatch between Graph sub-modules — `Update-Module Microsoft.Graph -Force` (or reinstall all at one version) |
| `Get-MgUser is not recognized` | `Microsoft.Graph.Users` missing, or module installed under a different PowerShell edition — check `$PSVersionTable` and `$env:PSModulePath` |
| `A parameter cannot be found that matches parameter name 'Filter'` | Old Graph SDK; `Get-MgUser -Filter` needs a 2.x module |
| Scripts blocked from running | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`; if the files were downloaded, `Get-ChildItem *.ps1 \| Unblock-File` |

---

## 9. Current state of this machine

```
Microsoft.Graph ................................ 2.38.1  (all sub-modules present)
Microsoft.Graph.Users ........................... 2.38.1
Microsoft.Graph.Users.Actions ................... 2.38.1
Microsoft.Graph.Groups .......................... 2.38.1
Microsoft.Graph.Identity.DirectoryManagement .... 2.38.1
ImportExcel ..................................... not installed (Python path in use)
AzureAD ......................................... not installed (not needed — Graph preferred)
Python (DSH bundled) ............................ 3.12.14 + openpyxl 3.1.5  OK
Python (PATH / laragon) ......................... 3.10      no openpyxl     <-- trap, see §10.3
```

All Graph cmdlets the scripts call resolve correctly here, and the bundled
Python at `.dsh\dsh-runtimes\...\python.exe` has `openpyxl`, so all four
functional scripts work on this machine as-is.

---

## 10. Python — who needs it and why

Short answer: **all four functional scripts need Python *today*, because
`ImportExcel` is not installed.** The three that use the shared helper look for
`ImportExcel` first and fall back to Python, so with no `ImportExcel` installed,
Python is the only path they have. Install `ImportExcel` and Python becomes
unnecessary for those three. `bulk_create_user.ps1` is Python-only and needs it
either way.

### 10.1 Dependency by script

| Script | Reads `.xlsx`? | Python needed now? |
| --- | --- | --- |
| `add_or_move_Entra-Users.ps1` | Yes (or use group source) | No — openpyxl primary, `ImportExcel` fallback |
| `Bulk-Delete-Entra-Users.ps1` | Yes (or use group source) | No — openpyxl primary, `ImportExcel` fallback |
| `Remove-Entra-Licenses.ps1` | Yes (or use group source) | No — openpyxl primary, `ImportExcel` fallback |
| `bulk_create_user.ps1` | Yes (always) | No — openpyxl primary, `ImportExcel` fallback |
| `menu.ps1` | No | No |

All four readers now try **Python/openpyxl first, `ImportExcel` second**. A
genuine "column not found" answer is re-thrown immediately rather than falling
through to the second reader, so a bad `-ColumnName` cannot masquerade as a
reader failure.

Python is no longer mandatory anywhere: install `ImportExcel` as a fallback, or
use the **`[2] Group`** source mode in the three scripts that offer it, which
reads members from Entra instead of a spreadsheet.

### 10.2 How the scripts find Python

`add_or_move`, `Bulk-Delete`, and `Remove-Entra-Licenses` share an identical
`Get-PythonExecutable` probe (see §10.2b for the dynamic replacement):

1. The DSH bundled interpreter — `%USERPROFILE%\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\python.exe` (listed twice: once as a hardcoded absolute path, once built from `$env:USERPROFILE`)
2. `python`, `python3`, `py` on `PATH` — **each tested with `import openpyxl`**

Only an interpreter that can `import openpyxl` is accepted. `bulk_create_user.ps1`
now uses the same shared resolver via `Get-PythonForOpenpyxl`, verifies
`openpyxl` before use, and falls back to `ImportExcel` when no suitable
interpreter exists.

> **No temp `.py` files.** `bulk_create_user.ps1` previously wrote its Python
> source to `%TEMP%` with `Set-Content` and then executed it. Because the script
> takes a `-WhatIf` switch, that write was suppressed in dry runs, so Python was
> handed a nonexistent path (`can't open file ...read_softnet_for_ps.py`). It now
> passes the code inline via `python -c`, so nothing is written to disk.

### 10.2b Dynamic resolution — `Resolve-PythonWithOpenpyxl.ps1`

The hardcoded path is not portable: it embeds the Windows user name
(`C:\Users\MAITA GOINGCO\...`), so it breaks on any other account or machine.
[Resolve-PythonWithOpenpyxl.ps1](Resolve-PythonWithOpenpyxl.ps1) is a drop-in
replacement that discovers Python at runtime. Tested under Windows PowerShell 5.1.

**Search order**

| # | Source | Catches |
| --- | --- | --- |
| 1 | `$env:ENTRA_PYTHON` | manual override / CI pinning |
| 2 | DSH bundled runtime, via `$env:USERPROFILE` | the current working setup |
| 3 | `py -0p` | all launcher-registered installs, incl. pyenv |
| 4 | Registry `PythonCore` (HKLM, WOW6432Node, HKCU) | installs never added to `PATH` |
| 5 | `python` / `python3` via `Get-Command -All` | ordinary PATH installs |

Each candidate must pass `importlib.util.find_spec('openpyxl')` before being
returned; otherwise the next one is tried.

**Two real traps it fixes**, both verified on this machine:

- **Store alias stubs.** `python3.exe` and `python.exe` in
  `AppData\Local\Microsoft\WindowsApps` are **0-byte reparse points**. `Test-Path`
  says they exist, but running them prints *"Python was not found"* and exits
  **9009**. A `Test-Path`-only check accepts them and fails later.
- **Hidden second PATH entry.** The plain `Get-Command python` returns only
  Laragon's 3.10 (no `openpyxl`) and stops. `-All` reveals the second entry, so
  a working interpreter further down `PATH` is still found.

**Usage**

```powershell
. "$PSScriptRoot\Resolve-PythonWithOpenpyxl.ps1"

$python = Resolve-PythonWithOpenpyxl
if (-not $python) {
    throw "No Python with openpyxl found. Install ImportExcel, or run: python -m pip install openpyxl"
}
```

`-SkipBundled` ignores the DSH runtime, for testing how a machine without DSH
behaves. `-NoCache` forces a fresh scan. Returns `$null` rather than throwing, so
callers keep control.

**Not a standalone script — it is a function library.** The file defines one
function and runs no main body. It must be **dot-sourced**; the file name looks
like a script, but running it directly with `&` or `. \Resolve-PythonWithOpenpyxl.ps1`
as a child process defines the function in a throwaway scope and discards it:

| How you load it | Works? | Why |
| --- | --- | --- |
| `. .\Resolve-PythonWithOpenpyxl.ps1` (dot-source) | **Yes** | function lands in the caller's scope |
| `& .\Resolve-PythonWithOpenpyxl.ps1` | No | child scope; function discarded afterwards |
| Rename to `.psm1` + `Import-Module` | Yes | module export; verified working |
| `powershell -File .\Resolve-...ps1` | No | function defined then process exits |

Dot-sourcing is verified: the function is visible afterwards, returns a
`System.String`, and leaks **no** variables into the caller. It does overwrite
`$LASTEXITCODE` (the final probe's exit code, normally `0`) — anything relying on
a previously-set `$LASTEXITCODE` should capture it before calling.

**Performance.** The probe uses `find_spec`, which merely locates the package
(~**82 ms**) instead of importing it (~**1407 ms**).

| Call | Cost |
| --- | --- |
| cold resolve | ~580 ms |
| **cached** repeat | **~1–7 ms** |
| `-SkipBundled` (uncached) | ~465 ms |
| `-NoCache` | ~250 ms |

Caching matters because the Excel helpers run twice per execution — once for
headers, once for data — so an uncached resolver would add ~0.5 s per run.
Successful lookups are cached per session in `$script:ResolvedPythonPath`;
`-SkipBundled` results are never cached and never overwrite it, so a test probe
cannot poison the normal path. A cached path that no longer exists is detected
and the scan re-runs automatically.

> Note: `Test-Path` alone is not a portability check — it is true for the 0-byte
> Store stubs. The `Length -eq 0` filter is what makes the PATH step safe.

**Wiring (applied).** All four functional scripts now use the shared resolver. Each
one dot-sources it near the top of its Setup block:

```powershell
$script:PyResolverPath = Join-Path $scriptRoot 'Resolve-PythonWithOpenpyxl.ps1'
if (Test-Path -LiteralPath $script:PyResolverPath) { . $script:PyResolverPath }
```

| Script | Change |
| --- | --- |
| `add_or_move_Entra-Users.ps1` | `Get-PythonExecutable` delegates to the helper |
| `Bulk-Delete-Entra-Users.ps1` | `Get-PythonExecutable` delegates to the helper |
| `Remove-Entra-Licenses.ps1` | `Get-PythonExecutable` delegates to the helper |
| `bulk_create_user.ps1` | Resolves `$Py` dynamically, plus a fail-fast `openpyxl` check |
| `menu.ps1` | **Unchanged** — no Python involvement |

Design points:

- **Backward compatible.** If `Resolve-PythonWithOpenpyxl.ps1` is deleted, each
  script falls back to its original inline probe (minus the hardcoded user name).
  Verified: deleting the helper still returns a working interpreter.
- **Delegation is all-or-nothing.** When the helper is present it is trusted; the
  scripts do not retry the old paths behind it. A half-broken install therefore
  fails consistently instead of working intermittently.
- **`bulk_create_user.ps1` fails fast.** Previously a missing `openpyxl` surfaced
  as a raw `ModuleNotFoundError` mid-loop. It now checks before processing and
  exits 1 with an actionable message.

Verified on this machine after wiring: all six files parse, all three
`Get-PythonExecutable` variants return the bundled interpreter, the fallback path
works with the helper removed, an Excel write→read round-trip through the resolved
interpreter succeeds, and `menu.ps1` still launches.

### 10.3 The trap on this machine

`python` on `PATH` resolves to Laragon's 3.10, which has **no `openpyxl`**:

```
python -> C:\laragon\bin\python\python-3.10\python.exe   (import openpyxl fails)
```

That is harmless for the three shared-probe scripts — they reject it at step 3
and pick the bundled 3.12 instead. But `bulk_create_user.ps1` only falls back to
bare `python` **when the bundled path is missing**, and then it will run
successfully and fail later with a raw `ModuleNotFoundError`. If you ever move
this folder to another machine, verify the interpreter first:

```powershell
python -c "import openpyxl; print(openpyxl.__version__)"
```

To install `openpyxl` into the PATH interpreter:

```powershell
python -m pip install openpyxl
```

### 10.4 Minimum Python baseline

- Python **3.8+**; 3.12.14 (bundled) and 3.10 (Laragon) both work
- `openpyxl` **3.x**; 3.1.5 present in the bundled interpreter
- No other third-party packages — the embedded scripts import only
  `openpyxl`, `sys`, `os`, `csv`, `json`, `re`
