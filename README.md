# Entra Scripts

A set of Windows PowerShell scripts for managing Microsoft Entra ID (Azure AD) users in bulk: create users, move them between groups, remove licenses, and delete users. A small menu launcher ties them together.

## Overview

| Script | What it does |
| --- | --- |
| `menu.ps1` | Launcher. Lets you pick one of the scripts below to run. |
| `bulk_create_user.ps1` | Creates many users from an Excel file. Sets the usage location (default `PH`) and assigns licenses. |
| `add_or_move_Entra-Users.ps1` | Adds users from an Excel file to a group, or moves them from one group to another. |
| `Remove-Entra-Licenses.ps1` | Removes licenses from the users listed in an Excel file. |
| `Bulk-Delete-Entra-Users.ps1` | Deletes users and strips their licenses. Users come from an Excel file or from a group, and are always limited to a chosen group. |
| `Find-Duplicate-Entra-Names.ps1` | Read-only scan of all users for duplicate or near-duplicate names (for example "Jon Smith" and "John Smith"). Changes nothing. Saves a CSV report to the `Output Report` folder. |
| `Resolve-PythonWithOpenpyxl.ps1` | Shared helper. Finds a Python install that has `openpyxl`, used to read Excel files. Not run on its own. |

Scripts write CSV reports to the `Output Report` folder.

## Dependencies

### Required

- **Windows PowerShell 5.1 or newer** (built into Windows).
- **Microsoft Graph PowerShell** modules. Install the full set with:

  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser -Force
  ```

  The scripts use these parts of it:
  - `Microsoft.Graph.Authentication` (sign in with `Connect-MgGraph`)
  - `Microsoft.Graph.Users`
  - `Microsoft.Graph.Users.Actions` (license assignment)
  - `Microsoft.Graph.Groups`
  - `Microsoft.Graph.Identity.DirectoryManagement` (license list)

- **Python 3 with `openpyxl`**, used to read `.xlsx` files:

  ```powershell
  python -m pip install openpyxl
  ```

### Optional

- **ImportExcel** module. Used as a fallback to read Excel files if Python/`openpyxl` is not available:

  ```powershell
  Install-Module ImportExcel -Scope CurrentUser -Force
  ```

- **AzureAD** module. Legacy and deprecated. Not needed, because the scripts use Microsoft Graph.

### Permissions

The scripts sign in with `Connect-MgGraph` and ask for these permissions (an admin must consent to them in Entra):

- `User.Read.All`, `User.ReadWrite.All`
- `Group.Read.All`, `GroupMember.Read.All`, `GroupMember.ReadWrite.All`
- `Directory.Read.All`
- `Organization.Read.All`

## Getting started

1. Install the dependencies above.
2. Open PowerShell in this folder.
3. Run the launcher:

   ```powershell
   .\menu.ps1
   ```

   Or run a script directly, for example `.\bulk_create_user.ps1`.
