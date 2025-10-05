# UserPasswordExpiration.ps1

A small interactive PowerShell utility to view and manage Microsoft Entra (Azure AD) password expiration settings for your tenant and users.

## What this script does

- View and change the tenant (domain) password expiration settings (PasswordValidityPeriodInDays and PasswordNotificationWindowInDays).
- Export a report of users' password expiration status to a CSV file (Reports folder).
- Update individual users' password expiration behavior by setting or clearing the DisablePasswordExpiration flag via a CSV file (Updates folder).

> Note: There is no UI in the Entra admin portal to view or set the domain-level password expiration options; this script uses the Microsoft Graph PowerShell SDK.

## Location

Place and run the script from the folder where `UserPasswordExpiration.ps1` resides. The script will create two subfolders as needed:

- `Reports` — CSV reports are saved here. Filename pattern: `UserPasswordExpiration_report_YYYY-MM-DD_HH-mm-ss.csv`.
- `Updates` — Place your update CSV files here (or edit the template the script creates).
- `Logs` — Script transcript saved here with timestamped log files.

## Prerequisites

- PowerShell (Windows PowerShell 5.1 or PowerShell 7+).
- Microsoft Graph PowerShell SDK installed and available: `Install-Module Microsoft.Graph -Scope CurrentUser`.
If `Connect-MgGraph` is not available, install the Microsoft.Graph module and re-open PowerShell.

## Execution policy

If execution of scripts is blocked, run PowerShell as Administrator and set an execution policy for the session or machine, for example:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## How to run

1. Open PowerShell (recommended: Run as Administrator if you need to change execution policy).
2. (Optional) Install Microsoft Graph module:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

3. Run the script:

```powershell
cd 'C:\Path\To\UserPasswordExpiration'
.\UserPasswordExpiration.ps1
```

4. Follow the interactive menu:
- Org settings — view or update tenant/domain password expiry and notification window.
- Report user password expiration — produces a CSV report of all users and their password expiration status.
- Update user password expiration — uses a CSV in the `Updates` folder to set or clear `DisablePasswordExpiration` per user.

## Update CSV format (for "Update user password expiration")

If no CSV file is found in the `Updates` folder, the script creates a template file named like `UserPasswordExpiration_update_YYYY-MM-DD_HH-mm-ss.csv` with the following content:

```
UserPrincipalName,DisablePasswordExpiration
jsmith@domain.com,TRUE
```

- Column names: `UserPrincipalName` and `DisablePasswordExpiration`.
- Allowed values for `DisablePasswordExpiration`: `TRUE` (to disable expiration) or any other value (to leave/clear it).
- Place your CSV in the `Updates` folder or edit the template the script creates. The script will pick the most recent CSV in that folder.

## Important behavior notes

- A domain value of `2147483647` (or a stored `null`) means the domain does NOT enforce password expiration (treated as "Never").
- If a user is about to hit their password change date and you set `DisablePasswordExpiration=TRUE`, they will NOT be prompted to change their password at next login.
- When changing domain-level settings, the script attempts to call `Update-MgDomain` and requires `Domain.ReadWrite.All` / `Directory.AccessAsUser.All` scopes.

## Output files

- Reports are saved to the `Reports` folder as CSV and openable in Excel.
- A transcript of the entire PowerShell session is created and moved to the `Logs` folder on exit.

## Troubleshooting

- "Connect-MgGraph is NOT available": install the Microsoft.Graph module, then re-open PowerShell and import or run `Connect-MgGraph`.
- Authentication/Permissions: ensure the account used to `Connect-MgGraph` has the required admin privileges and consent for the requested scopes.
- If `Update-MgUser`/`Update-MgDomain` calls fail, check that the Microsoft Graph module is up-to-date and that your signed-in account has write permissions.

## Safety and testing

- Test the report option first to confirm tenant domains and sample users.
- When using the update CSV, try a single user and "Yes" prompts rather than "Yes to All" until you're confident.

## Notes

- The script is interactive and does not accept command-line parameters; it relies on prompts, menus, and the Update CSV for bulk changes.
- There is a small helper module `ITAutomator.psm1` that the script expects to be present in the same folder. The script will error and exit with code 99 if that module is not found.

## License and contact

Use and modify this script within your organization as needed. If you need help adapting it, include the script and details of the change you want.

