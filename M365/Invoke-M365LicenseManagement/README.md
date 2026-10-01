# Invoke-M365LicenseManagement

Interactively manages Microsoft 365 license assignments through Microsoft Graph.

## Requirements

- PowerShell 7.0 or later
- Microsoft Graph PowerShell modules:
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.Identity.DirectoryManagement`
  - `Microsoft.Graph.Users`
  - `Microsoft.Graph.Users.Actions`
  - `Microsoft.Graph.Groups` when `PilotGroupId` is used
- Delegated Microsoft Graph permissions:
  - `User.Read.All`
  - `Organization.Read.All`
  - `LicenseAssignment.ReadWrite.All`
  - `GroupMember.Read.All` when `PilotGroupId` is used
- License Administrator or User Administrator role in Entra ID

The script checks for required Graph modules and installs missing modules for the current user.

## Parameters

- `OutputFolder`: Destination folder for CSV reports, JSON backups, and logs. Defaults to `.\LicenseManagementOutput`.
- `PilotUserCsv`: Optional CSV containing a `UserPrincipalName` column. Limits license replacement to listed users.
- `PilotGroupId`: Optional Entra ID group object ID. Limits license replacement to transitive group members.
- `DisabledPlansMode`: Controls target license service plans. `Preserve` carries matching disabled plans from the source license; `None` assigns the target license with all plans enabled.
- `IncludeDisabledAccounts`: Includes disabled accounts in replacement operations.
- `IncludeInactiveSkus`: Lists SKUs whose capability status is not `Enabled` or `Warning`.

## Usage

```powershell
.\Invoke-M365LicenseManagement.ps1
.\Invoke-M365LicenseManagement.ps1 -PilotUserCsv .\pilot.csv
.\Invoke-M365LicenseManagement.ps1 -PilotGroupId 00000000-0000-0000-0000-000000000000 -DisabledPlansMode Preserve
.\Invoke-M365LicenseManagement.ps1 -OutputFolder C:\Reports\M365Licenses -IncludeDisabledAccounts
```

## Menu Options

1. Show the licenses available in the tenant.
2. Export users holding a selected license.
3. Replace one license with another license.
4. Restore license assignments from a replacement backup.
5. Exit.

## Output

The script writes CSV reports, JSON rollback backups, and timestamped log files to `OutputFolder`.

License replacement always creates a pre-check CSV and a JSON backup before changes are made. The operator must type `EXECUTE` before any replacement is applied. Restore operations require typing `RESTORE` before the backup is replayed.

## Safety Notes

- Run against a pilot group or pilot CSV before broad tenant changes.
- Users whose source license is group-based are skipped because direct license removal would not remove the inherited assignment.
- Users without `UsageLocation` are skipped because Microsoft 365 licensing requires that value.
- Disabled accounts are skipped unless `IncludeDisabledAccounts` is specified.
- The replacement operation removes only the selected source SKU and preserves all other assigned SKUs.
