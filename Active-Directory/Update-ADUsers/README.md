# Update-ADUsers

Bulk updates Active Directory user attributes from an Excel worksheet.

## Requirements

- Active Directory PowerShell module (RSAT)
- `ImportExcel` PowerShell module
- Permission to modify the target users

## Parameters

- `ExcelPath`: Required path to the source workbook.
- `WorksheetName`: Worksheet containing user data. Defaults to `Users`.
- `LogPath`: Destination log path. Defaults to a timestamped file.
- `StopOnError`: Stops processing after the first failed user update.
- Common parameter `WhatIf`: Previews supported changes without applying them.

## Input

Each row must identify a user by `SamAccountName` or `UserPrincipalName`. Supported columns include profile fields, manager, enabled state, and extension attributes. Use `__CLEAR__` to clear a supported attribute.

## Usage

```powershell
.\Update-ADUsers.ps1 -ExcelPath C:\Data\Users.xlsx -WhatIf
.\Update-ADUsers.ps1 -ExcelPath C:\Data\Users.xlsx -WorksheetName Users
```

## Warning

This script modifies Active Directory. Review the workbook and run with `-WhatIf` first. Updates can disable email address policy processing for affected users.