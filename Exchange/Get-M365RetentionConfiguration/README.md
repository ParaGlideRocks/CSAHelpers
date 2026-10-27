# Get-M365RetentionConfiguration

Documents Microsoft 365 retention and hold configuration in read-only mode.

## Requirements

- PowerShell 7.0 or later (7.4 or later recommended)
- ExchangeOnlineManagement 3.10.1 or later
- Exchange Online read permissions
- Microsoft Purview read permissions for compliance reporting

## Usage

```powershell
.\Get-M365RetentionConfiguration.ps1
.\Get-M365RetentionConfiguration.ps1 -UserPrincipalName admin@contoso.com
.\Get-M365RetentionConfiguration.ps1 -IncludeSoftDeletedMailboxes
.\Get-M365RetentionConfiguration.ps1 -SkipComplianceConnection
```

## Parameters

- `OutputPath`: Destination directory. Defaults to `.\M365-Retention-Assessment`.
- `UserPrincipalName`: Optional account used for Exchange Online and Microsoft Purview connections.
- `IncludeSoftDeletedMailboxes`: Includes soft-deleted mailboxes in the mailbox inventory.
- `SkipComplianceConnection`: Skips the Microsoft Purview connection and compliance collection.
- `ContinueOnPreflightFailure`: Continues with partial access when blocking preflight checks fail.

## Output

The script creates a timestamped folder containing CSV and JSON reports for:

- Runtime, cmdlet, and effective RBAC preflight
- Exchange organization configuration
- Organization-wide holds
- Purview retention policies and rules
- eDiscovery case hold policies
- App retention policies
- Exchange MRM policies and tags
- Mailbox retention and hold configuration
- Collection failures

The script does not change Microsoft 365 configuration. The generated reports can contain security-sensitive tenant and mailbox information, so store them in an approved location.
