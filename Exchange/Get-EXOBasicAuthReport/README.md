# Get-EXOBasicAuthReport

Audits Exchange Online authentication policies and reports users for whom basic authentication may be enabled.

## Requirements

- Exchange Online PowerShell module
- An active `Connect-ExchangeOnline` session
- Exchange Administrator or Global Reader permissions

## Parameters

- `OutputPath`: Destination directory. Defaults to the current directory.
- `ExportFormat`: `CSV`, `JSON`, or `XML`. Defaults to `CSV`.
- `IncludeDetailedUserInfo`: Includes additional user properties in the report.

## Usage

```powershell
Connect-ExchangeOnline
.\Get-EXOBasicAuthReport.ps1 -OutputPath C:\Reports
.\Get-EXOBasicAuthReport.ps1 -OutputPath C:\Reports -ExportFormat JSON -IncludeDetailedUserInfo
```

## Output

The script exports organization configuration, authentication policies, affected users, and a summary. It is read-only, but output may contain security-sensitive tenant information.