# Export-ExchangeConfig

Exports on-premises Exchange configuration for auditing or backup and can generate a self-contained HTML report.

## Requirements

- Exchange 2013 or later
- Exchange Management Shell or remote PowerShell access
- Administrative permission to read the requested configuration

## Key Parameters

- `Servers`: Exchange server names to query.
- `Categories`: `All` or selected configuration categories.
- `OutputPath`: Destination directory.
- `Credential`: Optional remote PowerShell credential.
- `ExchangeUri`: Optional remote PowerShell endpoint.
- `GenerateHtmlReport`: Generates an HTML report in addition to JSON.
- `JsonInputFile`: Generates a report from an existing export without querying Exchange.

## Usage

```powershell
.\Export-ExchangeConfig.ps1 -Servers EX01,EX02 -Categories All -GenerateHtmlReport
.\Export-ExchangeConfig.ps1 -JsonInputFile .\ExchangeConfig.json -GenerateHtmlReport
```

## Output

The script writes a JSON configuration export, a timestamped log, and optionally an HTML report. Collection may take time in large Exchange environments.