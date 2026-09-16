# Export-DNSReport

Generates a searchable HTML report from JSON produced by `Export-DNSRecords.ps1`.

## Requirements

- PowerShell 5.0 or later
- A valid `Export-DNSRecords.ps1` JSON result

## Parameters

- `JsonPath`: Required source JSON path.
- `OutputPath`: Destination HTML path. Defaults to a name derived from the source file.

## Usage

```powershell
.\Export-DNSReport.ps1 -JsonPath .\20260916-120000_DNSRecordsExport.json
.\Export-DNSReport.ps1 -JsonPath .\dns.json -OutputPath .\DNSReport.html
```

## Output

The self-contained HTML report includes summary statistics, searchable domain details, and status indicators for MX, SPF, DMARC, and DKIM records. The script opens the report in the default browser.