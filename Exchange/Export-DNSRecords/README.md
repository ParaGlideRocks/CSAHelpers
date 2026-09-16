# Export-DNSRecords

Queries DNS records for domains listed in a CSV file and exports the results to JSON.

## Requirements

- PowerShell 5.1 or later
- Internet access and the `Resolve-DnsName` cmdlet

## Parameters

- `CsvPath`: Required CSV path. The file must contain a `Domain` column.
- `OutputPath`: Destination JSON path. Defaults to a timestamped file beside the input CSV.
- `DnsServer`: Optional resolver used to discover authoritative name servers.

## Usage

```powershell
.\Export-DNSRecords.ps1 -CsvPath .\domains.csv
.\Export-DNSRecords.ps1 -CsvPath .\domains.csv -OutputPath .\dns.json -DnsServer 8.8.8.8
```

## Output

The JSON includes MX, TXT, SPF, DMARC, and DKIM data, record TTLs, queried servers, and query timing information.

## Related Script

Use `Export-DNSReport.ps1` to turn the generated JSON into an HTML report.