# Analyze-DmarcReports

Parses DMARC aggregate (RUA) reports from a folder and summarises authentication results in an Excel workbook.

## Requirements

- PowerShell 5.1 or later
- A folder of DMARC aggregate reports (`.xml`, `.gz`, or `.zip`)
- `ImportExcel` module for XLSX output; without it the script writes CSV files instead

## Parameters

- `Path`: Source folder containing the reports. Defaults to the current directory.
- `OutFile`: Destination XLSX path. Defaults to a timestamped file inside the source folder.
- `ResolveHostnames`: Performs reverse DNS lookups on the sources that fail DMARC.
- `SkipDettaglio`: Limits the detail worksheet to failing records only.
- `ReportId`: Processes only the reports whose `report_id` matches the value, in whole or in part.

## Usage

```powershell
.\Analyze-DmarcReports.ps1 -Path C:\DMARC\raw
.\Analyze-DmarcReports.ps1 -Path C:\DMARC\raw -SkipDettaglio -ResolveHostnames
.\Analyze-DmarcReports.ps1 -Path C:\DMARC\raw -ReportId 1789520405
```

## Output

The workbook contains three worksheets:

- `Riepilogo`: one row per report, with reporter, reporting period, published policy, message volume, pass and fail counts, and quarantined or rejected totals.
- `DaVerificare`: sources that fail DMARC, grouped by IP address and ordered by volume. This is the view used to decide whether a domain is ready for `p=reject`.
- `Dettaglio`: one row per record, with source IP, evaluated disposition, SPF and DKIM alignment, DKIM selector, header and envelope sender, and any policy override reason.

Every row also carries the originating file name and full path, so a finding can be traced back to the attachment it came from. A summary is printed to the console at the end of the run, including the ten highest-volume failing sources.

Optional elements of the DMARC schema, such as `selector`, `sp`, `pct`, and `reason`, are read leniently: reports that omit them are still processed. Malformed files are skipped and listed once processing completes.

## Notes

The script is read-only. It does not require Outlook, an Exchange session, or administrative privileges, so it can be run on any workstation holding a copy of the reports. A DMARC failure does not by itself indicate spoofing: legitimate senders that are not yet aligned must be corrected before a domain is moved to `p=reject`.

## Related Script

Use `Export-DNSRecords.ps1` and `Export-DNSReport.ps1` to audit the published SPF, DKIM, and DMARC records for the same domains.
