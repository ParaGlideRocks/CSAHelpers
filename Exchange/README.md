# Exchange Scripts

This directory contains PowerShell scripts for managing and auditing Microsoft Exchange Online (EXO) configurations, DNS records, email authentication, and migration tasks.

## Script Guides

- [Analyze-DmarcReports](Analyze-DmarcReports/README.md)
- [Export-DmarcAttachments](Export-DmarcAttachments/README.md)
- [Export-DNSRecords](Export-DNSRecords/README.md)
- [Export-DNSReport](Export-DNSReport/README.md)
- [Export-ExchangeConfig](Export-ExchangeConfig/README.md)
- [Get-EXOBasicAuthReport](Get-EXOBasicAuthReport/README.md)
- [Remove-InvalidSMTP](Remove-InvalidSMTP/README.md)
- [Start-EXOMigrationBatch](Start-EXOMigrationBatch/README.md)

Each script now resides in its own directory. The guides above contain the current requirements, parameters, usage, outputs, and safety notes.

## Scripts

### 1. Analyze-DmarcReports.ps1

Parses DMARC aggregate (RUA) reports and summarises authentication results in an Excel workbook.

**Features:**
- Expands `.gz` and `.zip` attachments automatically
- Aggregates the sources that fail DMARC by IP address
- Reports SPF and DKIM alignment, DKIM selector, and policy override reasons
- Traces every row back to its originating report file
- Falls back to CSV output when the `ImportExcel` module is unavailable

**Parameters:**
- `-Path` (Optional): Folder containing the reports (defaults to the current directory)
- `-OutFile` (Optional): Output XLSX file path (defaults to a timestamped file in the source folder)
- `-ResolveHostnames` (Optional): Resolves reverse DNS for the failing sources
- `-SkipDettaglio` (Optional): Limits the detail worksheet to failing records
- `-ReportId` (Optional): Processes only reports matching the given report id

**Example:**
```powershell
.\Analyze-DmarcReports\Analyze-DmarcReports.ps1 -Path C:\DMARC\raw
.\Analyze-DmarcReports\Analyze-DmarcReports.ps1 -Path C:\DMARC\raw -SkipDettaglio -ResolveHostnames
```

**Report Sections:**
- Riepilogo (per report: reporter, period, published policy, volumes, pass and fail counts)
- DaVerificare (failing sources grouped by IP, ordered by volume)
- Dettaglio (per record: source IP, disposition, SPF and DKIM alignment, sender identifiers)

---

### 2. Export-DmarcAttachments.ps1

Saves DMARC aggregate report attachments from an Outlook folder to disk.

**Features:**
- Exports `.gz`, `.zip`, and `.xml` attachments from a mail folder
- Restricts messages by date through MAPI, so large RUA mailboxes stay workable
- Prefixes every file with the message timestamp to prevent name collisions
- Skips files already on disk, allowing incremental re-runs
- Optionally moves the processed messages to another folder
- Detects an elevated session before doing any work

**Parameters:**
- `-FolderPath` (Optional): Folder path from the mailbox root (prompts when omitted)
- `-OutputPath` (Optional): Destination directory (defaults to `.\DmarcReports`)
- `-Since` / `-Until` (Optional): Restricts the messages by received date
- `-Extensions` (Optional): Attachment extensions to export (defaults to `gz`, `zip`, `xml`)
- `-MoveToFolder` (Optional): Existing folder the exported messages are moved to
- `-IncludeSubfolders` (Optional): Also processes subfolders
- `-Overwrite` (Optional): Re-exports attachments already present

**Example:**
```powershell
.\Export-DmarcAttachments\Export-DmarcAttachments.ps1
.\Export-DmarcAttachments\Export-DmarcAttachments.ps1 -FolderPath "dmarc_rua@contoso.com\Inbox" -OutputPath C:\DMARC\raw
```

**Note:** Requires Outlook classic running at the same integrity level as the script. An elevated session fails with `CO_E_SERVER_EXEC_FAILURE` (0x80080005).

---

### 3. Export-DNSRecords.ps1

Exports MX, SPF, and DMARC DNS records for a list of domains from a CSV file.

**Features:**
- Queries authoritative nameservers for DNS records
- Tracks DNS record TTL (Time To Live) values
- Measures DNS query performance (milliseconds)
- Generates timestamp-prefixed JSON output
- Includes per-domain nameserver information

**Parameters:**
- `-CsvPath` (Mandatory): Path to input CSV file with "Domain" column
- `-OutputPath` (Optional): Output JSON file path (defaults to `yyyyMMdd-HHmmss_DNSRecordsExport.json`)
- `-DnsServer` (Optional): Custom DNS server for queries

**Example:**
```powershell
.\Export-DNSRecords\Export-DNSRecords.ps1 -CsvPath .\domains.csv
.\Export-DNSRecords\Export-DNSRecords.ps1 -CsvPath .\domains.csv -OutputPath .\dns_export.json -DnsServer 8.8.8.8
```

**Output Schema:**
```json
{
  "Domain": "example.com",
  "QueriedNameServer": "ns1.example.com",
  "QueryTimingMs": {
    "MX": 150,
    "TXT": 200,
    "DMARC": 100,
    "Total": 450
  },
  "MX": [
    {
      "Preference": 10,
      "Exchange": "mail.example.com",
      "TTL": 3600,
      "QueryDurationMs": 50
    }
  ],
  "TXT": [
    {
      "Value": "v=spf1 include:example.com ~all",
      "TTL": 3600,
      "QueryDurationMs": 75
    }
  ],
  "DMARC": [
    {
      "Value": "v=DMARC1; p=reject; ...",
      "TTL": 3600,
      "QueryDurationMs": 50
    }
  ]
}
```

---

### 4. Export-DNSReport.ps1

Generates a styled, searchable HTML report from DNS records exported by `Export-DNSRecords.ps1`.

**Features:**
- Displays MX records with preferences
- Shows SPF records filtered from TXT records
- Includes DMARC policy status with color-coded badges
- Shows TTL values per record
- Searchable and filterable interface
- Displays which nameserver was queried

**Parameters:**
- `-JsonPath` (Mandatory): Path to JSON file from `Export-DNSRecords.ps1`
- `-OutputPath` (Optional): Output HTML file path (defaults to `DNSReport.html`)

**Example:**
```powershell
.\Export-DNSReport\Export-DNSReport.ps1 -JsonPath .\20260618-185157_DNSRecordsExport.json
.\Export-DNSReport\Export-DNSReport.ps1 -JsonPath .\dns_export.json -OutputPath .\report.html
```

**Report Sections:**
- Summary statistics (Total domains, DMARC policy breakdown, SPF adoption)
- Searchable domain table with filters
- MX Records with TTL
- SPF Records with TTL
- DMARC Policy with TTL
- Status badges (SPF present, DMARC policy level)

---

### 5. Export-ExchangeConfig.ps1

Exports Exchange Online configuration and mailbox settings for auditing and backup purposes.

**Features:**
- Exports Exchange organization configuration
- Captures mailbox properties and settings
- Records recipient configurations
- Generates timestamped JSON output

**Parameters:**
- `-OutputPath` (Optional): Output JSON file path

**Example:**
```powershell
.\Export-ExchangeConfig\Export-ExchangeConfig.ps1 -Servers EX01
.\Export-ExchangeConfig\Export-ExchangeConfig.ps1 -Servers EX01 -GenerateHtmlReport
```

---

### 6. Get-EXOBasicAuthReport.ps1

Audits Exchange Online authentication policies and reports users for whom basic authentication may be enabled.

**Features:**
- Reports organization-level authentication settings
- Enumerates all authentication policies and their configurations
- Identifies users whose policy leaves basic auth enabled for any protocol
- Reports the default authentication policy assignment
- Produces summary statistics and timestamped output files

**Parameters:**
- `-OutputPath` (Optional): Destination directory (defaults to the current directory)
- `-ExportFormat` (Optional): `CSV`, `JSON`, or `XML` (defaults to `CSV`)
- `-IncludeDetailedUserInfo` (Optional): Includes additional user details in the report

**Example:**
```powershell
Connect-ExchangeOnline
.\Get-EXOBasicAuthReport\Get-EXOBasicAuthReport.ps1 -OutputPath C:\Reports
.\Get-EXOBasicAuthReport\Get-EXOBasicAuthReport.ps1 -OutputPath C:\Reports -ExportFormat JSON -IncludeDetailedUserInfo
```

**Note:** The script is read-only, but its output may contain security-sensitive tenant information.

---

### 7. Remove-InvalidSMTP.ps1

Removes invalid or duplicate SMTP addresses from Exchange Online mailboxes.

**Features:**
- Identifies invalid SMTP formats
- Removes duplicate email aliases
- Validates SMTP syntax before removal
- Reports changes made

**Parameters:**
- TBD (review script for complete documentation)

**Example:**
```powershell
.\Remove-InvalidSMTP\Remove-InvalidSMTP.ps1 -UserFile C:\Data\Users.txt -InvalidDomains olddomain.com -WhatIf
```

---

### 8. Start-EXOMigrationBatch.ps1

Manages and initiates Exchange Online migration batches.

**Features:**
- Creates migration batches from CSV source
- Monitors migration progress
- Handles batch status and completion

**Parameters:**
- TBD (review script for complete documentation)

**Example:**
```powershell
.\Start-EXOMigrationBatch\Start-EXOMigrationBatch.ps1 -UserFile C:\Migration\Users.txt -MigrationEndpoint OnpremEndpoint -TargetDeliveryDomain contoso.mail.onmicrosoft.com -WhatIf
```

---

## Workflow Examples

### Example 1: Complete DNS Audit Report

```powershell
# Step 1: Export DNS records for your domains
$domains = @"
Domain
contoso.com
fabrikam.com
northwindtraders.com
"@

$domains | Set-Content -Path domains.csv

.\Export-DNSRecords\Export-DNSRecords.ps1 -CsvPath domains.csv

# Step 2: Generate HTML report
$jsonFile = Get-ChildItem -Filter "*_DNSRecordsExport.json" | Select-Object -First 1 -ExpandProperty FullName

.\Export-DNSReport\Export-DNSReport.ps1 -JsonPath $jsonFile

# Step 3: Open report in browser
Invoke-Item .\DNSReport.html
```

### Example 2: DNS Audit with Custom Nameserver

```powershell
# Query specific DNS server (useful for validating zone updates)
.\Export-DNSRecords\Export-DNSRecords.ps1 -CsvPath domains.csv -DnsServer ns1.yourdomain.com
```

### Example 3: DMARC Enforcement Readiness

```powershell
# Step 1: Confirm the policy currently published for each domain
.\Export-DNSRecords\Export-DNSRecords.ps1 -CsvPath domains.csv

# Step 2: Export the aggregate reports from the RUA mailbox
.\Export-DmarcAttachments\Export-DmarcAttachments.ps1 -FolderPath "dmarc_rua@contoso.com\Inbox" -OutputPath C:\DMARC\raw

# Step 3: Analyse the exported reports
.\Analyze-DmarcReports\Analyze-DmarcReports.ps1 -Path C:\DMARC\raw -SkipDettaglio -ResolveHostnames

# Step 4: Review the DaVerificare worksheet and align any legitimate sender
# before moving a domain from p=quarantine to p=reject
```

---

## Prerequisites

- **PowerShell 5.0+**
- **Exchange Online PowerShell Module** (for Exchange-specific scripts)
- **Global Administrator or Exchange Administrator role** in Exchange Online
- Network access to DNS servers and Exchange Online endpoints

## Installation

1. Clone or download the scripts to your local machine
2. Set execution policy (if needed):
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
3. Connect to Exchange Online (if using Exchange scripts):
   ```powershell
   Connect-ExchangeOnline
   ```

## Version History

| Script | Version | Last Updated | Changes |
|--------|---------|--------------|---------|
| Analyze-DmarcReports.ps1 | 1.2.0 | 2026-09-21 | Report id filter, originating file columns |
| Export-DmarcAttachments.ps1 | 1.0.0 | 2026-09-21 | Initial release |
| Export-DNSRecords.ps1 | 1.2.0 | 2026-06-18 | Single nameserver query, per-record TTL |
| Export-DNSReport.ps1 | 1.2.0 | 2026-06-18 | SPF filtering, TTL columns |
| Export-ExchangeConfig.ps1 | 3.0.0 | 2026-09-21 | HTML report generation, report-only mode, horizontal scrolling fix |
| Get-EXOBasicAuthReport.ps1 | 1.0.0 | 2026-09-21 | Initial release |
| Remove-InvalidSMTP.ps1 | 1.0.0 | - | Initial release |
| Start-EXOMigrationBatch.ps1 | 1.0.0 | - | Initial release |

## Support

For issues or questions about these scripts, please review the comment-based help:

```powershell
Get-Help .\Analyze-DmarcReports\Analyze-DmarcReports.ps1 -Full
Get-Help .\Export-DNSRecords\Export-DNSRecords.ps1 -Full
Get-Help .\Export-DNSReport\Export-DNSReport.ps1 -Full
```

## License

See [LICENSE](../LICENSE) file in the repository root.
