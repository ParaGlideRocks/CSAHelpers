# Exchange Scripts

This directory contains PowerShell scripts for managing and auditing Microsoft Exchange Online (EXO) configurations, DNS records, and migration tasks.

## Script Guides

- [Dump-EXOBasicAuthReport](Dump-EXOBasicAuthReport/README.md)
- [Export-DNSRecords](Export-DNSRecords/README.md)
- [Export-DNSReport](Export-DNSReport/README.md)
- [Export-ExchangeConfig](Export-ExchangeConfig/README.md)
- [Remove-InvalidSMTP](Remove-InvalidSMTP/README.md)
- [Start-EXOMigrationBatch](Start-EXOMigrationBatch/README.md)

Each script now resides in its own directory. The guides above contain the current requirements, parameters, usage, outputs, and safety notes.

## Scripts

### 1. Export-DNSRecords.ps1

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

### 2. Export-DNSReport.ps1

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

### 3. Export-ExchangeConfig.ps1

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

### 4. Remove-InvalidSMTP.ps1

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

### 5. Start-EXOMigrationBatch.ps1

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
| Export-DNSRecords.ps1 | 1.2.0 | 2026-06-18 | Single nameserver query, per-record TTL |
| Export-DNSReport.ps1 | 1.2.0 | 2026-06-18 | SPF filtering, TTL columns |
| Export-ExchangeConfig.ps1 | 1.0.0 | - | Initial release |
| Remove-InvalidSMTP.ps1 | 1.0.0 | - | Initial release |
| Start-EXOMigrationBatch.ps1 | 1.0.0 | - | Initial release |

## Support

For issues or questions about these scripts, please review the comment-based help:

```powershell
Get-Help .\Export-DNSRecords\Export-DNSRecords.ps1 -Full
Get-Help .\Export-DNSReport\Export-DNSReport.ps1 -Full
```

## License

See [LICENSE](../LICENSE) file in the repository root.
