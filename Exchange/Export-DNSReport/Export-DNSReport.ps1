<#
.SYNOPSIS
    Generates an HTML report from a DNS records JSON export.

.DESCRIPTION
    Reads the JSON file produced by Export-DNSRecords.ps1 and generates a
    styled, searchable HTML report showing MX, TXT, DMARC, and DKIM records per domain.

.PARAMETER JsonPath
    Path to the JSON file produced by Export-DNSRecords.ps1.

.PARAMETER OutputPath
    Path for the HTML output file. Defaults to a matching filename derived from the JSON
    (e.g., "yyyyMMdd-HHmmss_DNSRecordsReport.html" for "yyyyMMdd-HHmmss_DNSRecordsExport.json").

.EXAMPLE
    .\Export-DNSReport.ps1 -JsonPath .\DNSRecordsExport.json

.EXAMPLE
    .\Export-DNSReport.ps1 -JsonPath .\DNSRecordsExport.json -OutputPath .\report.html
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$JsonPath,

    [Parameter()]
    [string]$OutputPath
)

if (-not $OutputPath) {
    $jsonDir    = Split-Path -Parent (Resolve-Path $JsonPath)
    $jsonFile   = Split-Path -Leaf (Resolve-Path $JsonPath)
    $htmlFile   = $jsonFile -replace 'DNSRecordsExport\.json$', 'DNSRecordsReport.html'
    if ($htmlFile -eq $jsonFile) {
        # Fallback if pattern doesn't match
        $htmlFile = (Get-Item -Path $JsonPath).BaseName + '_Report.html'
    }
    $OutputPath = Join-Path $jsonDir $htmlFile
}

$records = Get-Content -Path $JsonPath -Raw | ConvertFrom-Json

$generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$totalDomains = $records.Count

function Get-DmarcPolicy {
    param ($DmarcValues)
    foreach ($item in @($DmarcValues)) {
        $v = if ($item -is [string]) { $item } else { $item.Value }
        if ($v -match 'p=(\w+)') { return $matches[1].ToLower() }
    }
    return $null
}

function Get-SpfRecord {
    param ($TxtValues)
    foreach ($item in @($TxtValues)) {
        $v = if ($item -is [string]) { $item } else { $item.Value }
        if ($v -match '^v=spf1') { return $v }
    }
    return $null
}

function Get-DkimSelectors {
    param ($DkimRecords)
    return @($DkimRecords | ForEach-Object { if ($_ -is [string]) { $_ } else { $_.Selector } } | Sort-Object -Unique)
}

# Build domain rows HTML
$domainRows = foreach ($rec in $records) {
    $domain = $rec.Domain

    # MX rows
    $mxHtml = if ($rec.MX -and $rec.MX.Count -gt 0) {
        $rows = ($rec.MX | Sort-Object Preference, NameServer | ForEach-Object {
            $ns  = if ($_.NameServer) { " via $($_.NameServer)" } else { '' }
            "<tr><td class='pref'>$($_.Preference)</td><td>$($_.Exchange)$ns</td></tr>"
        }) -join ''
        "<table class='inner-table'><thead><tr><th>Pref</th><th>Exchange</th></tr></thead><tbody>$rows</tbody></table>"
    } else { "<span class='missing'>No MX records</span>" }
    
    # MX TTLs
    $mxTtls = @($rec.MX | Where-Object { $_.TTL } | ForEach-Object { $_.TTL } | Sort-Object -Unique) -join ', '
    $mxTtlHtml = if ($mxTtls) { $mxTtls } else { "<span class='muted'>-</span>" }

    # SPF records (TXT with v=spf1 only)
    $spf  = Get-SpfRecord -TxtValues $rec.TXT
    $spfHtml = if ($rec.TXT -and $rec.TXT.Count -gt 0) {
        $spfRecords = @($rec.TXT | ForEach-Object {
            $value = if ($_ -is [string]) { $_ } else { $_.Value }
            if ($value -match '^v=spf1') {
                "<tr><td>$([System.Web.HttpUtility]::HtmlEncode($value))</td></tr>"
            }
        }) -join ''
        
        if ($spfRecords) {
            "<table class='inner-table txt-table'><thead><tr><th>Value</th></tr></thead><tbody>$spfRecords</tbody></table>"
        } else {
            "<span class='missing'>No SPF record</span>"
        }
    } else { "<span class='missing'>No SPF record</span>" }
    
    # SPF TTLs
    $spfTtls = @($rec.TXT | Where-Object { $_ -isnot [string] -and $_.Value -match '^v=spf1' } | ForEach-Object { $_.TTL } | Sort-Object -Unique) -join ', '
    $spfTtlHtml = if ($spfTtls) { $spfTtls } else { "<span class='muted'>-</span>" }

    # DMARC
    $dmarcPolicy = Get-DmarcPolicy -DmarcValues $rec.DMARC
    $dmarcHtml = if ($rec.DMARC -and $rec.DMARC.Count -gt 0) {
        $badgeClass = switch ($dmarcPolicy) {
            'reject'     { 'badge-reject' }
            'quarantine' { 'badge-quarantine' }
            'none'       { 'badge-none' }
            default      { 'badge-unknown' }
        }
        $dmarcFirstItem = $rec.DMARC[0]
        $dmarcValue = if ($dmarcFirstItem -is [string]) { $dmarcFirstItem } else { $dmarcFirstItem.Value }
        $encoded = [System.Web.HttpUtility]::HtmlEncode($dmarcValue)
        "<span class='badge $badgeClass'>$dmarcPolicy</span><br><span class='dmarc-value'>$encoded</span>"
    } else { "<span class='badge badge-warn'>No DMARC record</span>" }
    
    # DMARC TTLs
    $dmarcTtls = @($rec.DMARC | Where-Object { $_.TTL } | ForEach-Object { $_.TTL } | Sort-Object -Unique) -join ', '
    $dmarcTtlHtml = if ($dmarcTtls) { $dmarcTtls } else { "<span class='muted'>-</span>" }

    # DKIM records (CNAME)
    $dkimSelectors = Get-DkimSelectors -DkimRecords $rec.DKIM
    $dkimHtml = if ($rec.DKIM -and $rec.DKIM.Count -gt 0) {
        $rows = ($rec.DKIM | ForEach-Object {
            $selector = if ($_ -is [string]) { $_ } else { $_.Selector }
            $value = if ($_ -is [string]) { '' } else { $_.Value }
            "<tr><td class='dkim-sel'>$selector</td><td class='dkim-val'>$([System.Web.HttpUtility]::HtmlEncode($value))</td></tr>"
        }) -join ''
        "<table class='inner-table txt-table'><thead><tr><th>Selector</th><th>CNAME Target</th></tr></thead><tbody>$rows</tbody></table>"
    } else { "<span class='missing'>No DKIM records</span>" }
    
    # DKIM TTLs
    $dkimTtls = @($rec.DKIM | Where-Object { $_.TTL } | ForEach-Object { $_.TTL } | Sort-Object -Unique) -join ', '
    $dkimTtlHtml = if ($dkimTtls) { $dkimTtls } else { "<span class='muted'>-</span>" }

    # SPF / DMARC status badges for summary column
    $spfBadge   = if ($spf)        { "<span class='badge badge-ok'>SPF</span>" }  else { "<span class='badge badge-warn'>No SPF</span>" }
    $dmarcBadge = if ($dmarcPolicy -eq 'reject')     { "<span class='badge badge-reject'>DMARC: reject</span>" }
                  elseif ($dmarcPolicy -eq 'quarantine') { "<span class='badge badge-quarantine'>DMARC: quarantine</span>" }
                  elseif ($dmarcPolicy -eq 'none')    { "<span class='badge badge-none'>DMARC: none</span>" }
                  else                                { "<span class='badge badge-warn'>No DMARC</span>" }
    $dkimBadge  = if ($dkimSelectors.Count -gt 0) { "<span class='badge badge-ok'>DKIM: $($dkimSelectors.Count)</span>" }
                  else                              { "<span class='badge badge-warn'>No DKIM</span>" }

    @"
    <tr class="domain-row">
      <td class="domain-name">$domain</td>
      <td class="ns-cell">$($rec.QueriedNameServer)</td>
      <td>$mxHtml</td>
      <td class="ttl-cell">$mxTtlHtml</td>
      <td>$spfHtml</td>
      <td class="ttl-cell">$spfTtlHtml</td>
      <td>$dmarcHtml</td>
      <td class="ttl-cell">$dmarcTtlHtml</td>
      <td>$dkimHtml</td>
      <td class="ttl-cell">$dkimTtlHtml</td>
      <td class="status-cell">$spfBadge $dmarcBadge $dkimBadge</td>
    </tr>
"@
}

# Summary counters
$noSpf   = ($records | Where-Object { -not (Get-SpfRecord @($_.TXT)) }).Count
$noDmarc = ($records | Where-Object { -not ($_.DMARC -and $_.DMARC.Count -gt 0) }).Count
$dmarcNone       = ($records | Where-Object { (Get-DmarcPolicy @($_.DMARC)) -eq 'none' }).Count
$dmarcQuarantine = ($records | Where-Object { (Get-DmarcPolicy @($_.DMARC)) -eq 'quarantine' }).Count
$dmarcReject     = ($records | Where-Object { (Get-DmarcPolicy @($_.DMARC)) -eq 'reject' }).Count
$noDkim  = ($records | Where-Object { -not ($_.DKIM -and $_.DKIM.Count -gt 0) }).Count

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>DNS Records Report</title>
  <style>
    :root {
      --bg: #f4f6f9;
      --card: #ffffff;
      --primary: #0078d4;
      --text: #1a1a2e;
      --muted: #6c757d;
      --border: #dee2e6;
      --reject: #c0392b;
      --quarantine: #e67e22;
      --none-color: #7f8c8d;
      --ok: #27ae60;
      --warn: #e74c3c;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', Arial, sans-serif; background: var(--bg); color: var(--text); font-size: 13px; }

    header { background: var(--primary); color: #fff; padding: 18px 32px; display: flex; align-items: center; justify-content: space-between; }
    header h1 { font-size: 1.4rem; font-weight: 600; }
    header .meta { font-size: 0.8rem; opacity: 0.85; }

    .summary { display: flex; gap: 14px; padding: 18px 32px; flex-wrap: wrap; }
    .stat-card { background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 14px 20px; min-width: 130px; text-align: center; box-shadow: 0 1px 3px rgba(0,0,0,.06); }
    .stat-card .num { font-size: 1.8rem; font-weight: 700; color: var(--primary); }
    .stat-card .num.warn { color: var(--warn); }
    .stat-card .num.ok   { color: var(--ok); }
    .stat-card .num.rej  { color: var(--reject); }
    .stat-card .num.qua  { color: var(--quarantine); }
    .stat-card .lbl { font-size: 0.75rem; color: var(--muted); margin-top: 4px; }

    .controls { padding: 0 32px 14px; display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
    .controls input[type=text] {
      padding: 7px 12px; border: 1px solid var(--border); border-radius: 6px;
      font-size: 13px; width: 280px; outline: none;
    }
    .controls input[type=text]:focus { border-color: var(--primary); box-shadow: 0 0 0 2px rgba(0,120,212,.15); }
    .controls label { font-size: 12px; color: var(--muted); }
    .controls select { padding: 7px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 13px; }

    .table-wrapper { padding: 0 32px 32px; overflow-x: auto; }
    table.main { width: 100%; border-collapse: collapse; background: var(--card); border-radius: 8px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.08); }
    table.main thead tr { background: var(--primary); color: #fff; }
    table.main thead th { padding: 11px 14px; text-align: left; font-weight: 600; font-size: 12px; letter-spacing: .03em; }
    table.main tbody tr { border-bottom: 1px solid var(--border); transition: background .12s; }
    table.main tbody tr:last-child { border-bottom: none; }
    table.main tbody tr:hover { background: #f0f7ff; }
    table.main td { padding: 10px 14px; vertical-align: top; }

    td.domain-name { font-weight: 600; color: var(--primary); white-space: nowrap; min-width: 160px; }
    td.ns-cell { font-family: 'Courier New', monospace; font-size: 12px; color: var(--muted); white-space: nowrap; }
    td.ttl-cell { font-family: 'Courier New', monospace; font-size: 12px; color: var(--muted); text-align: center; white-space: nowrap; }
    td.status-cell { white-space: nowrap; }

    table.inner-table { border-collapse: collapse; width: 100%; font-size: 12px; }
    table.inner-table thead tr { background: #e8f0fe; }
    table.inner-table th { padding: 4px 8px; font-weight: 600; color: #444; text-align: left; }
    table.inner-table td { padding: 3px 8px; border-top: 1px solid #eee; }
    table.txt-table td { word-break: break-all; }
    table.txt-table td.ttl-inline { text-align: center; white-space: nowrap; color: var(--muted); }
    td.pref { text-align: right; color: var(--muted); width: 36px; }

    ul.txt-list { list-style: none; padding: 0; margin: 0; font-size: 12px; }
    ul.txt-list li { padding: 2px 0; border-bottom: 1px dashed #eee; word-break: break-all; }
    ul.txt-list li:last-child { border-bottom: none; }
    ul.txt-list li.spf { color: var(--ok); font-weight: 600; }

    .dmarc-value { font-size: 11px; color: var(--muted); word-break: break-all; display: block; margin-top: 4px; }
    
    td.dkim-sel { font-weight: 600; color: var(--primary); white-space: nowrap; }
    td.dkim-val { font-family: 'Courier New', monospace; font-size: 11px; color: var(--muted); word-break: break-all; }

    .badge {
      display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px;
      font-weight: 600; margin: 2px 2px 2px 0; white-space: nowrap;
    }
    .badge-reject     { background: #eafaf1; color: var(--ok); border: 1px solid #a9dfbf; }
    .badge-quarantine { background: #fef3e2; color: var(--quarantine); border: 1px solid #fddba8; }
    .badge-none       { background: #fdecea; color: var(--warn); border: 1px solid #f5c6c3; }
    .badge-unknown    { background: #f9f9f9; color: #999; border: 1px solid #ddd; }
    .badge-ok         { background: #eafaf1; color: var(--ok); border: 1px solid #a9dfbf; }
    .badge-warn       { background: #fdecea; color: var(--warn); border: 1px solid #f5c6c3; }

    .missing { color: #aaa; font-style: italic; font-size: 12px; }
    .muted { color: #ccc; font-style: italic; }

    .hidden { display: none; }
    .no-results { text-align: center; padding: 30px; color: var(--muted); font-style: italic; }

    footer { text-align: center; padding: 16px; font-size: 11px; color: var(--muted); }
  </style>
</head>
<body>

<header>
  <h1>&#x1F4CB; DNS Records Report</h1>
  <div class="meta">Generated: $generatedAt &nbsp;|&nbsp; $totalDomains domains</div>
</header>

<div class="summary">
  <div class="stat-card"><div class="num">$totalDomains</div><div class="lbl">Total Domains</div></div>
  <div class="stat-card"><div class="num ok">$dmarcReject</div><div class="lbl">DMARC: reject</div></div>
  <div class="stat-card"><div class="num qua">$dmarcQuarantine</div><div class="lbl">DMARC: quarantine</div></div>
  <div class="stat-card"><div class="num warn">$dmarcNone</div><div class="lbl">DMARC: none</div></div>
  <div class="stat-card"><div class="num warn">$noDmarc</div><div class="lbl">No DMARC</div></div>
  <div class="stat-card"><div class="num warn">$noSpf</div><div class="lbl">No SPF</div></div>
  <div class="stat-card"><div class="num warn">$noDkim</div><div class="lbl">No DKIM</div></div>
</div>

<div class="controls">
  <input type="text" id="searchBox" placeholder="&#x1F50D; Filter by domain or record value..." oninput="filterTable()" />
  <label for="dmarcFilter">DMARC policy:</label>
  <select id="dmarcFilter" onchange="filterTable()">
    <option value="">All</option>
    <option value="reject">reject</option>
    <option value="quarantine">quarantine</option>
    <option value="none">none</option>
    <option value="missing">missing</option>
  </select>
</div>

<div class="table-wrapper">
  <table class="main" id="mainTable">
    <thead>
      <tr>
        <th>Domain</th>
        <th>Nameserver</th>
        <th>MX Records</th>
        <th>MX TTL</th>
        <th>SPF</th>
        <th>SPF TTL</th>
        <th>DMARC</th>
        <th>DMARC TTL</th>
        <th>DKIM</th>
        <th>DKIM TTL</th>
        <th>Status</th>
      </tr>
    </thead>
    <tbody id="tableBody">
$($domainRows -join "`n")
    </tbody>
  </table>
  <div class="no-results hidden" id="noResults">No domains match your filter.</div>
</div>

<footer>DNS Records Report &mdash; Generated by Export-DNSReport.ps1</footer>

<script>
  function filterTable() {
    const search = document.getElementById('searchBox').value.toLowerCase();
    const dmarc  = document.getElementById('dmarcFilter').value.toLowerCase();
    const rows   = document.querySelectorAll('#tableBody tr.domain-row');
    let visible  = 0;

    rows.forEach(row => {
      const text        = row.textContent.toLowerCase();
      const statusCell  = row.querySelector('.status-cell') ? row.querySelector('.status-cell').textContent.toLowerCase() : '';
      const dmarcCell   = row.cells[3] ? row.cells[3].textContent.toLowerCase() : '';

      const matchSearch = !search || text.includes(search);

      let matchDmarc = true;
      if (dmarc === 'missing') {
        matchDmarc = !dmarcCell.includes('v=dmarc1');
      } else if (dmarc) {
        matchDmarc = dmarcCell.includes('p=' + dmarc) || statusCell.includes('dmarc: ' + dmarc);
      }

      if (matchSearch && matchDmarc) {
        row.classList.remove('hidden');
        visible++;
      } else {
        row.classList.add('hidden');
      }
    });

    document.getElementById('noResults').classList.toggle('hidden', visible > 0);
  }
</script>
</body>
</html>
"@

$html | Set-Content -Path $OutputPath -Encoding UTF8

Write-Host "Done. Report saved to: $OutputPath" -ForegroundColor Green
Invoke-Item $OutputPath
