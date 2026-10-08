<#
.SYNOPSIS
    Analyzes aggregate DMARC reports (RUA) and produces an Excel/CSV summary.

.DESCRIPTION
    Reads all DMARC reports in a folder (.xml, .gz, .zip), extracts them,
    parses their records, and produces three views:

      1. Summary          - one row per report: reporter, period, policy, volumes, results
      2. SourcesToReview  - only sources that FAIL DMARC, aggregated by IP
      3. Details          - one row per record: source IP, DMARC result, SPF/DKIM, reason

    If the ImportExcel module is available, a multi-sheet .xlsx file is generated;
    otherwise, three equivalent .csv files are generated.

    All optional fields in the DMARC specification (selector, sp, pct, reason,
    envelope_from, etc.) are read tolerantly: if they are missing, the report
    is still processed.

.PARAMETER Path
    Folder containing the reports (.xml / .gz / .zip). Default: .\

.PARAMETER XMLFolderPath
    Optional path to the folder containing the XML reports. When specified,
    this value overrides Path for report discovery and the default output location.

.PARAMETER OutFile
    Destination Excel file. By default, the file is created in the selected
    input folder as DMARC_Analysis_<date>.xlsx.

.PARAMETER ResolveHostnames
    If specified, attempts PTR resolution for IPs that fail DMARC
    (useful for identifying the sending service). Slows execution.

.PARAMETER SkipDetails
    Includes only FAIL records in the Details sheet. Useful with thousands of
    reports when only the decision-making view is needed.

.PARAMETER ReportId
    Filters reports whose report_id matches the specified value, including
    partial matches. Useful for locating a specific report.

.EXAMPLE
    .\Analyze-DmarcReports.ps1 -Path C:\DMARC\raw

.EXAMPLE
    .\Analyze-DmarcReports.ps1 -XMLFolderPath C:\DMARC\xml

.EXAMPLE
    .\Analyze-DmarcReports.ps1 -Path .\raw\ -ResolveHostnames -OutFile C:\DMARC\report.xlsx

.NOTES
    Version 1.2.0 - does not require Outlook or administrative privileges.
    To install the Excel module (once, from a non-elevated PowerShell session):
        Install-Module ImportExcel -Scope CurrentUser
#>

[CmdletBinding()]
param(
    [string] $Path = ".",
    [string] $XMLFolderPath,
    [string] $OutFile,
    [switch] $ResolveHostnames,
    [switch] $SkipDetails,
    [string] $ReportId
)

# Do not use StrictMode: DMARC reports have many optional elements, and
# accessing a missing node must return an empty value rather than stop execution.
Set-StrictMode -Off
$ErrorActionPreference = "Stop"

$inputPath = if ($XMLFolderPath) { $XMLFolderPath } else { $Path }

if (-not (Test-Path -LiteralPath $inputPath -PathType Container)) {
    throw "Folder not found: $inputPath"
}
if (-not $OutFile) {
    $OutFile = Join-Path (Resolve-Path $inputPath) ("DMARC_Analysis_{0:yyyyMMdd-HHmm}.xlsx" -f (Get-Date))
}

# ---------------------------------------------------------------------------
# Helper: tolerant XML node reading
# ---------------------------------------------------------------------------

function Get-XVal {
    <# Returns the text of a child element, or "" if it is missing. #>
    param($Node, [string] $Name, [string] $Default = "")
    if ($null -eq $Node) { return $Default }
    try {
        $prop = $Node.PSObject.Properties[$Name]
        if ($null -eq $prop) { return $Default }
        $v = $prop.Value
        if ($null -eq $v) { return $Default }
        if ($v -is [System.Xml.XmlElement]) {
            $t = $v.InnerText
            if ([string]::IsNullOrWhiteSpace($t)) { return $Default }
            return $t.Trim()
        }
        $s = [string]$v
        if ([string]::IsNullOrWhiteSpace($s)) { return $Default }
        return $s.Trim()
    }
    catch { return $Default }
}

function Get-XNode {
    <# Returns a child node (or array of child nodes), or $null if missing. #>
    param($Node, [string] $Name)
    if ($null -eq $Node) { return $null }
    try {
        $prop = $Node.PSObject.Properties[$Name]
        if ($null -eq $prop) { return $null }
        return $prop.Value
    }
    catch { return $null }
}

function ConvertFrom-UnixTime {
    param($Value)
    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    try { return [System.DateTimeOffset]::FromUnixTimeSeconds([int64]$Value).LocalDateTime }
    catch { return $null }
}

function Expand-GzipFile {
    param([string] $Source, [string] $Destination)
    $in = $null; $gz = $null; $out = $null
    try {
        $in  = [System.IO.File]::OpenRead($Source)
        $gz  = New-Object System.IO.Compression.GZipStream($in, [System.IO.Compression.CompressionMode]::Decompress)
        $out = [System.IO.File]::Create($Destination)
        $gz.CopyTo($out)
    }
    finally {
        if ($out) { $out.Dispose() }
        if ($gz)  { $gz.Dispose()  }
        if ($in)  { $in.Dispose()  }
    }
}

# ---------------------------------------------------------------------------
# 1. Collection and decompression
# ---------------------------------------------------------------------------

$workDir = Join-Path ([System.IO.Path]::GetTempPath()) ("dmarc_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

Write-Host "Collecting reports from: $inputPath" -ForegroundColor Cyan
$sourceFiles = @(Get-ChildItem -LiteralPath $inputPath -File -Recurse -Include *.xml, *.gz, *.zip)

if ($sourceFiles.Count -eq 0) {
    throw "No .xml/.gz/.zip files found in $inputPath"
}

# Each item tracks both the XML file to read and the SOURCE file
# (the .gz/.zip/.xml attachment as saved from the RUA mailbox).
$xmlFiles = New-Object System.Collections.Generic.List[object]
$skipped  = New-Object System.Collections.Generic.List[string]

$i = 0
foreach ($f in $sourceFiles) {
    $i++
    if ($i % 250 -eq 0) { Write-Host "  extracted $i / $($sourceFiles.Count)..." -ForegroundColor DarkGray }
    try {
        switch ($f.Extension.ToLower()) {
            ".xml" {
                $xmlFiles.Add([pscustomobject]@{ XmlPath = $f.FullName; SourcePath = $f.FullName })
            }
            ".gz"  {
                # DMARC reports are single-file gzip archives, NOT tar.gz archives.
                $target = Join-Path $workDir ([guid]::NewGuid().ToString("N") + ".xml")
                Expand-GzipFile -Source $f.FullName -Destination $target
                $xmlFiles.Add([pscustomobject]@{ XmlPath = $target; SourcePath = $f.FullName })
            }
            ".zip" {
                $sub = Join-Path $workDir ([guid]::NewGuid().ToString("N"))
                New-Item -ItemType Directory -Path $sub -Force | Out-Null
                Expand-Archive -LiteralPath $f.FullName -DestinationPath $sub -Force
                Get-ChildItem -LiteralPath $sub -Recurse -Filter *.xml | ForEach-Object {
                    # For ZIP files, the source also identifies the internal entry.
                    $xmlFiles.Add([pscustomobject]@{
                        XmlPath    = $_.FullName
                        SourcePath = "$($f.FullName)!$($_.Name)"
                    })
                }
            }
        }
    }
    catch {
        $skipped.Add("$($f.Name) -> $($_.Exception.Message)")
    }
}

Write-Host "Reports to analyze: $($xmlFiles.Count)" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# 2. Parsing
# ---------------------------------------------------------------------------

$details = New-Object System.Collections.Generic.List[object]
$summary = New-Object System.Collections.Generic.List[object]

$i = 0
foreach ($entry in $xmlFiles) {

    $i++
    if ($i % 250 -eq 0) { Write-Host "  analyzed $i / $($xmlFiles.Count)..." -ForegroundColor DarkGray }

    $file       = $entry.XmlPath
    $sourcePath = $entry.SourcePath
    $sourceName = Split-Path $sourcePath -Leaf

    $doc = $null
    try { $doc = [xml](Get-Content -LiteralPath $file -Raw -Encoding UTF8) }
    catch {
        $skipped.Add("$sourceName -> invalid XML")
        continue
    }

    $feedback = Get-XNode $doc "feedback"
    if ($null -eq $feedback) {
        $skipped.Add("$sourceName -> not a DMARC report")
        continue
    }

    $meta   = Get-XNode $feedback "report_metadata"
    $policy = Get-XNode $feedback "policy_published"
    $range  = Get-XNode $meta "date_range"

    $begin = ConvertFrom-UnixTime (Get-XVal $range "begin")
    $end   = ConvertFrom-UnixTime (Get-XVal $range "end")

    $reporter = Get-XVal $meta   "org_name"
    $repId    = Get-XVal $meta   "report_id"

    # Optional report_id filter (partial, case-insensitive match).
    if ($ReportId -and ($repId -notlike "*$ReportId*")) { continue }

    $domain   = Get-XVal $policy "domain"
    $pP       = Get-XVal $policy "p"
    $pSP      = Get-XVal $policy "sp"
    $pPct     = Get-XVal $policy "pct"
    $pAspf    = Get-XVal $policy "aspf"
    $pAdkim   = Get-XVal $policy "adkim"

    $records = @(Get-XNode $feedback "record")
    $messageCount = 0; $passCount = 0; $failCount = 0; $quarantineCount = 0; $rejectCount = 0

    foreach ($rec in $records) {
        if ($null -eq $rec) { continue }

        $row = Get-XNode $rec "row"
        $pe  = Get-XNode $row "policy_evaluated"

        $count = 1
        $rawCount = Get-XVal $row "count"
        if ($rawCount -match '^\d+$') { $count = [int]$rawCount }
        if ($count -le 0) { $count = 1 }

        $dkimEval = (Get-XVal $pe "dkim").ToLower()
        $spfEval  = (Get-XVal $pe "spf").ToLower()
        $disp     = (Get-XVal $pe "disposition").ToLower()

        # DMARC passes if AT LEAST ONE of SPF and DKIM passes AND is aligned.
        $dmarcPass = ($dkimEval -eq "pass") -or ($spfEval -eq "pass")

        # Override reasons applied by the recipient (for example, arc=fail or forwarded).
        $reasonType = ""; $reasonComment = ""
        $reasons = @(Get-XNode $pe "reason")
        if ($reasons.Count -gt 0) {
            $rt = @(); $rc = @()
            foreach ($r in $reasons) {
                if ($null -eq $r) { continue }
                $t = Get-XVal $r "type";    if ($t) { $rt += $t }
                $c = Get-XVal $r "comment"; if ($c) { $rc += $c }
            }
            $reasonType    = ($rt -join "; ")
            $reasonComment = ($rc -join "; ")
        }

        $ident = Get-XNode $rec "identifiers"
        $auth  = Get-XNode $rec "auth_results"

        # auth_results may contain multiple DKIM signatures, and 'selector' is optional.
        $authDkim = ""; $authDkimSel = ""; $authSpf = ""; $authSpfDom = ""
        if ($null -ne $auth) {
            $dkimNodes = @(Get-XNode $auth "dkim")
            if ($dkimNodes.Count -gt 0) {
                $d1 = @(); $d2 = @()
                foreach ($d in $dkimNodes) {
                    if ($null -eq $d) { continue }
                    $dd = Get-XVal $d "domain"
                    $dr = Get-XVal $d "result"
                    $ds = Get-XVal $d "selector" "(n/a)"
                    $d1 += "$dd=$dr"
                    $d2 += $ds
                }
                $authDkim    = ($d1 -join "; ")
                $authDkimSel = ($d2 -join "; ")
            }
            $spfNodes = @(Get-XNode $auth "spf")
            if ($spfNodes.Count -gt 0) {
                $s1 = @(); $s2 = @()
                foreach ($s in $spfNodes) {
                    if ($null -eq $s) { continue }
                    $s1 += (Get-XVal $s "result")
                    $s2 += (Get-XVal $s "domain")
                }
                $authSpf    = ($s1 -join "; ")
                $authSpfDom = ($s2 -join "; ")
            }
        }

        $messageCount += $count
        if ($dmarcPass) { $passCount += $count } else { $failCount += $count }
        if ($disp -eq "quarantine") { $quarantineCount += $count }
        if ($disp -eq "reject")     { $rejectCount += $count }

        if ((-not $SkipDetails) -or (-not $dmarcPass)) {
            $details.Add([pscustomobject][ordered]@{
                Reporter       = $reporter
                ReportId       = $repId
                PeriodFrom     = $begin
                PeriodTo       = $end
                Domain         = $domain
                PolicyP        = $pP
                PolicySP       = $pSP
                Pct            = $pPct
                SourceIP       = (Get-XVal $row "source_ip")
                Messages       = $count
                Disposition    = $disp
                DMARC          = $(if ($dmarcPass) { "PASS" } else { "FAIL" })
                DKIM_Aligned   = $dkimEval
                SPF_Aligned    = $spfEval
                HeaderFrom     = (Get-XVal $ident "header_from")
                EnvelopeFrom   = (Get-XVal $ident "envelope_from")
                DKIM_Auth      = $authDkim
                DKIM_Selector  = $authDkimSel
                SPF_Auth       = $authSpf
                SPF_Domain     = $authSpfDom
                OverrideType   = $reasonType
                OverrideNote   = $reasonComment
                SourceFile     = $sourceName
                SourceFilePath = $sourcePath
            })
        }
    }

    $summary.Add([pscustomobject][ordered]@{
        Reporter        = $reporter
        ReportId        = $repId
        PeriodFrom      = $begin
        PeriodTo        = $end
        Domain          = $domain
        PublishedPolicy = "p=$pP; sp=$pSP; pct=$pPct; aspf=$pAspf; adkim=$pAdkim"
        TotalMessages   = $messageCount
        DMARC_Pass      = $passCount
        DMARC_Fail      = $failCount
        PassPercentage  = $(if ($messageCount -gt 0) { [math]::Round(100.0 * $passCount / $messageCount, 2) } else { 0 })
        Quarantined     = $quarantineCount
        Rejected        = $rejectCount
        SourceFile      = $sourceName
        SourceFilePath  = $sourcePath
    })
}

if ($summary.Count -eq 0) {
    if ($ReportId) {
        throw "No reports have a report_id matching '$ReportId'."
    }
    throw "No valid DMARC reports found. Check the folder contents."
}

if ($ReportId) {
    Write-Host "`nReports matching '$ReportId': $($summary.Count)" -ForegroundColor Yellow
    $summary | Select-Object ReportId, Reporter, Domain, TotalMessages, SourceFile |
        Format-Table -AutoSize | Out-String | Write-Host
}

# ---------------------------------------------------------------------------
# 3. Sources to review (FAIL records aggregated by IP)
# ---------------------------------------------------------------------------

$sourcesToReview = @(
    $details |
        Where-Object { $_.DMARC -eq "FAIL" } |
        Group-Object SourceIP, Domain |
        ForEach-Object {
            $g = $_.Group
            [pscustomobject][ordered]@{
                SourceIP       = $g[0].SourceIP
                Domain         = $g[0].Domain
                Hostname       = ""
                TotalMessages  = ($g | Measure-Object Messages -Sum).Sum
                SPF_Aligned    = (($g | Select-Object -ExpandProperty SPF_Aligned  -Unique) -join ", ")
                DKIM_Aligned   = (($g | Select-Object -ExpandProperty DKIM_Aligned -Unique) -join ", ")
                Disposition    = (($g | Select-Object -ExpandProperty Disposition  -Unique) -join ", ")
                HeaderFrom     = (($g | Select-Object -ExpandProperty HeaderFrom   -Unique) -join ", ")
                Reporter       = (($g | Select-Object -ExpandProperty Reporter     -Unique) -join ", ")
                FirstSeen      = ($g | Sort-Object PeriodFrom | Select-Object -First 1).PeriodFrom
                LastSeen       = ($g | Sort-Object PeriodTo   | Select-Object -Last  1).PeriodTo
                SourceFileCount = @($g | Select-Object -ExpandProperty SourceFile -Unique).Count
                SourceFile      = $(
                    $f = @($g | Select-Object -ExpandProperty SourceFile -Unique)
                    if ($f.Count -gt 5) { ($f[0..4] -join "; ") + "; ... (+$($f.Count - 5))" }
                    else { $f -join "; " }
                )
            }
        } | Sort-Object TotalMessages -Descending
)

if ($ResolveHostnames -and $sourcesToReview.Count -gt 0) {
    Write-Host "Resolving PTR records for $($sourcesToReview.Count) failing sources..." -ForegroundColor Cyan
    foreach ($r in $sourcesToReview) {
        try   { $r.Hostname = [System.Net.Dns]::GetHostEntry($r.SourceIP).HostName }
        catch { $r.Hostname = "(unresolved)" }
    }
}

# ---------------------------------------------------------------------------
# 4. Output
# ---------------------------------------------------------------------------

$totalMessages    = ($summary | Measure-Object TotalMessages -Sum).Sum
$totalPass        = ($summary | Measure-Object DMARC_Pass    -Sum).Sum
$totalFail        = ($summary | Measure-Object DMARC_Fail    -Sum).Sum
$totalQuarantined = ($summary | Measure-Object Quarantined   -Sum).Sum
$totalRejected    = ($summary | Measure-Object Rejected      -Sum).Sum

$hasExcel = $null -ne (Get-Module -ListAvailable -Name ImportExcel)

if ($hasExcel) {
    Import-Module ImportExcel
    if (Test-Path -LiteralPath $OutFile) { Remove-Item -LiteralPath $OutFile -Force }

    $summary | Sort-Object PeriodFrom -Descending |
        Export-Excel -Path $OutFile -WorksheetName "Summary" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter

    if ($sourcesToReview.Count -gt 0) {
        $sourcesToReview |
            Export-Excel -Path $OutFile -WorksheetName "SourcesToReview" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
    }

    $details | Sort-Object PeriodFrom -Descending |
        Export-Excel -Path $OutFile -WorksheetName "Details" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter

    Write-Host "`nExcel file generated: $OutFile" -ForegroundColor Green
}
else {
    $base = [System.IO.Path]::ChangeExtension($OutFile, $null).TrimEnd('.')
    $summary         | Export-Csv "$base-Summary.csv"         -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    $sourcesToReview | Export-Csv "$base-SourcesToReview.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    $details         | Export-Csv "$base-Details.csv"         -NoTypeInformation -Encoding UTF8 -Delimiter ';'

    Write-Warning "The ImportExcel module is not available: generated 3 CSV files instead of an Excel file."
    Write-Host "For Excel output: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    Write-Host "`nCSV files generated with prefix: $base" -ForegroundColor Green
}

Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# 5. On-screen summary
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "=== SUMMARY ===" -ForegroundColor Cyan
Write-Host ("Reports analyzed : {0}" -f $summary.Count)
Write-Host ("Total messages   : {0}" -f $totalMessages)
Write-Host ("DMARC PASS       : {0}" -f $totalPass) -ForegroundColor Green
Write-Host ("DMARC FAIL       : {0}" -f $totalFail) -ForegroundColor $(if ($totalFail -gt 0) { "Yellow" } else { "Green" })
if ($totalMessages -gt 0) {
    Write-Host ("PASS percentage  : {0}%" -f ([math]::Round(100.0 * $totalPass / $totalMessages, 2)))
}
Write-Host ("Quarantined      : {0}" -f $totalQuarantined)
Write-Host ("Rejected         : {0}" -f $totalRejected)
Write-Host ("Sources to review: {0}" -f $sourcesToReview.Count) -ForegroundColor $(if ($sourcesToReview.Count -gt 0) { "Yellow" } else { "Green" })

if ($sourcesToReview.Count -gt 0) {
    Write-Host ""
    Write-Host "Top 10 failing sources by volume:" -ForegroundColor Yellow
    $sourcesToReview | Select-Object -First 10 SourceIP, Domain, TotalMessages, SPF_Aligned, DKIM_Aligned |
        Format-Table -AutoSize | Out-String | Write-Host
}

if ($skipped.Count -gt 0) {
    Write-Host ""
    Write-Warning "Unprocessed files: $($skipped.Count)"
    $skipped | Select-Object -First 20 | ForEach-Object { Write-Host "  - $_" -ForegroundColor DarkYellow }
    if ($skipped.Count -gt 20) { Write-Host "  ... and $($skipped.Count - 20) more" -ForegroundColor DarkYellow }
}

Write-Host ""
Write-Host "Note: a DMARC FAIL does not automatically indicate spoofing." -ForegroundColor DarkGray
Write-Host "Distinguish legitimate unaligned services (which must be fixed BEFORE" -ForegroundColor DarkGray
Write-Host "moving to p=reject) from genuinely malicious traffic." -ForegroundColor DarkGray
