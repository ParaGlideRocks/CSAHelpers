<#
.SYNOPSIS
    Analizza i report aggregati DMARC (RUA) e produce un riepilogo Excel/CSV.

.DESCRIPTION
    Legge tutti i report DMARC presenti in una cartella (.xml, .gz, .zip),
    li espande, ne estrae i record e produce tre viste:

      1. Riepilogo      - una riga per report: reporter, periodo, policy, volumi, esiti
      2. DaVerificare   - solo le sorgenti che FALLISCONO il DMARC, aggregate per IP
      3. Dettaglio      - una riga per record: IP sorgente, esito DMARC, SPF/DKIM, reason

    Se il modulo ImportExcel e' presente viene generato un .xlsx multi-foglio,
    altrimenti tre file .csv equivalenti.

    Tutti i campi opzionali della specifica DMARC (selector, sp, pct, reason,
    envelope_from, ...) sono letti in modo tollerante: se mancano, il report
    viene comunque elaborato.

.PARAMETER Path
    Cartella contenente i report (.xml / .gz / .zip). Default: .\

.PARAMETER OutFile
    File Excel di destinazione. Default: .\DMARC_Analisi_<data>.xlsx

.PARAMETER ResolveHostnames
    Se specificato, tenta la risoluzione PTR degli IP che falliscono il DMARC
    (utile per identificare il servizio mittente). Rallenta l'esecuzione.

.PARAMETER SkipDettaglio
    Nel foglio Dettaglio include solo i record in FAIL. Utile con migliaia di
    report, quando serve solo la vista decisionale.

.PARAMETER ReportId
    Filtra i soli report il cui report_id corrisponde (anche parzialmente) al
    valore indicato. Utile per risalire a uno specifico report.

.EXAMPLE
    .\Analyze-DmarcReports.ps1 -Path C:\DMARC\raw

.EXAMPLE
    .\Analyze-DmarcReports.ps1 -Path .\raw\ -ResolveHostnames -OutFile C:\DMARC\report.xlsx

.NOTES
    Versione 1.2.0 - non richiede Outlook ne' privilegi amministrativi.
    Per installare il modulo Excel (una tantum, da PowerShell NON elevato):
        Install-Module ImportExcel -Scope CurrentUser
#>

[CmdletBinding()]
param(
    [string] $Path = ".",
    [string] $OutFile,
    [switch] $ResolveHostnames,
    [switch] $SkipDettaglio,
    [string] $ReportId
)

# NIENTE StrictMode: i report DMARC hanno molti elementi opzionali e
# l'accesso a un nodo assente deve restituire vuoto, non interrompere.
Set-StrictMode -Off
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Path)) {
    throw "Cartella non trovata: $Path"
}
if (-not $OutFile) {
    $OutFile = Join-Path (Resolve-Path $Path) ("DMARC_Analisi_{0:yyyyMMdd-HHmm}.xlsx" -f (Get-Date))
}

# ---------------------------------------------------------------------------
# Helper: lettura tollerante dei nodi XML
# ---------------------------------------------------------------------------

function Get-XVal {
    <# Restituisce il testo di un sottoelemento, o "" se assente. #>
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
    <# Restituisce un sottonodo (o array di sottonodi), o $null se assente. #>
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
# 1. Raccolta e decompressione
# ---------------------------------------------------------------------------

$workDir = Join-Path ([System.IO.Path]::GetTempPath()) ("dmarc_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

Write-Host "Raccolta report da: $Path" -ForegroundColor Cyan
$sourceFiles = @(Get-ChildItem -LiteralPath $Path -File -Recurse -Include *.xml, *.gz, *.zip)

if ($sourceFiles.Count -eq 0) {
    throw "Nessun file .xml/.gz/.zip trovato in $Path"
}

# ogni elemento tiene traccia sia del file XML da leggere sia del file di
# ORIGINE (l'allegato .gz/.zip/.xml come salvato dalla casella RUA)
$xmlFiles = New-Object System.Collections.Generic.List[object]
$skipped  = New-Object System.Collections.Generic.List[string]

$i = 0
foreach ($f in $sourceFiles) {
    $i++
    if ($i % 250 -eq 0) { Write-Host "  espansi $i / $($sourceFiles.Count)..." -ForegroundColor DarkGray }
    try {
        switch ($f.Extension.ToLower()) {
            ".xml" {
                $xmlFiles.Add([pscustomobject]@{ XmlPath = $f.FullName; Origine = $f.FullName })
            }
            ".gz"  {
                # i report DMARC sono gzip a file singolo, NON tar.gz
                $target = Join-Path $workDir ([guid]::NewGuid().ToString("N") + ".xml")
                Expand-GzipFile -Source $f.FullName -Destination $target
                $xmlFiles.Add([pscustomobject]@{ XmlPath = $target; Origine = $f.FullName })
            }
            ".zip" {
                $sub = Join-Path $workDir ([guid]::NewGuid().ToString("N"))
                New-Item -ItemType Directory -Path $sub -Force | Out-Null
                Expand-Archive -LiteralPath $f.FullName -DestinationPath $sub -Force
                Get-ChildItem -LiteralPath $sub -Recurse -Filter *.xml | ForEach-Object {
                    # per gli zip l'origine indica anche la voce interna
                    $xmlFiles.Add([pscustomobject]@{
                        XmlPath = $_.FullName
                        Origine = "$($f.FullName)!$($_.Name)"
                    })
                }
            }
        }
    }
    catch {
        $skipped.Add("$($f.Name) -> $($_.Exception.Message)")
    }
}

Write-Host "Report da analizzare: $($xmlFiles.Count)" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# 2. Parsing
# ---------------------------------------------------------------------------

$dettaglio = New-Object System.Collections.Generic.List[object]
$riepilogo = New-Object System.Collections.Generic.List[object]

$i = 0
foreach ($entry in $xmlFiles) {

    $i++
    if ($i % 250 -eq 0) { Write-Host "  analizzati $i / $($xmlFiles.Count)..." -ForegroundColor DarkGray }

    $file    = $entry.XmlPath
    $origine = $entry.Origine
    $origineNome = Split-Path $origine -Leaf

    $doc = $null
    try { $doc = [xml](Get-Content -LiteralPath $file -Raw -Encoding UTF8) }
    catch {
        $skipped.Add("$origineNome -> XML non valido")
        continue
    }

    $feedback = Get-XNode $doc "feedback"
    if ($null -eq $feedback) {
        $skipped.Add("$origineNome -> non e' un report DMARC")
        continue
    }

    $meta   = Get-XNode $feedback "report_metadata"
    $policy = Get-XNode $feedback "policy_published"
    $range  = Get-XNode $meta "date_range"

    $begin = ConvertFrom-UnixTime (Get-XVal $range "begin")
    $end   = ConvertFrom-UnixTime (Get-XVal $range "end")

    $reporter = Get-XVal $meta   "org_name"
    $repId    = Get-XVal $meta   "report_id"

    # filtro opzionale su report_id (match parziale, case-insensitive)
    if ($ReportId -and ($repId -notlike "*$ReportId*")) { continue }

    $dominio  = Get-XVal $policy "domain"
    $pP       = Get-XVal $policy "p"
    $pSP      = Get-XVal $policy "sp"
    $pPct     = Get-XVal $policy "pct"
    $pAspf    = Get-XVal $policy "aspf"
    $pAdkim   = Get-XVal $policy "adkim"

    $records = @(Get-XNode $feedback "record")
    $tot = 0; $ok = 0; $ko = 0; $quar = 0; $rej = 0

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

        # DMARC passa se ALMENO UNO tra SPF e DKIM passa ED e' allineato
        $dmarcPass = ($dkimEval -eq "pass") -or ($spfEval -eq "pass")

        # motivi di override applicati dal destinatario (es. arc=fail, forwarded)
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

        # auth_results: possono esserci piu' firme DKIM, e 'selector' e' opzionale
        $authDkim = ""; $authDkimSel = ""; $authSpf = ""; $authSpfDom = ""
        if ($null -ne $auth) {
            $dkimNodes = @(Get-XNode $auth "dkim")
            if ($dkimNodes.Count -gt 0) {
                $d1 = @(); $d2 = @()
                foreach ($d in $dkimNodes) {
                    if ($null -eq $d) { continue }
                    $dd = Get-XVal $d "domain"
                    $dr = Get-XVal $d "result"
                    $ds = Get-XVal $d "selector" "(n/d)"
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

        $tot += $count
        if ($dmarcPass) { $ok += $count } else { $ko += $count }
        if ($disp -eq "quarantine") { $quar += $count }
        if ($disp -eq "reject")     { $rej  += $count }

        if ((-not $SkipDettaglio) -or (-not $dmarcPass)) {
            $dettaglio.Add([pscustomobject][ordered]@{
                Reporter        = $reporter
                ReportId        = $repId
                PeriodoDa       = $begin
                PeriodoA        = $end
                Dominio         = $dominio
                PolicyP         = $pP
                PolicySP        = $pSP
                Pct             = $pPct
                SourceIP        = (Get-XVal $row "source_ip")
                Messaggi        = $count
                Disposition     = $disp
                DMARC           = $(if ($dmarcPass) { "PASS" } else { "FAIL" })
                DKIM_Allineato  = $dkimEval
                SPF_Allineato   = $spfEval
                HeaderFrom      = (Get-XVal $ident "header_from")
                EnvelopeFrom    = (Get-XVal $ident "envelope_from")
                DKIM_Auth       = $authDkim
                DKIM_Selector   = $authDkimSel
                SPF_Auth        = $authSpf
                SPF_Dominio     = $authSpfDom
                OverrideTipo    = $reasonType
                OverrideNota    = $reasonComment
                FileOrigine     = $origineNome
                PercorsoOrigine = $origine
            })
        }
    }

    $riepilogo.Add([pscustomobject][ordered]@{
        Reporter          = $reporter
        ReportId          = $repId
        PeriodoDa         = $begin
        PeriodoA          = $end
        Dominio           = $dominio
        PolicyPubblicata  = "p=$pP; sp=$pSP; pct=$pPct; aspf=$pAspf; adkim=$pAdkim"
        MessaggiTotali    = $tot
        DMARC_Pass        = $ok
        DMARC_Fail        = $ko
        PercentualePass   = $(if ($tot -gt 0) { [math]::Round(100.0 * $ok / $tot, 2) } else { 0 })
        InQuarantena      = $quar
        Respinti          = $rej
        FileOrigine       = $origineNome
        PercorsoOrigine   = $origine
    })
}

if ($riepilogo.Count -eq 0) {
    if ($ReportId) {
        throw "Nessun report con report_id corrispondente a '$ReportId'."
    }
    throw "Nessun report DMARC valido trovato: verifica il contenuto della cartella."
}

if ($ReportId) {
    Write-Host "`nReport corrispondenti a '$ReportId': $($riepilogo.Count)" -ForegroundColor Yellow
    $riepilogo | Select-Object ReportId, Reporter, Dominio, MessaggiTotali, FileOrigine |
        Format-Table -AutoSize | Out-String | Write-Host
}

# ---------------------------------------------------------------------------
# 3. Sorgenti da verificare (i FAIL aggregati per IP)
# ---------------------------------------------------------------------------

$daVerificare = @(
    $dettaglio |
        Where-Object { $_.DMARC -eq "FAIL" } |
        Group-Object SourceIP, Dominio |
        ForEach-Object {
            $g = $_.Group
            [pscustomobject][ordered]@{
                SourceIP        = $g[0].SourceIP
                Dominio         = $g[0].Dominio
                Hostname        = ""
                MessaggiTotali  = ($g | Measure-Object Messaggi -Sum).Sum
                SPF_Allineato   = (($g | Select-Object -ExpandProperty SPF_Allineato  -Unique) -join ", ")
                DKIM_Allineato  = (($g | Select-Object -ExpandProperty DKIM_Allineato -Unique) -join ", ")
                Disposition     = (($g | Select-Object -ExpandProperty Disposition    -Unique) -join ", ")
                HeaderFrom      = (($g | Select-Object -ExpandProperty HeaderFrom     -Unique) -join ", ")
                Reporter        = (($g | Select-Object -ExpandProperty Reporter       -Unique) -join ", ")
                VistoDa         = ($g | Sort-Object PeriodoDa | Select-Object -First 1).PeriodoDa
                VistoFinoA      = ($g | Sort-Object PeriodoA  | Select-Object -Last  1).PeriodoA
                NumFileOrigine  = @($g | Select-Object -ExpandProperty FileOrigine -Unique).Count
                FileOrigine     = $(
                    $f = @($g | Select-Object -ExpandProperty FileOrigine -Unique)
                    if ($f.Count -gt 5) { ($f[0..4] -join "; ") + "; ... (+$($f.Count - 5))" }
                    else { $f -join "; " }
                )
            }
        } | Sort-Object MessaggiTotali -Descending
)

if ($ResolveHostnames -and $daVerificare.Count -gt 0) {
    Write-Host "Risoluzione PTR di $($daVerificare.Count) sorgenti in errore..." -ForegroundColor Cyan
    foreach ($r in $daVerificare) {
        try   { $r.Hostname = [System.Net.Dns]::GetHostEntry($r.SourceIP).HostName }
        catch { $r.Hostname = "(non risolto)" }
    }
}

# ---------------------------------------------------------------------------
# 4. Output
# ---------------------------------------------------------------------------

$totMsg  = ($riepilogo | Measure-Object MessaggiTotali -Sum).Sum
$totPass = ($riepilogo | Measure-Object DMARC_Pass     -Sum).Sum
$totFail = ($riepilogo | Measure-Object DMARC_Fail     -Sum).Sum
$totQuar = ($riepilogo | Measure-Object InQuarantena   -Sum).Sum
$totRej  = ($riepilogo | Measure-Object Respinti       -Sum).Sum

$hasExcel = $null -ne (Get-Module -ListAvailable -Name ImportExcel)

if ($hasExcel) {
    Import-Module ImportExcel
    if (Test-Path -LiteralPath $OutFile) { Remove-Item -LiteralPath $OutFile -Force }

    $riepilogo | Sort-Object PeriodoDa -Descending |
        Export-Excel -Path $OutFile -WorksheetName "Riepilogo" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter

    if ($daVerificare.Count -gt 0) {
        $daVerificare |
            Export-Excel -Path $OutFile -WorksheetName "DaVerificare" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
    }

    $dettaglio | Sort-Object PeriodoDa -Descending |
        Export-Excel -Path $OutFile -WorksheetName "Dettaglio" -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter

    Write-Host "`nExcel generato: $OutFile" -ForegroundColor Green
}
else {
    $base = [System.IO.Path]::ChangeExtension($OutFile, $null).TrimEnd('.')
    $riepilogo    | Export-Csv "$base-Riepilogo.csv"    -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    $daVerificare | Export-Csv "$base-DaVerificare.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    $dettaglio    | Export-Csv "$base-Dettaglio.csv"    -NoTypeInformation -Encoding UTF8 -Delimiter ';'

    Write-Warning "Modulo ImportExcel non presente: generati 3 file CSV al posto dell'Excel."
    Write-Host "Per l'output Excel: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    Write-Host "`nCSV generati con prefisso: $base" -ForegroundColor Green
}

Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# 5. Sintesi a video
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "=== SINTESI ===" -ForegroundColor Cyan
Write-Host ("Report analizzati      : {0}" -f $riepilogo.Count)
Write-Host ("Messaggi totali        : {0}" -f $totMsg)
Write-Host ("DMARC PASS             : {0}" -f $totPass) -ForegroundColor Green
Write-Host ("DMARC FAIL             : {0}" -f $totFail) -ForegroundColor $(if ($totFail -gt 0) { "Yellow" } else { "Green" })
if ($totMsg -gt 0) {
    Write-Host ("Percentuale PASS       : {0}%" -f ([math]::Round(100.0 * $totPass / $totMsg, 2)))
}
Write-Host ("Messi in quarantena    : {0}" -f $totQuar)
Write-Host ("Respinti               : {0}" -f $totRej)
Write-Host ("Sorgenti da verificare : {0}" -f $daVerificare.Count) -ForegroundColor $(if ($daVerificare.Count -gt 0) { "Yellow" } else { "Green" })

if ($daVerificare.Count -gt 0) {
    Write-Host ""
    Write-Host "Prime 10 sorgenti in errore per volume:" -ForegroundColor Yellow
    $daVerificare | Select-Object -First 10 SourceIP, Dominio, MessaggiTotali, SPF_Allineato, DKIM_Allineato |
        Format-Table -AutoSize | Out-String | Write-Host
}

if ($skipped.Count -gt 0) {
    Write-Host ""
    Write-Warning "File non elaborati: $($skipped.Count)"
    $skipped | Select-Object -First 20 | ForEach-Object { Write-Host "  - $_" -ForegroundColor DarkYellow }
    if ($skipped.Count -gt 20) { Write-Host "  ... e altri $($skipped.Count - 20)" -ForegroundColor DarkYellow }
}

Write-Host ""
Write-Host "Nota: un DMARC FAIL non significa automaticamente spoofing." -ForegroundColor DarkGray
Write-Host "Vanno distinti i servizi legittimi non allineati (da sistemare PRIMA" -ForegroundColor DarkGray
Write-Host "del passaggio a p=reject) dal traffico realmente malevolo." -ForegroundColor DarkGray
