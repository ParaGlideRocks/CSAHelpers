<#
.SYNOPSIS
    Exports MX, TXT, DMARC, and DKIM DNS records for a list of domains.

.DESCRIPTION
    Reads a CSV file containing a "Domain" column and queries DNS for MX, TXT,
    DMARC (_dmarc TXT), and DKIM (selector._domainkey CNAME) records for each domain.
    Queries common DKIM selectors (default, selector1, selector2, google, mandrill).
    Queries are sent to each domain's authoritative name servers when available.
    Results include record TTL and DNS query timings, and are exported as JSON.

.PARAMETER CsvPath
    Path to the input CSV file. Must contain a "Domain" column.

.PARAMETER OutputPath
    Path for the JSON output file. Defaults to a timestamp-prefixed file name
    like "yyyyMMdd-HHmmss_DNSRecordsExport.json" in the same directory as the CSV.

.PARAMETER DnsServer
    Optional DNS server used as a bootstrap resolver for NS discovery and as
    fallback when authoritative name servers are unavailable.

.EXAMPLE
    .\Export-DNSRecords.ps1 -CsvPath .\domains.csv

.EXAMPLE
    .\Export-DNSRecords.ps1 -CsvPath .\domains.csv -OutputPath .\results.json -DnsServer 8.8.8.8
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [string]$DnsServer
)

function Resolve-DnsRecordSafe {
    param (
        [string]$Name,
        [string]$Type,
        [string]$Server
    )

    $params = @{ Name = $Name; Type = $Type; ErrorAction = 'SilentlyContinue' }
    if ($Server) { $params['Server'] = $Server }

    try {
        Resolve-DnsName @params
    }
    catch {
        $null
    }
}

function Get-AuthoritativeNameServers {
    param(
        [Parameter(Mandatory)]
        [string]$Domain,
        [string]$BootstrapServer
    )

    Resolve-DnsRecordSafe -Name $Domain -Type NS -Server $BootstrapServer |
        Where-Object { $_.Type -eq 'NS' } |
        ForEach-Object { $_.NameHost.TrimEnd('.') } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
}

function Resolve-DnsAcrossServers {
    param(
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [string]$Type,
        [string[]]$Servers,
        [string]$FallbackServer
    )

    $responses = [System.Collections.Generic.List[object]]::new()

    foreach ($server in @($Servers)) {
        $queryTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $response = Resolve-DnsRecordSafe -Name $Name -Type $Type -Server $server
        $queryTimer.Stop()

        foreach ($record in @($response)) {
            $responses.Add([PSCustomObject]@{
                Record          = $record
                QueriedServer   = $server
                QueryDurationMs = $queryTimer.ElapsedMilliseconds
            })
        }
    }

    # If authoritative queries returned nothing, use bootstrap/fallback resolver.
    if ($responses.Count -eq 0 -and $FallbackServer -and (@($Servers) -notcontains $FallbackServer)) {
        $queryTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $response = Resolve-DnsRecordSafe -Name $Name -Type $Type -Server $FallbackServer
        $queryTimer.Stop()

        foreach ($record in @($response)) {
            $responses.Add([PSCustomObject]@{
                Record          = $record
                QueriedServer   = $FallbackServer
                QueryDurationMs = $queryTimer.ElapsedMilliseconds
            })
        }
    }

    # If there are no candidate servers and no fallback was supplied, use system resolver.
    if ($responses.Count -eq 0 -and (-not @($Servers) -or @($Servers).Count -eq 0) -and -not $FallbackServer) {
        $queryTimer = [System.Diagnostics.Stopwatch]::StartNew()
        $response = Resolve-DnsRecordSafe -Name $Name -Type $Type
        $queryTimer.Stop()

        foreach ($record in @($response)) {
            $responses.Add([PSCustomObject]@{
                Record          = $record
                QueriedServer   = 'SystemResolver'
                QueryDurationMs = $queryTimer.ElapsedMilliseconds
            })
        }
    }

    $responses
}

function Join-TxtStrings {
    param(
        [Parameter(Mandatory)]
        $Record
    )

    if ($null -eq $Record.Strings) {
        return $null
    }

    # TXT RDATA can be split into multiple character-strings; join into one value.
    ($Record.Strings -join '').Trim()
}

# Set default output path next to the CSV
if (-not $OutputPath) {
    $csvDir  = Split-Path -Parent (Resolve-Path $CsvPath)
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $OutputPath = Join-Path $csvDir ("${timestamp}_DNSRecordsExport.json")
}

$domains = Import-Csv -Path $CsvPath

if (-not ($domains | Get-Member -Name 'Domain' -MemberType NoteProperty)) {
    Write-Error "CSV file must contain a 'Domain' column."
    exit 1
}

$results = [System.Collections.Generic.List[object]]::new()
$total   = ($domains | Measure-Object).Count
$index   = 0

foreach ($row in $domains) {
    $index++
    $domain = ($row.Domain -replace '\s+', '').Trim()

    if ([string]::IsNullOrWhiteSpace($domain)) { continue }

    Write-Progress -Activity "Querying DNS records" `
                   -Status "[$index/$total] $domain" `
                   -PercentComplete (($index / $total) * 100)

    Write-Verbose "Processing: $domain"

    $authoritativeNameServers = @(Get-AuthoritativeNameServers -Domain $domain -BootstrapServer $DnsServer)
    $queriedNameServer = if ($authoritativeNameServers.Count -gt 0) { $authoritativeNameServers[0] } else { $null }
    $queryServers = if ($queriedNameServer) { @($queriedNameServer) } else { @() }

    $mxQueryTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $mxResponses = Resolve-DnsAcrossServers -Name $domain -Type MX -Servers $queryServers -FallbackServer $DnsServer
    $mxQueryTimer.Stop()

    # MX records
    $mxRecords = $mxResponses |
        Where-Object { $_.Record.Type -eq 'MX' } |
        ForEach-Object {
            [PSCustomObject]@{
                Preference      = $_.Record.Preference
                Exchange        = $_.Record.NameExchange
                TTL             = $_.Record.TTL
                QueryDurationMs = $_.QueryDurationMs
            }
        } |
        Sort-Object Preference, Exchange, TTL -Unique

    $txtQueryTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $txtResponses = Resolve-DnsAcrossServers -Name $domain -Type TXT -Servers $queryServers -FallbackServer $DnsServer
    $txtQueryTimer.Stop()

    # TXT records (root domain)
    $txtRecords = $txtResponses |
        Where-Object { $_.Record.Type -eq 'TXT' } |
        ForEach-Object {
            $txtValue = Join-TxtStrings -Record $_.Record
            if ($txtValue) {
                [PSCustomObject]@{
                    Value           = $txtValue
                    TTL             = $_.Record.TTL
                    QueryDurationMs = $_.QueryDurationMs
                }
            }
        } |
        Sort-Object Value, TTL -Unique

    $dmarcQueryTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $dmarcResponses = Resolve-DnsAcrossServers -Name "_dmarc.$domain" -Type TXT -Servers $queryServers -FallbackServer $DnsServer
    $dmarcQueryTimer.Stop()

    # DMARC record (_dmarc TXT)
    $dmarcRecords = $dmarcResponses |
        Where-Object { $_.Record.Type -eq 'TXT' } |
        ForEach-Object {
            $txtValue = Join-TxtStrings -Record $_.Record
            if ($txtValue -and $txtValue -match '^(?i)v=DMARC1(?:;|\s|$)') {
                [PSCustomObject]@{
                    Value           = $txtValue
                    TTL             = $_.Record.TTL
                    QueryDurationMs = $_.QueryDurationMs
                }
            }
        } |
        Sort-Object Value, TTL -Unique

    # DKIM records (selector._domainkey CNAME) - query common selectors
    $dkimSelectors = @('default', 'selector1', 'selector2', 'google', 'mandrill')
    $dkimQueryTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $dkimRecords = [System.Collections.Generic.List[object]]::new()

    foreach ($selector in $dkimSelectors) {
        $dkimName = "${selector}._domainkey.$domain"
        $dkimResponses = Resolve-DnsAcrossServers -Name $dkimName -Type CNAME -Servers $queryServers -FallbackServer $DnsServer

        $dkimResponses |
            Where-Object { $_.Record.Type -eq 'CNAME' } |
            ForEach-Object {
                $dkimRecords.Add([PSCustomObject]@{
                    Selector        = $selector
                    Value           = $_.Record.NameHost.TrimEnd('.')
                    TTL             = $_.Record.TTL
                    QueryDurationMs = $_.QueryDurationMs
                })
            }
    }
    $dkimQueryTimer.Stop()
    $dkimRecordsSorted = $dkimRecords | Sort-Object Selector, Value, TTL -Unique

    $results.Add([PSCustomObject]@{
        Domain             = $domain
        QueriedNameServer  = $queriedNameServer
        QueryTimingMs      = [PSCustomObject]@{
            MX    = $mxQueryTimer.ElapsedMilliseconds
            TXT   = $txtQueryTimer.ElapsedMilliseconds
            DMARC = $dmarcQueryTimer.ElapsedMilliseconds
            DKIM  = $dkimQueryTimer.ElapsedMilliseconds
            Total = ($mxQueryTimer.ElapsedMilliseconds + $txtQueryTimer.ElapsedMilliseconds + $dmarcQueryTimer.ElapsedMilliseconds + $dkimQueryTimer.ElapsedMilliseconds)
        }
        MX                 = @($mxRecords)
        TXT                = @($txtRecords)
        DMARC              = @($dmarcRecords)
        DKIM               = @($dkimRecordsSorted)
    })
}

Write-Progress -Activity "Querying DNS records" -Completed

$results | ConvertTo-Json -Depth 5 | Set-Content -Path $OutputPath -Encoding UTF8

Write-Host "Done. $($results.Count) domain(s) processed." -ForegroundColor Green
Write-Host "Output: $OutputPath" -ForegroundColor Cyan
