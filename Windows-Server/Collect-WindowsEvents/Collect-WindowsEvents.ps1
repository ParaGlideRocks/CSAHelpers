<#
.SYNOPSIS
    Collects Windows event log entries from one or more remote servers based on a JSON configuration file.

.DESCRIPTION
    Reads a JSON configuration file that specifies servers, event logs, event IDs, and an optional
    timeframe. Queries each server using Get-WinEvent and exports results to a JSON output file.

    Filtering logic (any combination is supported):
    - If eventIds array is non-empty, filter by those IDs.
    - If timeframe startTime/endTime (or lastHours) are specified, apply time filtering.
    - If eventIds is empty or omitted, all events from that log within the timeframe are collected.

.PARAMETER ConfigPath
    Path to the JSON configuration file. Defaults to 'EventCollection-Config.json' in the same
    directory as this script.

.PARAMETER OutputPath
    Path for the JSON output file. Defaults to 'EventCollection-Results_<timestamp>.json' in the
    same directory as this script.

.PARAMETER Credential
    PSCredential to use when connecting to remote servers. If not supplied, the current user
    context is used.

.PARAMETER MaxEventsPerQuery
    Maximum number of events to retrieve per server/log/eventId combination. Default is 1000.

.EXAMPLE
    .\Collect-WindowsEvents.ps1

.EXAMPLE
    .\Collect-WindowsEvents.ps1 -ConfigPath "C:\Configs\myconfig.json" -OutputPath "C:\Results\events.json"

.EXAMPLE
    $cred = Get-Credential
    .\Collect-WindowsEvents.ps1 -Credential $cred -MaxEventsPerQuery 500
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'EventCollection-Config.json'),

    [Parameter()]
    [string]$OutputPath = (Join-Path $PSScriptRoot ("EventCollection-Results_{0}.json" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential,

    [Parameter()]
    [int]$MaxEventsPerQuery = 1000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

#region Helpers

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $color = switch ($Level) {
        'WARN'  { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Cyan' }
    }
    Write-Host "[$timestamp][$Level] $Message" -ForegroundColor $color
}

function Resolve-Timeframe {
    param([pscustomobject]$TimeframeCfg)

    $start = $null
    $end   = $null

    if ($null -ne $TimeframeCfg) {
        # lastHours takes precedence over explicit start/end
        if ($TimeframeCfg.lastHours -and $TimeframeCfg.lastHours -gt 0) {
            $end   = Get-Date
            $start = $end.AddHours(-[double]$TimeframeCfg.lastHours)
        }
        else {
            if ($TimeframeCfg.startTime) {
                $start = [datetime]::Parse($TimeframeCfg.startTime)
            }
            if ($TimeframeCfg.useCurrentTimeAsEnd -eq $true) {
                $end = Get-Date
            }
            elseif ($TimeframeCfg.endTime) {
                $end = [datetime]::Parse($TimeframeCfg.endTime)
            }
        }
    }

    return @{ Start = $start; End = $end }
}

function Build-FilterHashtable {
    param(
        [string]$LogName,
        [int[]]$EventIds,
        [datetime]$Start,
        [datetime]$End
    )

    $filter = @{ LogName = $LogName }

    if ($EventIds -and $EventIds.Count -gt 0) {
        $filter['Id'] = $EventIds
    }
    if ($Start) {
        $filter['StartTime'] = $Start
    }
    if ($End) {
        $filter['EndTime'] = $End
    }

    return $filter
}

function ConvertTo-EventObject {
    param([System.Diagnostics.Eventing.Reader.EventLogRecord]$Event, [string]$ServerName)

    # Safely retrieve message text — may fail for logs without a registered provider
    $message = ''
    try { $message = $Event.FormatDescription() } catch {}

    [PSCustomObject]@{
        Server      = $ServerName
        LogName     = $Event.LogName
        EventId     = $Event.Id
        Level       = $Event.LevelDisplayName
        TimeCreated = $Event.TimeCreated.ToString('o')
        Source      = $Event.ProviderName
        Message     = $message
        RecordId    = $Event.RecordId
        MachineName = $Event.MachineName
        UserId      = if ($Event.UserId) { $Event.UserId.Value } else { $null }
        Keywords    = $Event.KeywordsDisplayNames -join ', '
        TaskCategory = $Event.TaskDisplayName
    }
}

#endregion

#region Main

Write-Log "Starting Windows Event Collection"
Write-Log "Config  : $ConfigPath"
Write-Log "Output  : $OutputPath"

# --- Load and validate config ---
if (-not (Test-Path $ConfigPath)) {
    Write-Log "Config file not found: $ConfigPath" -Level 'ERROR'
    exit 1
}

try {
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
}
catch {
    Write-Log "Failed to parse config file: $_" -Level 'ERROR'
    exit 1
}

if (-not $config.servers -or $config.servers.Count -eq 0) {
    Write-Log "No servers defined in config." -Level 'WARN'
    exit 0
}

# --- Resolve global timeframe ---
$timeframe = Resolve-Timeframe -TimeframeCfg $config.timeframe

if ($timeframe.Start) { Write-Log "Time Start : $($timeframe.Start)" }
if ($timeframe.End)   { Write-Log "Time End   : $($timeframe.End)" }

# --- Prepare shared Get-WinEvent params ---
$winEventBase = @{ MaxEvents = $MaxEventsPerQuery }
if ($Credential) { $winEventBase['Credential'] = $Credential }

# --- Collect events ---
$allResults  = [System.Collections.Generic.List[PSCustomObject]]::new()
$serverStats = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($server in $config.servers) {
    $serverName   = $server.name
    $serverErrors = [System.Collections.Generic.List[string]]::new()
    $serverCount  = 0

    Write-Log "Processing server: $serverName"

    if (-not $server.eventLogs -or $server.eventLogs.Count -eq 0) {
        Write-Log "  No event logs configured for $serverName — skipping." -Level 'WARN'
        continue
    }

    foreach ($logCfg in $server.eventLogs) {
        $logName  = $logCfg.logName
        $eventIds = @()
        if ($logCfg.eventIds -and $logCfg.eventIds.Count -gt 0) {
            $eventIds = [int[]]$logCfg.eventIds
        }

        Write-Log ("  Log: {0} | IDs: {1}" -f $logName, (($eventIds.Count -gt 0) ? ($eventIds -join ',') : 'ALL'))

        $filter = Build-FilterHashtable `
            -LogName  $logName `
            -EventIds $eventIds `
            -Start    $timeframe.Start `
            -End      $timeframe.End

        $params = $winEventBase.Clone()
        $params['FilterHashtable'] = $filter

        # Local vs remote
        $isLocal = ($serverName -eq $env:COMPUTERNAME) -or ($serverName -eq 'localhost') -or ($serverName -eq '.')
        if (-not $isLocal) {
            $params['ComputerName'] = $serverName
        }

        try {
            $events = Get-WinEvent @params -ErrorAction Stop
            foreach ($evt in $events) {
                $allResults.Add((ConvertTo-EventObject -Event $evt -ServerName $serverName))
            }
            $count = $events.Count
            $serverCount += $count
            Write-Log ("    -> {0} event(s) collected." -f $count)
        }
        catch [System.Exception] {
            $errMsg = $_.Exception.Message
            # "No events found" is informational, not an error
            if ($errMsg -match 'No events were found') {
                Write-Log "    -> No events matched the criteria." -Level 'WARN'
            }
            else {
                Write-Log ("    -> ERROR querying {0}\{1}: {2}" -f $serverName, $logName, $errMsg) -Level 'ERROR'
                $serverErrors.Add("[$logName] $errMsg")
            }
        }
    }

    $serverStats.Add([PSCustomObject]@{
        Server     = $serverName
        EventCount = $serverCount
        Errors     = if ($serverErrors.Count -gt 0) { $serverErrors -join '; ' } else { $null }
    })
}

# --- Build output document ---
$output = [PSCustomObject]@{
    CollectionMetadata = [PSCustomObject]@{
        GeneratedAt     = (Get-Date -Format 'o')
        ConfigFile      = $ConfigPath
        TimeframeStart  = if ($timeframe.Start) { $timeframe.Start.ToString('o') } else { $null }
        TimeframeEnd    = if ($timeframe.End)   { $timeframe.End.ToString('o')   } else { $null }
        TotalEvents     = $allResults.Count
        ServerSummary   = $serverStats
    }
    Events = $allResults
}

# --- Write output ---
try {
    $output | ConvertTo-Json -Depth 10 | Set-Content -Path $OutputPath -Encoding UTF8
    Write-Log "Results written to: $OutputPath"
    Write-Log ("Total events collected: {0}" -f $allResults.Count)
}
catch {
    Write-Log "Failed to write output file: $_" -Level 'ERROR'
    exit 1
}

Write-Log "Collection complete."

#endregion
