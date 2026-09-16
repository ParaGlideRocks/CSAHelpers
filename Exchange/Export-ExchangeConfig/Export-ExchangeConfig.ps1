#Requires -Version 5.1
<#
.SYNOPSIS
    Exports Exchange on-premises configuration to JSON and optionally renders an HTML report.

.DESCRIPTION
    Connects to Exchange on-premises servers and dumps selected configuration categories
    (accepted domains, adpermissions, adserversettings, authconfig, authenticationpolicy,
    databaseavailabilitygroup*, *virtualdirectory, federation*, foreignconnector,
    frontendtransportservice, hybridconfiguration, settings, *connector, transport*)
    into a structured JSON file, and optionally produces a self-contained HTML report.

.PARAMETER Servers
    Array of Exchange server hostnames or FQDNs to query.
    Not required when -JsonInputFile is used alone.

.PARAMETER Categories
    One or more configuration categories to export, or "All".
    Not required when -JsonInputFile is used alone.

.PARAMETER OutputPath
    Folder where JSON / HTML output files will be saved. Defaults to the current directory.

.PARAMETER LogPath
    Path to the log file. Defaults to .\ExchangeConfigDump_<timestamp>.log.

.PARAMETER Credential
    PSCredential for Exchange remote PowerShell (falls back to Kerberos if omitted).

.PARAMETER ExchangeUri
    Override the remote PowerShell URI (default: http://<first server>/PowerShell/).

.PARAMETER GenerateHtmlReport
    Switch: generate an HTML report after the JSON export (or from -JsonInputFile).

.PARAMETER JsonInputFile
    Path to an existing JSON dump produced by this script.
    When specified together with -GenerateHtmlReport, only the HTML report is built
    (no Exchange connection is made).

.EXAMPLE
    # Full export + HTML report
    .\Export-ExchangeConfig.ps1 -Servers "EX01","EX02" -Categories All -GenerateHtmlReport

.EXAMPLE
    # Export specific categories only, no report
    .\Export-ExchangeConfig.ps1 -Servers "EX01" -Categories AcceptedDomains,Connectors,Transport `
        -OutputPath "C:\Reports" -Credential (Get-Credential)

.EXAMPLE
    # Build HTML report from an existing JSON dump (no Exchange connection needed)
    .\Export-ExchangeConfig.ps1 -GenerateHtmlReport `
        -JsonInputFile "C:\Reports\ExchangeConfigDump_20240101_120000.json"

.NOTES
    Author  : Exchange Admin Toolkit
    Requires: Exchange Management Shell or remote PowerShell access to Exchange.

    Version history:
      1.0.0 - Initial release. JSON export of all Exchange config categories.
      2.0.0 - Added HTML report generation (-GenerateHtmlReport / -JsonInputFile).
              Added ReportOnly parameter set (no Exchange connection needed).
              Lazy on-demand rendering: JSON embedded as JS variable, each
              category rendered client-side only when selected.
      3.0.0 - Fixed horizontal scrolling (overflow-x on body, #main, .cmdlet-block).
              Updated version metadata in script header and runtime log messages.
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = "Export")]
param (
    [Parameter(ParameterSetName = "Export",          Mandatory)]
    [Parameter(ParameterSetName = "ExportAndReport", Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string[]] $Servers,

    [Parameter(ParameterSetName = "Export",          Mandatory)]
    [Parameter(ParameterSetName = "ExportAndReport", Mandatory)]
    [ValidateSet(
        "All","AcceptedDomains","ADPermissions","ADServerSettings","AuthConfig",
        "AuthenticationPolicy","DatabaseAvailabilityGroup","VirtualDirectories",
        "Federation","ForeignConnector","FrontendTransportService",
        "HybridConfiguration","Settings","Connectors","Transport"
    )]
    [string[]] $Categories,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $OutputPath = (Get-Location).Path,

    [Parameter()]
    [string] $LogPath,

    [Parameter()]
    [System.Management.Automation.PSCredential] $Credential,

    [Parameter()]
    [string] $ExchangeUri,

    [Parameter(ParameterSetName = "ExportAndReport", Mandatory)]
    [Parameter(ParameterSetName = "ReportOnly",      Mandatory)]
    [switch] $GenerateHtmlReport,

    [Parameter(ParameterSetName = "ReportOnly", Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string] $JsonInputFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ===========================================================================
#  REGION: Logging
# ===========================================================================
$script:LogFile = if ($LogPath) { $LogPath } else {
    Join-Path $OutputPath ("ExchangeConfigDump_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Message,
        [ValidateSet("INFO","WARN","ERROR","DEBUG","SUCCESS")][string] $Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry     = "[$timestamp] [$Level] $Message"
    $color     = switch ($Level) {
        "INFO"    { "Cyan"   }
        "WARN"    { "Yellow" }
        "ERROR"   { "Red"    }
        "DEBUG"   { "Gray"   }
        "SUCCESS" { "Green"  }
        default   { "White"  }
    }
    Write-Host $entry -ForegroundColor $color
    $logDir = Split-Path $script:LogFile -Parent
    if ($logDir -and -not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    Add-Content -Path $script:LogFile -Value $entry -Encoding UTF8
}

function Write-LogInfo    { param([string]$m) Write-Log -Message $m -Level "INFO"    }
function Write-LogWarn    { param([string]$m) Write-Log -Message $m -Level "WARN"    }
function Write-LogError   { param([string]$m) Write-Log -Message $m -Level "ERROR"   }
function Write-LogDebug   { param([string]$m) Write-Log -Message $m -Level "DEBUG"   }
function Write-LogSuccess { param([string]$m) Write-Log -Message $m -Level "SUCCESS" }

# ===========================================================================
#  REGION: Helpers
# ===========================================================================

function ConvertTo-SafeObject {
    param([object]$InputObject, [int]$Depth = 0, [int]$MaxDepth = 8)
    if ($null -eq $InputObject) { return $null }
    if ($Depth -gt $MaxDepth)   { return $InputObject.ToString() }
    if ($InputObject -is [string]  -or $InputObject -is [bool]    -or
        $InputObject -is [int]     -or $InputObject -is [long]    -or
        $InputObject -is [double]  -or $InputObject -is [decimal] -or
        $InputObject -is [datetime]) { return $InputObject }
    if ($InputObject -is [System.Enum]) { return $InputObject.ToString() }
    if ($InputObject -is [System.Collections.IEnumerable] -and
        $InputObject -isnot [string] -and
        $InputObject -isnot [System.Collections.IDictionary]) {
        $list = [System.Collections.Generic.List[object]]::new()
        foreach ($item in $InputObject) {
            $list.Add((ConvertTo-SafeObject $item -Depth ($Depth+1) -MaxDepth $MaxDepth))
        }
        return $list.ToArray()
    }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $ht = [ordered]@{}
        foreach ($key in $InputObject.Keys) {
            $ht[$key.ToString()] = ConvertTo-SafeObject $InputObject[$key] -Depth ($Depth+1) -MaxDepth $MaxDepth
        }
        return $ht
    }
    if ($InputObject -is [System.Management.Automation.PSObject] -or
        $InputObject.GetType().FullName -match "^Microsoft\.Exchange") {
        $ht = [ordered]@{}
        foreach ($prop in $InputObject.PSObject.Properties) {
            try   { $ht[$prop.Name] = ConvertTo-SafeObject $prop.Value -Depth ($Depth+1) -MaxDepth $MaxDepth }
            catch { $ht[$prop.Name] = "<serialisation error: $($_.Exception.Message)>" }
        }
        return $ht
    }
    return $InputObject.ToString()
}

function Invoke-ExchangeCommand {
    param(
        [string]    $CmdletName,
        [hashtable] $Parameters  = @{},
        [string]    $ServerName  = "",
        [switch]    $PerServer
    )
    $label = if ($ServerName) { "[$ServerName] " } else { "" }
    Write-LogDebug "${label}Running: $CmdletName $(($Parameters.GetEnumerator() | ForEach-Object { "-$($_.Key) $($_.Value)" }) -join ' ')"
    try {
        $cmd    = Get-Command $CmdletName -ErrorAction Stop
        $result = & $cmd @Parameters
        if ($null -eq $result) { Write-LogDebug "${label}$CmdletName returned no data."; return @() }
        return $result
    }
    catch [System.Management.Automation.CommandNotFoundException] {
        Write-LogWarn "${label}Cmdlet '$CmdletName' not found – skipping."
        return $null
    }
    catch {
        Write-LogError "${label}Error running '$CmdletName': $($_.Exception.Message)"
        return $null
    }
}

# ===========================================================================
#  REGION: Category definitions
# ===========================================================================
$CategoryMap = [ordered]@{
    AcceptedDomains = @(
        @{ Cmdlet = "Get-AcceptedDomain"; PerServer = $false }
    )
    ADPermissions = @(
        @{ Cmdlet = "Get-ADPermission"; PerServer = $true; ServerParam = "Identity" }
    )
    ADServerSettings = @(
        @{ Cmdlet = "Get-ADServerSettings"; PerServer = $false }
    )
    AuthConfig = @(
        @{ Cmdlet = "Get-AuthConfig"; PerServer = $false }
    )
    AuthenticationPolicy = @(
        @{ Cmdlet = "Get-AuthenticationPolicy"; PerServer = $false }
    )
    DatabaseAvailabilityGroup = @(
        @{ Cmdlet = "Get-DatabaseAvailabilityGroup";        PerServer = $false }
        @{ Cmdlet = "Get-DatabaseAvailabilityGroupNetwork"; PerServer = $false }
        @{ Cmdlet = "Get-MailboxDatabase";                  PerServer = $false }
        @{ Cmdlet = "Get-MailboxDatabaseCopyStatus";        PerServer = $true; ServerParam = "Server" }
    )
    VirtualDirectories = @(
        @{ Cmdlet = "Get-OwaVirtualDirectory";          PerServer = $true; ServerParam = "Server" }
        @{ Cmdlet = "Get-EcpVirtualDirectory";          PerServer = $true; ServerParam = "Server" }
        @{ Cmdlet = "Get-WebServicesVirtualDirectory";  PerServer = $true; ServerParam = "Server" }
        @{ Cmdlet = "Get-ActiveSyncVirtualDirectory";   PerServer = $true; ServerParam = "Server" }
        @{ Cmdlet = "Get-OabVirtualDirectory";          PerServer = $true; ServerParam = "Server" }
        @{ Cmdlet = "Get-AutodiscoverVirtualDirectory"; PerServer = $true; ServerParam = "Server" }
        @{ Cmdlet = "Get-MapiVirtualDirectory";         PerServer = $true; ServerParam = "Server" }
        @{ Cmdlet = "Get-OutlookAnywhere";              PerServer = $true; ServerParam = "Server" }
        @{ Cmdlet = "Get-PowerShellVirtualDirectory";   PerServer = $true; ServerParam = "Server" }
    )
    Federation = @(
        @{ Cmdlet = "Get-FederationTrust";                 PerServer = $false }
        @{ Cmdlet = "Get-FederatedOrganizationIdentifier"; PerServer = $false }
        @{ Cmdlet = "Get-FederationInformation";           PerServer = $false }
        @{ Cmdlet = "Get-OrganizationRelationship";        PerServer = $false }
        @{ Cmdlet = "Get-SharingPolicy";                   PerServer = $false }
    )
    ForeignConnector = @(
        @{ Cmdlet = "Get-ForeignConnector"; PerServer = $false }
    )
    FrontendTransportService = @(
        @{ Cmdlet = "Get-FrontendTransportService"; PerServer = $true; ServerParam = "Identity" }
    )
    HybridConfiguration = @(
        @{ Cmdlet = "Get-HybridConfiguration";            PerServer = $false }
        @{ Cmdlet = "Get-OnPremisesOrganization";         PerServer = $false }
        @{ Cmdlet = "Get-IntraOrganizationConnector";     PerServer = $false }
        @{ Cmdlet = "Get-IntraOrganizationConfiguration"; PerServer = $false }
    )
    Settings = @(
        @{ Cmdlet = "Get-ExchangeServer";                  PerServer = $true; ServerParam = "Identity" }
        @{ Cmdlet = "Get-OrganizationConfig";              PerServer = $false }
        @{ Cmdlet = "Get-AdminAuditLogConfig";             PerServer = $false }
        @{ Cmdlet = "Get-ExchangeDiagnosticInfo";          PerServer = $true; ServerParam = "Server";
           Parameters = @{ Process = "EdgeTransport"; Component = "ResourceThrottling" } }
        @{ Cmdlet = "Get-SettingOverride";                 PerServer = $false }
        @{ Cmdlet = "Get-ExchangeServerAccessLicenseUser"; PerServer = $false }
    )
    Connectors = @(
        @{ Cmdlet = "Get-ReceiveConnector"; PerServer = $true; ServerParam = "Server" }
        @{ Cmdlet = "Get-SendConnector";    PerServer = $false }
    )
    Transport = @(
        @{ Cmdlet = "Get-TransportConfig";   PerServer = $false }
        @{ Cmdlet = "Get-TransportRule";     PerServer = $false }
        @{ Cmdlet = "Get-TransportService";  PerServer = $true; ServerParam = "Identity" }
        @{ Cmdlet = "Get-TransportAgent";    PerServer = $false }
        @{ Cmdlet = "Get-TransportPipeline"; PerServer = $false }
    )
}

# ===========================================================================
#  REGION: Session management
# ===========================================================================
$script:ExchangeSession = $null

function Connect-ExchangeRemote {
    param([string]$ServerName)
    $uri = if ($ExchangeUri) { $ExchangeUri } else { "http://$ServerName/PowerShell/" }
    Write-LogInfo "Connecting to Exchange remote PowerShell at: $uri"
    $sessionParams = @{
        ConfigurationName = "Microsoft.Exchange"
        ConnectionUri     = $uri
        Authentication    = "Kerberos"
        AllowRedirection  = $true
        SessionOption     = (New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck)
        ErrorAction       = "Stop"
    }
    if ($Credential) {
        $sessionParams["Credential"]     = $Credential
        $sessionParams["Authentication"] = "Basic"
    }
    try {
        $session = New-PSSession @sessionParams
        Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
        $script:ExchangeSession = $session
        Write-LogSuccess "Exchange session established (SessionId=$($session.Id))."
    }
    catch {
        Write-LogError "Failed to create Exchange remote PowerShell session: $($_.Exception.Message)"
        throw
    }
}

function Disconnect-ExchangeRemote {
    if ($script:ExchangeSession) {
        try   { Remove-PSSession $script:ExchangeSession -ErrorAction SilentlyContinue; Write-LogInfo "Exchange session closed." }
        catch { Write-LogWarn "Could not cleanly close the Exchange session: $($_.Exception.Message)" }
        $script:ExchangeSession = $null
    }
}

function Test-ExchangeCmdletsAvailable {
    try { $null = Get-Command "Get-ExchangeServer" -ErrorAction Stop; return $true }
    catch { return $false }
}

# ===========================================================================
#  REGION: Core export logic
# ===========================================================================

function Export-CategoryData {
    param([string]$CategoryName, [string[]]$TargetServers)
    Write-LogInfo "--- Processing category: $CategoryName ---"
    $categoryResult = [ordered]@{}
    foreach ($cmdDef in $CategoryMap[$CategoryName]) {
        $cmdlet      = $cmdDef.Cmdlet
        $perServer   = if (([System.Collections.IDictionary]$cmdDef).Contains('PerServer')   -and $cmdDef.PerServer)   { $true } else { $false }
        $srvParam    = if (([System.Collections.IDictionary]$cmdDef).Contains('ServerParam') -and $cmdDef.ServerParam) { $cmdDef.ServerParam } else { "Server" }
        $extraParams = if (([System.Collections.IDictionary]$cmdDef).Contains('Parameters')  -and $cmdDef.Parameters)  { $cmdDef.Parameters  } else { @{} }
        $cmdletResult = [ordered]@{}
        if ($perServer) {
            foreach ($srv in $TargetServers) {
                Write-LogInfo "  [$srv] $cmdlet"
                $params = $extraParams.Clone(); $params[$srvParam] = $srv
                $raw = Invoke-ExchangeCommand -CmdletName $cmdlet -Parameters $params -ServerName $srv -PerServer
                $cmdletResult[$srv] = if ($null -ne $raw) { ConvertTo-SafeObject $raw } else { $null }
            }
        }
        else {
            Write-LogInfo "  [org-wide] $cmdlet"
            $raw = Invoke-ExchangeCommand -CmdletName $cmdlet -Parameters $extraParams
            $cmdletResult["__OrgWide__"] = if ($null -ne $raw) { ConvertTo-SafeObject $raw } else { $null }
        }
        $categoryResult[$cmdlet] = $cmdletResult
    }
    return $categoryResult
}

# ===========================================================================
#  REGION: HTML Report
# ===========================================================================

function New-HtmlReport {
    <#
    .SYNOPSIS
        Builds a self-contained, lazily-rendered HTML report from a JSON dump.
        All data is embedded as a JS variable; each category is rendered on demand
        when the user selects it, keeping initial page load fast regardless of
        how large the JSON is.
    #>
    param([string]$JsonFilePath, [string]$DestinationFolder)

    Write-LogInfo "Building HTML report from: $JsonFilePath"
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

    # ── Parse JSON ─────────────────────────────────────────────────────────
    try {
        $raw  = [System.IO.File]::ReadAllText($JsonFilePath, [System.Text.Encoding]::UTF8)
        $data = $raw | ConvertFrom-Json
    }
    catch {
        Write-LogError "Failed to parse JSON file: $($_.Exception.Message)"
        throw
    }

    # ── Metadata for sidebar ───────────────────────────────────────────────
    $meta       = $data.ExportMetadata
    $exportDate = if ($null -ne $meta -and $null -ne $meta.ExportDate)    { [System.Web.HttpUtility]::HtmlEncode($meta.ExportDate) }    else { "N/A" }
    $exportHost = if ($null -ne $meta -and $null -ne $meta.ExportHost)    { [System.Web.HttpUtility]::HtmlEncode($meta.ExportHost) }    else { "N/A" }
    $exportUser = if ($null -ne $meta -and $null -ne $meta.ExportUser)    { [System.Web.HttpUtility]::HtmlEncode($meta.ExportUser) }    else { "N/A" }
    $scriptVer  = if ($null -ne $meta -and $null -ne $meta.ScriptVersion) { [System.Web.HttpUtility]::HtmlEncode($meta.ScriptVersion) } else { "N/A" }
    $servers    = if ($null -ne $meta -and $null -ne $meta.TargetServers) { [System.Web.HttpUtility]::HtmlEncode((@($meta.TargetServers) -join ", ")) } else { "N/A" }
    $errCount   = if ($null -ne $meta -and $null -ne $meta.CategoriesWithErrors) { [int]$meta.CategoriesWithErrors } else { 0 }
    $statusCls  = if ($errCount -gt 0) { "status-warn" } else { "status-ok" }
    $statusTxt  = if ($errCount -gt 0) { "&#9888; $errCount error(s)" } else { "&#10004; No errors" }

    # ── Build sidebar nav items (just links — no content rendered yet) ─────
    $navHtml  = [System.Text.StringBuilder]::new()
    $catNames = @($data.Data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)
    foreach ($cat in $catNames) {
        $esc = [System.Web.HttpUtility]::HtmlEncode($cat)
        [void]$navHtml.Append("<li><a class=`"nav-link`" href=`"#`" onclick=`"loadCategory(event,'$esc')`">$esc</a></li>`n")
    }

    # ── Errors sidebar widget ──────────────────────────────────────────────
    $errListHtml = ""
    if ($null -ne $meta -and $null -ne $meta.Errors -and @($meta.Errors).Count -gt 0) {
        $items = (@($meta.Errors) | ForEach-Object { "<li>$([System.Web.HttpUtility]::HtmlEncode($_))</li>" }) -join "`n"
        $errListHtml = "<div class=`"sidebar-errors`"><strong>Export errors</strong><ul>$items</ul></div>"
    }

    # ── Embed the raw JSON (re-serialised for safety) ──────────────────────
    # We convert to JSON again to ensure it is compact and valid for JS embedding.
    # Single-quotes in the JSON are escaped; the whole thing goes into a JS const.
    $jsonForJs = ($data.Data | ConvertTo-Json -Depth 20 -Compress -WarningAction SilentlyContinue) -replace '</script>', '<\/script>'

    # ── Assemble the HTML shell ─────────────────────────────────────────────
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Exchange Configuration Report</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#f4f6f9;--surface:#fff;--sidebar-bg:#1a2236;--sidebar-fg:#c9d1e0;
  --accent:#0078d4;--accent2:#005ea2;--border:#dde3ed;
  --text:#1e2a3a;--muted:#6b7a99;--ok:#107c10;--warn:#b86800;--err:#c42b1c;
  --radius:6px;--font:'Segoe UI',system-ui,sans-serif;--mono:'Cascadia Code','Consolas',monospace;
}
[data-theme=dark]{
  --bg:#111827;--surface:#1f2937;--border:#374151;
  --text:#f1f5f9;--muted:#94a3b8;--sidebar-bg:#0d1117;--sidebar-fg:#94a3b8;
}
html{scroll-behavior:smooth}
body{font-family:var(--font);background:var(--bg);color:var(--text);display:flex;min-height:100vh;font-size:14px;overflow-x:auto}

/* Sidebar */
#sidebar{width:260px;min-width:260px;background:var(--sidebar-bg);color:var(--sidebar-fg);display:flex;flex-direction:column;position:sticky;top:0;height:100vh;overflow-y:auto;padding-bottom:24px;z-index:10}
.s-head{padding:20px 18px 14px;border-bottom:1px solid rgba(255,255,255,.08)}
.s-head h1{font-size:13px;font-weight:700;color:#fff;line-height:1.4}
.s-head p{font-size:11px;margin-top:3px}
.s-meta{padding:12px 18px;font-size:11px;line-height:1.9;border-bottom:1px solid rgba(255,255,255,.06)}
.s-meta span{color:#fff;font-weight:600}
.status-badge{display:inline-block;margin-top:5px;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700}
.status-ok{background:rgba(16,124,16,.25);color:#6fcf97}
.status-warn{background:rgba(184,104,0,.25);color:#f2c94c}
.s-search{padding:10px 14px}
.s-search input{width:100%;padding:7px 10px;border-radius:var(--radius);border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.07);color:#fff;font-size:12px;outline:none}
.s-search input::placeholder{color:rgba(255,255,255,.35)}
.s-nav{flex:1;padding:4px 0}
.s-nav ul{list-style:none}
.nav-link{display:block;padding:7px 18px;font-size:12px;color:var(--sidebar-fg);text-decoration:none;border-left:3px solid transparent;transition:background .12s,border-color .12s;cursor:pointer}
.nav-link:hover,.nav-link.active{background:rgba(255,255,255,.07);border-left-color:var(--accent);color:#fff}
.sidebar-errors{margin:10px 14px 0;padding:10px 12px;background:rgba(196,43,28,.15);border-radius:var(--radius);font-size:11px}
.sidebar-errors strong{color:#ff8a80;display:block;margin-bottom:5px}
.sidebar-errors ul{padding-left:14px}
.sidebar-errors li{margin-bottom:3px;color:#ffb3ae}

/* Main */
#main{flex:1;overflow-x:auto;padding:0 28px 60px;min-width:0}

/* Topbar */
.topbar{position:sticky;top:0;z-index:5;background:var(--surface);border-bottom:1px solid var(--border);padding:10px 0;margin-bottom:20px;display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap}
.topbar-title{font-size:15px;font-weight:700}
.topbar-right{display:flex;gap:6px;align-items:center;flex-wrap:wrap}
.btn{padding:6px 12px;border-radius:var(--radius);border:1px solid var(--border);background:var(--surface);color:var(--text);font-size:12px;cursor:pointer;font-family:var(--font);transition:background .12s;white-space:nowrap}
.btn:hover{background:var(--bg)}
.btn-primary{background:var(--accent);color:#fff;border-color:var(--accent)}
.btn-primary:hover{background:var(--accent2)}

/* Welcome / placeholder */
#welcome{padding:40px 0;color:var(--muted);font-size:14px}
#welcome h2{font-size:22px;font-weight:700;color:var(--text);margin-bottom:10px}
#welcome p{margin-bottom:6px}
.cat-grid{display:flex;flex-wrap:wrap;gap:10px;margin-top:20px}
.cat-pill{padding:6px 14px;border-radius:20px;background:var(--surface);border:1px solid var(--border);font-size:12px;cursor:pointer;transition:background .12s,border-color .12s}
.cat-pill:hover{background:var(--accent);color:#fff;border-color:var(--accent)}

/* Category view */
.cat-title{font-size:20px;font-weight:700;border-bottom:2px solid var(--accent);padding-bottom:7px;margin-bottom:16px}
.breadcrumb{font-size:12px;color:var(--muted);margin-bottom:14px}
.breadcrumb a{color:var(--accent);text-decoration:none;cursor:pointer}
.breadcrumb a:hover{text-decoration:underline}

/* Cmdlet block */
.cmdlet-block{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);margin-bottom:10px;overflow:visible}
.cmdlet-toggle{width:100%;text-align:left;padding:11px 14px;background:none;border:none;cursor:pointer;font-family:var(--mono);font-size:13px;font-weight:600;color:var(--accent);display:flex;align-items:center;gap:8px}
.cmdlet-toggle:hover{background:rgba(0,120,212,.04)}
.chevron{transition:transform .18s;display:inline-block}
.cmdlet-toggle[aria-expanded=false] .chevron{transform:rotate(-90deg)}
.cmdlet-body{padding:2px 14px 12px}

/* Scope */
.scope-block{margin-bottom:10px}
.scope-label{display:inline-block;font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px;margin:7px 0 5px;letter-spacing:.3px;text-transform:uppercase}
.scope-org{background:#e8f4fd;color:#0550ae}
.scope-server{background:#f0fdf4;color:#166534}
[data-theme=dark] .scope-org{background:rgba(5,80,174,.3);color:#7dd3fc}
[data-theme=dark] .scope-server{background:rgba(22,101,52,.3);color:#86efac}

/* Tables */
.scroll-x{overflow-x:auto;border-radius:var(--radius);border:1px solid var(--border)}
.data-table,.props-table{width:100%;border-collapse:collapse;font-size:12px}
.data-table th{background:var(--bg);text-align:left;padding:7px 10px;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px;color:var(--muted);border-bottom:2px solid var(--border);white-space:nowrap;cursor:pointer;user-select:none}
.data-table th:hover{color:var(--accent)}
.sort-icon{font-size:10px;opacity:.5}
.data-table td,.props-table td{padding:6px 10px;border-bottom:1px solid var(--border);vertical-align:top}
.data-table tr:last-child td,.props-table tr:last-child td{border-bottom:none}
.data-table tr:hover td{background:rgba(0,120,212,.03)}
.props-table{border:1px solid var(--border);border-radius:var(--radius)}
.prop-key{font-family:var(--mono);font-size:11px;color:var(--muted);white-space:nowrap;width:210px}
.nested-obj{background:var(--bg);border:1px solid var(--border);border-radius:var(--radius);padding:5px;margin:2px 0}

/* Value decorators */
.null{color:var(--muted);font-style:italic;font-size:11px}
.badge{display:inline-block;padding:1px 7px;border-radius:10px;background:#e9eef6;color:var(--text);font-size:11px;margin:1px}
[data-theme=dark] .badge{background:#374151;color:#d1d5db}
.bool{font-weight:700;font-family:var(--mono);font-size:12px}
.bool-true{color:var(--ok)}
.bool-false{color:var(--err)}
.no-data{color:var(--muted);font-style:italic;font-size:12px;padding:4px 0}
.error-msg{color:var(--err);font-size:12px;padding:4px 0;font-weight:600}
.loading{color:var(--muted);padding:30px 0;font-style:italic}

/* Search highlight */
mark{background:#fff3b0;color:inherit;border-radius:2px;padding:0 1px}
[data-theme=dark] mark{background:#854d0e}

@media print{
  #sidebar{display:none}
  .topbar-right{display:none}
  #main{padding:20px}
  .cmdlet-body{display:block!important}
}
</style>
</head>
<body data-theme="light">

<nav id="sidebar" aria-label="Category navigation">
  <div class="s-head">
    <h1>&#9881; Exchange Config Report</h1>
    <p>Export-ExchangeConfig.ps1 v$scriptVer</p>
  </div>
  <div class="s-meta">
    <div>Date: <span>$exportDate</span></div>
    <div>Host: <span>$exportHost</span></div>
    <div>User: <span>$exportUser</span></div>
    <div>Servers: <span>$servers</span></div>
    <div><span class="status-badge $statusCls">$statusTxt</span></div>
  </div>
  <div class="s-search">
    <input id="sidebarSearch" type="text" placeholder="Filter categories..." oninput="filterNav(this.value)" aria-label="Filter categories"/>
  </div>
  <div class="s-nav"><ul id="navList">
$($navHtml.ToString())
  </ul></div>
  $errListHtml
</nav>

<div id="main">
  <div class="topbar">
    <span class="topbar-title" id="pageTitle">Exchange On-Premises Configuration</span>
    <div class="topbar-right">
      <button class="btn" id="expandAllBtn" onclick="expandAll()" style="display:none">Expand all</button>
      <button class="btn" id="collapseAllBtn" onclick="collapseAll()" style="display:none">Collapse all</button>
      <button class="btn btn-primary" id="themeBtn" onclick="toggleTheme()">&#127769; Dark mode</button>
      <button class="btn" onclick="window.print()">&#128424; Print</button>
    </div>
  </div>
  <div id="content">
    <div id="welcome">
      <h2>Select a category</h2>
      <p>Choose a category from the left sidebar to view its configuration data.</p>
      <p style="color:var(--muted);font-size:12px">Data is rendered on demand — only the selected category is loaded at a time.</p>
      <div class="cat-grid" id="catGrid"></div>
    </div>
  </div>
</div>

<script>
// ── Embedded data ─────────────────────────────────────────────────────────
const DATA = $jsonForJs;

// ── Populate welcome grid ─────────────────────────────────────────────────
(function(){
  const grid = document.getElementById('catGrid');
  Object.keys(DATA).forEach(cat => {
    const btn = document.createElement('button');
    btn.className = 'cat-pill';
    btn.textContent = cat;
    btn.onclick = () => loadCategoryByName(cat);
    grid.appendChild(btn);
  });
})();

// ── Nav helpers ───────────────────────────────────────────────────────────
function loadCategory(e, name){ e.preventDefault(); loadCategoryByName(name); }
function filterNav(v){
  v = v.toLowerCase();
  document.querySelectorAll('#navList li').forEach(li => {
    li.style.display = li.textContent.toLowerCase().includes(v) ? '' : 'none';
  });
}

// ── Lazy category renderer ────────────────────────────────────────────────
function loadCategoryByName(catName) {
  // Update nav active state
  document.querySelectorAll('.nav-link').forEach(a => {
    a.classList.toggle('active', a.textContent.trim() === catName);
  });

  document.getElementById('pageTitle').textContent = catName;
  document.getElementById('expandAllBtn').style.display   = '';
  document.getElementById('collapseAllBtn').style.display = '';

  const content = document.getElementById('content');
  content.innerHTML = '<p class="loading">&#8987; Rendering ' + esc(catName) + '...</p>';

  // Yield to browser to paint the loading indicator, then render
  setTimeout(() => {
    const catData = DATA[catName];
    if (!catData) { content.innerHTML = '<p class="error-msg">Category not found in data.</p>'; return; }

    const frag = document.createDocumentFragment();

    // Breadcrumb
    const bc = document.createElement('div');
    bc.className = 'breadcrumb';
    bc.innerHTML = '<a onclick="showWelcome()">&#8962; Home</a> &rsaquo; ' + esc(catName);
    frag.appendChild(bc);

    // Title
    const h2 = document.createElement('h2');
    h2.className = 'cat-title';
    h2.textContent = catName;
    frag.appendChild(h2);

    // Cmdlets
    Object.keys(catData).forEach(cmdlet => {
      const cmdData = catData[cmdlet];
      const block   = document.createElement('div');
      block.className = 'cmdlet-block';

      const btn = document.createElement('button');
      btn.className = 'cmdlet-toggle';
      btn.setAttribute('aria-expanded','true');
      btn.onclick = () => toggleBlock(btn);
      btn.innerHTML = '<span class="chevron">&#9660;</span> ' + esc(cmdlet);
      block.appendChild(btn);

      const body = document.createElement('div');
      body.className = 'cmdlet-body';

      if (!cmdData || Object.keys(cmdData).length === 0) {
        body.innerHTML = '<p class="no-data">No data.</p>';
      } else {
        Object.keys(cmdData).forEach(scope => {
          const scopeData  = cmdData[scope];
          const scopeLabel = scope === '__OrgWide__' ? 'Organisation-wide' : esc(scope);
          const scopeClass = scope === '__OrgWide__' ? 'scope-org' : 'scope-server';

          const sb = document.createElement('div');
          sb.className = 'scope-block';
          sb.innerHTML = '<div class="scope-label ' + scopeClass + '">' + scopeLabel + '</div>';

          if (scopeData === null || scopeData === undefined) {
            sb.innerHTML += '<p class="no-data">No data returned.</p>';
          } else if (scopeData && scopeData.__Error__) {
            sb.innerHTML += '<p class="error-msg">&#9940; ' + esc(String(scopeData.__Error__)) + '</p>';
          } else {
            sb.appendChild(renderValue(scopeData, 0));
          }
          body.appendChild(sb);
        });
      }

      block.appendChild(body);
      frag.appendChild(block);
    });

    content.innerHTML = '';
    content.appendChild(frag);
  }, 20);
}

function showWelcome(){
  document.querySelectorAll('.nav-link').forEach(a => a.classList.remove('active'));
  document.getElementById('pageTitle').textContent = 'Exchange On-Premises Configuration';
  document.getElementById('expandAllBtn').style.display   = 'none';
  document.getElementById('collapseAllBtn').style.display = 'none';
  document.getElementById('content').innerHTML = document.getElementById('welcomeTpl').innerHTML;
}

// ── Value renderer ────────────────────────────────────────────────────────
function renderValue(val, depth) {
  const wrap = document.createElement('span');

  if (val === null || val === undefined) {
    wrap.innerHTML = '<span class="null">null</span>'; return wrap;
  }

  // Object
  if (typeof val === 'object' && !Array.isArray(val)) {
    const keys = Object.keys(val);
    if (keys.length === 0) { wrap.innerHTML = '<span class="null">{ }</span>'; return wrap; }

    const table = document.createElement('table');
    table.className = 'props-table';
    const tbody = document.createElement('tbody');

    keys.forEach(k => {
      const tr = document.createElement('tr');
      const td1 = document.createElement('td'); td1.className = 'prop-key'; td1.textContent = k;
      const td2 = document.createElement('td'); td2.className = 'prop-val';
      td2.appendChild(renderValue(val[k], depth + 1));
      tr.appendChild(td1); tr.appendChild(td2); tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    if (depth > 0) { const d = document.createElement('div'); d.className = 'nested-obj'; d.appendChild(table); wrap.appendChild(d); }
    else { wrap.appendChild(table); }
    return wrap;
  }

  // Array
  if (Array.isArray(val)) {
    if (val.length === 0) { wrap.innerHTML = '<span class="null">[ ]</span>'; return wrap; }
    const isObjArray = val.length > 0 && typeof val[0] === 'object' && val[0] !== null && !Array.isArray(val[0]);
    if (isObjArray) {
      // Collect all unique keys
      const keys = [];
      val.forEach(item => { if(item && typeof item === 'object') Object.keys(item).forEach(k => { if (!keys.includes(k)) keys.push(k); }); });
      const outer = document.createElement('div'); outer.className = 'scroll-x';
      const table = document.createElement('table'); table.className = 'data-table sortable';
      const thead = document.createElement('thead'); const hrow = document.createElement('tr');
      keys.forEach(k => {
        const th = document.createElement('th');
        th.innerHTML = esc(k) + ' <span class="sort-icon">&#8645;</span>';
        th.onclick = () => sortTable(th);
        hrow.appendChild(th);
      });
      thead.appendChild(hrow); table.appendChild(thead);
      const tbody = document.createElement('tbody');
      val.forEach(item => {
        const tr = document.createElement('tr');
        keys.forEach(k => {
          const td = document.createElement('td');
          td.appendChild(renderValue(item ? item[k] : null, depth + 1));
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
      table.appendChild(tbody); outer.appendChild(table); wrap.appendChild(outer);
    } else {
      val.forEach(v => {
        const s = document.createElement('span'); s.className = 'badge';
        s.textContent = v === null ? 'null' : String(v);
        wrap.appendChild(s); wrap.appendChild(document.createTextNode(' '));
      });
    }
    return wrap;
  }

  // Boolean
  if (typeof val === 'boolean') {
    wrap.innerHTML = '<span class="bool bool-' + val + '">' + val + '</span>'; return wrap;
  }

  // Scalar
  wrap.textContent = String(val);
  return wrap;
}

// ── Collapse / expand ─────────────────────────────────────────────────────
function toggleBlock(btn){
  const body = btn.nextElementSibling, open = btn.getAttribute('aria-expanded') === 'true';
  btn.setAttribute('aria-expanded', String(!open));
  body.style.display = open ? 'none' : '';
}
function expandAll(){ document.querySelectorAll('.cmdlet-toggle').forEach(b => { b.setAttribute('aria-expanded','true'); b.nextElementSibling.style.display = ''; }); }
function collapseAll(){ document.querySelectorAll('.cmdlet-toggle').forEach(b => { b.setAttribute('aria-expanded','false'); b.nextElementSibling.style.display = 'none'; }); }

// ── Table sort ────────────────────────────────────────────────────────────
function sortTable(th){
  const table = th.closest('table'), tbody = table.querySelector('tbody');
  const idx = Array.from(th.parentNode.children).indexOf(th);
  const asc = th.dataset.sort !== 'asc';
  th.dataset.sort = asc ? 'asc' : 'desc';
  th.querySelector('.sort-icon').textContent = asc ? '\u25b2' : '\u25bc';
  Array.from(tbody.querySelectorAll('tr')).sort((a,b) => {
    const av = a.children[idx]?.textContent.trim() ?? '';
    const bv = b.children[idx]?.textContent.trim() ?? '';
    const n  = v => isNaN(v) ? v : Number(v);
    return asc ? (n(av) > n(bv) ? 1 : -1) : (n(av) < n(bv) ? 1 : -1);
  }).forEach(r => tbody.appendChild(r));
}

// ── Theme ─────────────────────────────────────────────────────────────────
function toggleTheme(){
  const dark = document.body.getAttribute('data-theme') === 'dark';
  document.body.setAttribute('data-theme', dark ? 'light' : 'dark');
  document.getElementById('themeBtn').innerHTML = dark ? '&#127769; Dark mode' : '&#9728; Light mode';
}

// ── Utilities ─────────────────────────────────────────────────────────────
function esc(s){ const d=document.createElement('div');d.textContent=String(s);return d.innerHTML; }
</script>
</body>
</html>
"@

    $DestinationFolder = Convert-Path -LiteralPath $DestinationFolder
    $base    = [System.IO.Path]::GetFileNameWithoutExtension($JsonFilePath)
    $outFile = Join-Path $DestinationFolder "$base.html"
    try {
        [System.IO.File]::WriteAllText($outFile, $html, [System.Text.Encoding]::UTF8)
        Write-LogSuccess "HTML report written to: $outFile"
        return $outFile
    }
    catch {
        Write-LogError "Failed to write HTML report: $($_.Exception.Message)"
        throw
    }
}
# ===========================================================================
#  REGION: Main
# ===========================================================================

function Main {
    Write-LogInfo "========================================================"
    Write-LogInfo " Exchange Configuration Dump v3.0 -- $(Get-Date -Format 'u')"
    Write-LogInfo "========================================================"

    if (-not (Test-Path $OutputPath)) {
        try   { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null; Write-LogInfo "Created output directory: $OutputPath" }
        catch { Write-LogError "Cannot create output directory '$OutputPath': $($_.Exception.Message)"; exit 1 }
    }

    $jsonFile = $null

    # ── ReportOnly: skip all Exchange work ────────────────────────────────
    if ($PSCmdlet.ParameterSetName -eq "ReportOnly") {
        # Convert-Path resolves relative paths against the PS working directory,
        # preventing .NET methods from anchoring to C:\Windows\system32
        $JsonInputFile = Convert-Path -LiteralPath $JsonInputFile
        Write-LogInfo "Mode: ReportOnly -- using JSON: $JsonInputFile"
        $jsonFile = $JsonInputFile
    }
    else {
        # ── Export (+ optional report) ────────────────────────────────────
        Write-LogInfo "Mode: $($PSCmdlet.ParameterSetName)"
        Write-LogInfo "Servers    : $($Servers -join ', ')"
        Write-LogInfo "Categories : $($Categories -join ', ')"

        $effectiveCategories = if ($Categories -contains "All") { $CategoryMap.Keys } else { $Categories }

        $sessionCreated = $false
        if (-not (Test-ExchangeCmdletsAvailable)) {
            try   { Connect-ExchangeRemote -ServerName $Servers[0]; $sessionCreated = $true }
            catch { Write-LogError "Aborting: could not connect to Exchange."; exit 1 }
        }
        else { Write-LogInfo "Exchange Management Shell cmdlets already loaded." }

        $output = [ordered]@{
            ExportMetadata = [ordered]@{
                ExportDate    = (Get-Date -Format "o")
                ExportHost    = $env:COMPUTERNAME
                ExportUser    = "$env:USERDOMAIN\$env:USERNAME"
                ScriptVersion = "3.0.0"
                TargetServers = $Servers
                Categories    = $effectiveCategories
            }
            Data = [ordered]@{}
        }

        $total  = @($effectiveCategories).Count
        $idx    = 0
        $errors = [System.Collections.Generic.List[string]]::new()

        foreach ($cat in $effectiveCategories) {
            $idx++
            Write-Progress -Activity "Exporting Exchange configuration" `
                           -Status   "Category $idx / $total : $cat" `
                           -PercentComplete ([int](($idx / $total) * 100))
            try {
                $output.Data[$cat] = Export-CategoryData -CategoryName $cat -TargetServers $Servers
                Write-LogSuccess "Category '$cat' completed."
            }
            catch {
                $msg = "Category '$cat' failed: $($_.Exception.Message)"
                Write-LogError $msg; $errors.Add($msg)
                $output.Data[$cat] = @{ "__Error__" = $_.Exception.Message }
            }
        }
        Write-Progress -Activity "Exporting Exchange configuration" -Completed

        $output.ExportMetadata["TotalCategoriesRequested"] = $total
        $output.ExportMetadata["CategoriesWithErrors"]     = $errors.Count
        $output.ExportMetadata["Errors"]                   = $errors.ToArray()

        if ($sessionCreated) { Disconnect-ExchangeRemote }

        $ts       = Get-Date -Format "yyyyMMdd_HHmmss"
        $jsonFile = Join-Path (Convert-Path -LiteralPath $OutputPath) ("ExchangeConfigDump_{0}.json" -f $ts)
        try {
            $json = $output | ConvertTo-Json -Depth 20 -WarningAction SilentlyContinue
            [System.IO.File]::WriteAllText($jsonFile, $json, [System.Text.Encoding]::UTF8)
            Write-LogSuccess "JSON written to: $jsonFile"
        }
        catch { Write-LogError "Failed to write JSON: $($_.Exception.Message)"; exit 1 }
    }

    # ── HTML report ───────────────────────────────────────────────────────
    $htmlFile = $null
    if ($GenerateHtmlReport) {
        try {
            $dest     = if ($PSCmdlet.ParameterSetName -eq "ReportOnly") { Split-Path $JsonInputFile -Parent } else { Convert-Path -LiteralPath $OutputPath }
            $htmlFile = New-HtmlReport -JsonFilePath $jsonFile -DestinationFolder $dest
        }
        catch { Write-LogError "HTML report generation failed: $($_.Exception.Message)" }
    }

    Write-LogInfo "========================================================"
    Write-LogInfo " Done."
    Write-LogInfo " JSON : $jsonFile"
    if ($htmlFile) { Write-LogInfo " HTML : $htmlFile" }
    Write-LogInfo " Log  : $script:LogFile"
    Write-LogInfo "========================================================"

    return [PSCustomObject]@{ JsonFile = $jsonFile; HtmlFile = $htmlFile }
}

# ===========================================================================
#  ENTRYPOINT
# ===========================================================================
try {
    $result = Main
    Write-Host "`nJSON  : $($result.JsonFile)" -ForegroundColor Cyan
    if ($result.HtmlFile) { Write-Host "HTML  : $($result.HtmlFile)" -ForegroundColor Green }
    Write-Host "Log   : $script:LogFile"       -ForegroundColor Cyan
}
catch {
    Write-LogError "Unhandled fatal error: $($_.Exception.Message)"
    Write-LogError $_.ScriptStackTrace
    exit 1
}
