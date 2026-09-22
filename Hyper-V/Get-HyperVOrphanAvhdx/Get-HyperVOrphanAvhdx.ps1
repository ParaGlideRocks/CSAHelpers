#Requires -Version 5.1

<#
.SYNOPSIS
    Read-only analysis of AVHD/AVHDX differencing disks on a Hyper-V host:
    rebuilds every disk chain, finds orphaned files and reports chain health.

.DESCRIPTION
    This tool NEVER modifies anything. It does not merge, delete, move, rename,
    stop a VM, remove a checkpoint or restart a service. Its only outputs are the
    console report, an optional transcript and the optional report files.

    Analysis performed:
      * Builds a HOST-WIDE index of every disk referenced by any VM on the host
        (attached disks + disks of every checkpoint), walking each ParentPath up
        to the root. A file is an orphan only when NO VM on the host references
        it - scanning a shared 'Virtual Hard Disks' folder therefore cannot
        mistake another VM's live disk for an orphan.
      * Scans the relevant folders for .avhd / .avhdx files and classifies each
        one as in-chain or orphaned.
      * Diagnoses each orphan: declared parent, parent present, parent already in
        an active chain, disk/parent identifier mismatch (0xC03A000E), file
        currently locked, readability.
      * Reports VM-level health: broken chains, checkpoints whose disk files are
        missing (ghost checkpoints), recovery checkpoints left behind by a
        host-level backup, and VMs currently backing up or merging.

    Run without parameters for an interactive menu; pass parameters for
    unattended use (scheduled task, monitoring, CI).

.PARAMETER VMName
    One or more VM names to analyse. Wildcards are supported. Omit together with
    -AllVM to get the interactive picker.

.PARAMETER AllVM
    Analyse every VM on the host without prompting.

.PARAMETER AdditionalSearchPath
    Extra folders to include in the AVHD/AVHDX scan (for example an old backup
    location). Folders are scanned recursively.

.PARAMETER ExportFormat
    One or more of CSV, HTML, JSON. Omit to produce no files.

.PARAMETER ReportPath
    Folder that receives the report files. Defaults to the script folder.

.PARAMETER SkipLockCheck
    Skip the per-file lock probe (a short read-only open attempt). Use it if the
    scan must not touch the files at all.

.PARAMETER NoTranscript
    Do not write the session transcript log.

.PARAMETER Quiet
    Suppress the console report; only produce the export files. Implies
    non-interactive.

.EXAMPLE
    .\Get-HyperVOrphanAvhdx.ps1
    Interactive: pick the VMs, pick the export formats.

.EXAMPLE
    .\Get-HyperVOrphanAvhdx.ps1 -AllVM -ExportFormat HTML,CSV,JSON
    Unattended full-host analysis with all three reports.

.EXAMPLE
    .\Get-HyperVOrphanAvhdx.ps1 -VMName 'SQL*' -AdditionalSearchPath 'D:\HyperV\Old'

.NOTES
    Read-only by design. Nothing in this script changes host or VM state.
    License: MIT. See the LICENSE file at the root of this repository.
#>

[CmdletBinding()]
param(
    [string[]]$VMName,

    [switch]$AllVM,

    [string[]]$AdditionalSearchPath,

    [ValidateSet('CSV', 'HTML', 'JSON')]
    [string[]]$ExportFormat,

    [string]$ReportPath,

    [switch]$SkipLockCheck,

    [switch]$NoTranscript,

    [switch]$Quiet
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ToolName    = 'Get-HyperVOrphanAvhdx'
$script:ToolVersion = '1.0.0'
$script:Quiet       = [bool]$Quiet

# =====================================================================
# Console helpers
# =====================================================================

function Write-Line {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '',
        Justification = 'Interactive console report; colour is part of the UX.')]
    param(
        [Parameter(ValueFromPipeline)][string]$Text = '',
        [System.ConsoleColor]$Color,
        [switch]$Force
    )
    process {
        if ($script:Quiet -and -not $Force) { return }
        if ($PSBoundParameters.ContainsKey('Color')) {
            Write-Host $Text -ForegroundColor $Color
        } else {
            Write-Host $Text
        }
    }
}

function Write-Section { param([string]$Text) Write-Line ''; Write-Line "==== $Text ====" -Color Cyan }
function Write-Ok      { param([string]$Text) Write-Line "[OK]    $Text" -Color Green }
function Write-Note    { param([string]$Text) Write-Line "[INFO]  $Text" -Color Gray }
function Write-Caution { param([string]$Text) Write-Line "[WARN]  $Text" -Color Yellow }
function Write-Problem { param([string]$Text) Write-Line "[ERROR] $Text" -Color Red }

function Read-Choice {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [Parameter(Mandatory)][string[]]$Valid
    )
    while ($true) {
        $answer = (Read-Host $Prompt).Trim().ToLowerInvariant()
        if ($Valid -contains $answer) { return $answer }
        Write-Caution "Invalid answer. Valid options: $($Valid -join ', ')"
    }
}

# =====================================================================
# Path helpers
#   Paths are compared through ConvertTo-ComparablePath and stored in
#   case-insensitive sets, because Windows paths are case-insensitive and
#   Hyper-V reports them with inconsistent casing.
# =====================================================================

function New-PathSet {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Pure in-memory factory, changes no system state.')]
    param([string[]]$Value)
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($v in $Value) { if ($v) { [void]$set.Add($v) } }
    return , $set
}

# An explicit OrdinalIgnoreCase comparer is used instead of the @{} literal:
# the default PowerShell hashtable comparer is culture sensitive, which would
# mis-handle paths under locales such as tr-TR (dotless-i problem).
function New-PathMap {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Pure in-memory factory, changes no system state.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseLiteralInitializerForHashtable', '',
        Justification = 'Needs an ordinal, culture-invariant comparer.')]
    param()
    return New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
}

# Canonical form used as dictionary/set key. Never throws.
function ConvertTo-ComparablePath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $value = $Path.Trim().Trim('"')
    try {
        $resolved = (Resolve-Path -LiteralPath $value -ErrorAction Stop).ProviderPath
        if ($resolved) { $value = $resolved }
    } catch {
        # Path cannot be resolved (missing file, unusual syntax): keep the raw
        # value so the comparison still works on the literal string.
        try { $value = [System.IO.Path]::GetFullPath($value) }
        catch { Write-Verbose "Cannot normalize '$value': $($_.Exception.Message)" }
    }
    return $value.TrimEnd('\')
}

function Test-AvhdPath {
    param([string]$Path)
    if (-not $Path) { return $false }
    # Exact extension test: '*.avhd' as a wildcard would also match '.avhdx'.
    $ext = [System.IO.Path]::GetExtension($Path)
    return ($ext -ieq '.avhd' -or $ext -ieq '.avhdx')
}

function Test-PathIsUnder {
    param([string]$Path, [string]$Parent)
    if (-not $Path -or -not $Parent) { return $false }
    $p = $Path.TrimEnd('\') + '\'
    $q = $Parent.TrimEnd('\') + '\'
    if ($p.Length -le $q.Length) { return $false }
    return $p.StartsWith($q, [System.StringComparison]::OrdinalIgnoreCase)
}

# Drop duplicates and folders already covered by another root, so a recursive
# scan never walks the same subtree twice.
function Get-DistinctSearchRoot {
    param([string[]]$Path)
    $normalized = [System.Collections.Generic.List[string]]::new()
    $seen = New-PathSet
    foreach ($p in $Path) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $n = ConvertTo-ComparablePath $p
        if ($n -and $seen.Add($n)) { $normalized.Add($n) }
    }
    $result = [System.Collections.Generic.List[string]]::new()
    foreach ($candidate in ($normalized | Sort-Object { $_.Length })) {
        $covered = $false
        foreach ($kept in $result) {
            if (Test-PathIsUnder -Path $candidate -Parent $kept) { $covered = $true; break }
        }
        if (-not $covered) { $result.Add($candidate) }
    }
    return , $result.ToArray()
}

function Get-FileSizeSafe {
    param([string]$Path)
    try { return [int64](Get-Item -LiteralPath $Path -Force -ErrorAction Stop).Length }
    catch { return $null }
}

function Format-Size {
    param([Nullable[int64]]$Bytes)
    if ($null -eq $Bytes) { return 'n/a' }
    $units = 'B', 'KB', 'MB', 'GB', 'TB', 'PB'
    $value = [double]$Bytes
    $i = 0
    while ($value -ge 1024 -and $i -lt ($units.Count - 1)) { $value /= 1024; $i++ }
    return ('{0:N1} {1}' -f $value, $units[$i])
}

# Read-only lock probe: opens the file for reading with no sharing. It never
# writes and never creates anything.
function Test-FileLocked {
    param([string]$Path)
    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
        return $false
    } catch [System.IO.IOException] {
        return $true
    } catch {
        return $null   # access denied or path problem: undetermined
    } finally {
        if ($stream) { $stream.Dispose() }
    }
}

# =====================================================================
# Chain walking
#   The VHD resolver is injected so the logic can be unit-tested without a
#   Hyper-V host.
# =====================================================================

function Get-VhdInfoSafe {
    param([string]$Path)
    return (Get-VHD -Path $Path -ErrorAction Stop)
}

<#
Walks Path -> parent -> ... -> root.
Returns @{ Status; Detail; Links[] } where Status is one of
Ok | MissingFile | Unreadable | Cycle | DepthExceeded.
#>
function Resolve-DiskChain {
    param(
        [Parameter(Mandatory)][string]$Path,
        [scriptblock]$Resolver = { param($p) Get-VhdInfoSafe -Path $p },
        [int]$MaxDepth = 256
    )

    $links   = [System.Collections.Generic.List[pscustomobject]]::new()
    $visited = New-PathSet
    $status  = 'Ok'
    $detail  = $null
    $current = $Path
    $level   = 0

    while (-not [string]::IsNullOrWhiteSpace($current)) {
        $key = ConvertTo-ComparablePath $current
        if (-not $visited.Add($key)) {
            $status = 'Cycle'
            $detail = "Parent chain loops back to '$current'."
            break
        }
        if ($level -ge $MaxDepth) {
            $status = 'DepthExceeded'
            $detail = "Chain deeper than $MaxDepth levels; stopped at '$current'."
            break
        }

        $exists = Test-Path -LiteralPath $current -PathType Leaf
        $info   = $null
        $err    = $null
        if ($exists) {
            try { $info = & $Resolver $current }
            catch { $err = $_.Exception.Message }
        }

        $parentPath = $null
        $diskId     = $null
        $parentId   = $null
        $vhdType    = $null
        if ($info) {
            if ($info.PSObject.Properties['ParentPath'])           { $parentPath = $info.ParentPath }
            if ($info.PSObject.Properties['DiskIdentifier'])       { $diskId     = $info.DiskIdentifier }
            if ($info.PSObject.Properties['ParentDiskIdentifier']) { $parentId   = $info.ParentDiskIdentifier }
            if ($info.PSObject.Properties['VhdType'])              { $vhdType    = [string]$info.VhdType }
        }

        $links.Add([pscustomobject]@{
            Level                = $level
            Path                 = $current
            Key                  = $key
            Exists               = $exists
            ParentPath           = $parentPath
            DiskIdentifier       = $diskId
            ParentDiskIdentifier = $parentId
            VhdType              = $vhdType
            Error                = $err
        })

        if (-not $exists) {
            $status = 'MissingFile'
            $detail = "File not found: '$current'."
            break
        }
        if ($err) {
            $status = 'Unreadable'
            $detail = "Cannot read '$current': $err"
            break
        }

        $current = $parentPath
        $level++
    }

    return [pscustomobject]@{
        Status = $status
        Detail = $detail
        Links  = $links.ToArray()
    }
}

# =====================================================================
# Host-wide disk reference index
#   Key   : comparable path
#   Value : list of @{ VMName; VMId; Context }
#   Every VM on the host is indexed, not only the selected ones.
# =====================================================================

function Add-DiskReference {
    param(
        [Parameter(Mandatory)][hashtable]$Index,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$VMName,
        [string]$VMId,
        [Parameter(Mandatory)][string]$Context
    )
    $key = ConvertTo-ComparablePath $Path
    if (-not $key) { return }
    if (-not $Index.ContainsKey($key)) {
        $Index[$key] = [System.Collections.Generic.List[pscustomobject]]::new()
    }
    $existing = $Index[$key] | Where-Object { $_.VMName -eq $VMName -and $_.Context -eq $Context }
    if (-not $existing) {
        $Index[$key].Add([pscustomobject]@{
            VMName  = $VMName
            VMId    = $VMId
            Context = $Context
            Path    = $Path
        })
    }
}

# Disks directly referenced by a VM: attached controllers + every checkpoint.
function Get-VmDiskReference {
    param([Parameter(Mandatory)]$Vm)

    $refs = [System.Collections.Generic.List[pscustomobject]]::new()

    try {
        foreach ($d in @(Get-VMHardDiskDrive -VM $Vm -ErrorAction Stop)) {
            if ($d.Path) {
                $refs.Add([pscustomobject]@{ Path = $d.Path; Context = 'Attached' })
            }
        }
    } catch {
        Write-Caution "Cannot read attached disks of '$($Vm.Name)': $($_.Exception.Message)"
    }

    try {
        foreach ($snap in @(Get-VMSnapshot -VM $Vm -ErrorAction Stop)) {
            foreach ($d in @(Get-VMHardDiskDrive -VMSnapshot $snap -ErrorAction Stop)) {
                if ($d.Path) {
                    $refs.Add([pscustomobject]@{
                        Path    = $d.Path
                        Context = "Checkpoint '$($snap.Name)'"
                    })
                }
            }
        }
    } catch {
        Write-Caution "Cannot read checkpoint disks of '$($Vm.Name)': $($_.Exception.Message)"
    }

    return , $refs.ToArray()
}

<#
Builds the host-wide index and, for each VM, the resolved chains.
Returns @{ Index; Chains } where Chains is a list of per-reference results.
#>
function New-DiskOwnerIndex {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Read-only: builds an in-memory index.')]
    param(
        [Parameter(Mandatory)][object[]]$Vm,
        [scriptblock]$Resolver = { param($p) Get-VhdInfoSafe -Path $p }
    )

    $index  = New-PathMap
    $chains = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($v in $Vm) {
        $vmId = if ($v.PSObject.Properties['Id']) { [string]$v.Id } else { '' }
        foreach ($ref in (Get-VmDiskReference -Vm $v)) {
            $chain = Resolve-DiskChain -Path $ref.Path -Resolver $Resolver
            foreach ($link in $chain.Links) {
                $ctx = if ($link.Level -eq 0) { $ref.Context } else { "$($ref.Context) -> parent L$($link.Level)" }
                Add-DiskReference -Index $index -Path $link.Path -VMName $v.Name -VMId $vmId -Context $ctx
            }
            $chains.Add([pscustomobject]@{
                VMName    = $v.Name
                VMId      = $vmId
                Context   = $ref.Context
                RootPath  = $ref.Path
                Status    = $chain.Status
                Detail    = $chain.Detail
                Depth     = $chain.Links.Count
                Links     = $chain.Links
            })
        }
    }

    return [pscustomobject]@{
        Index  = $index
        Chains = $chains.ToArray()
    }
}

# =====================================================================
# Folder discovery and scanning
# =====================================================================

function Get-SearchRoot {
    param(
        [Parameter(Mandatory)][object[]]$Vm,
        [Parameter(Mandatory)][hashtable]$Index,
        [string[]]$Extra
    )

    $candidates = [System.Collections.Generic.List[string]]::new()

    foreach ($key in $Index.Keys) {
        $dir = try { [System.IO.Path]::GetDirectoryName($key) } catch { $null }
        if ($dir) { $candidates.Add($dir) }
    }

    foreach ($v in $Vm) {
        foreach ($prop in 'Path', 'SnapshotFileLocation', 'ConfigurationLocation') {
            if ($v.PSObject.Properties[$prop] -and $v.$prop) {
                $candidates.Add([string]$v.$prop)
                $candidates.Add((Join-Path ([string]$v.$prop) 'Virtual Hard Disks'))
            }
        }
    }

    try {
        $vmHost = Get-VMHost -ErrorAction Stop
        if ($vmHost.VirtualHardDiskPath) { $candidates.Add([string]$vmHost.VirtualHardDiskPath) }
    } catch {
        Write-Verbose "Get-VMHost unavailable: $($_.Exception.Message)"
    }

    foreach ($e in $Extra) { if ($e) { $candidates.Add($e) } }

    $existing = foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c -PathType Container -ErrorAction SilentlyContinue) { $c }
    }
    return Get-DistinctSearchRoot -Path @($existing)
}

function Get-AvhdFile {
    param([Parameter(Mandatory)][string[]]$Root)

    $found = New-PathMap
    foreach ($r in $Root) {
        try {
            $items = Get-ChildItem -LiteralPath $r -Recurse -File -Force -ErrorAction SilentlyContinue
        } catch {
            Write-Caution "Cannot scan '$r': $($_.Exception.Message)"
            continue
        }
        foreach ($item in $items) {
            if (Test-AvhdPath $item.Name) {
                $key = ConvertTo-ComparablePath $item.FullName
                if ($key -and -not $found.ContainsKey($key)) { $found[$key] = $item }
            }
        }
    }
    return , @($found.Values | Sort-Object FullName)
}

# =====================================================================
# Orphan classification
# =====================================================================

function Get-OrphanDiagnostic {
    param(
        [Parameter(Mandatory)]$File,
        [Parameter(Mandatory)][hashtable]$Index,
        [scriptblock]$Resolver = { param($p) Get-VhdInfoSafe -Path $p },
        [switch]$SkipLock
    )

    $path       = $File.FullName
    $parentPath = $null
    $diskId     = $null
    $parentId   = $null
    $vhdType    = $null
    $readError  = $null

    try {
        $info = & $Resolver $path
        if ($info.PSObject.Properties['ParentPath'])           { $parentPath = $info.ParentPath }
        if ($info.PSObject.Properties['DiskIdentifier'])       { $diskId     = $info.DiskIdentifier }
        if ($info.PSObject.Properties['ParentDiskIdentifier']) { $parentId   = $info.ParentDiskIdentifier }
        if ($info.PSObject.Properties['VhdType'])              { $vhdType    = [string]$info.VhdType }
    } catch {
        $readError = $_.Exception.Message
    }

    $parentExists  = $false
    $parentInChain = $false
    $parentOwners  = @()
    $idMismatch    = $null

    if ($parentPath) {
        $parentExists = Test-Path -LiteralPath $parentPath -PathType Leaf
        $parentKey    = ConvertTo-ComparablePath $parentPath
        if ($parentKey -and $Index.ContainsKey($parentKey)) {
            $parentInChain = $true
            $parentOwners  = @($Index[$parentKey] | Select-Object -ExpandProperty VMName -Unique)
        }
        if ($parentExists -and $parentId) {
            try {
                $pInfo = & $Resolver $parentPath
                if ($pInfo.PSObject.Properties['DiskIdentifier'] -and $pInfo.DiskIdentifier) {
                    $idMismatch = ([string]$pInfo.DiskIdentifier -ne [string]$parentId)
                }
            } catch {
                Write-Verbose "Cannot read parent '$parentPath': $($_.Exception.Message)"
            }
        }
    }

    $locked = if ($SkipLock) { $null } else { Test-FileLocked -Path $path }

    $severity = 'Info'
    $reasons  = [System.Collections.Generic.List[string]]::new()

    if ($readError) {
        $severity = 'Error'
        $reasons.Add('File is unreadable as a VHD/VHDX (possibly corrupt or in use).')
    }
    if ($parentPath -and -not $parentExists) {
        $severity = 'Error'
        $reasons.Add('Declared parent file does not exist: the chain cannot be rebuilt.')
    }
    if ($idMismatch -eq $true) {
        $severity = 'Error'
        $reasons.Add('Parent/child identifier mismatch (merge would fail with 0xC03A000E).')
    }
    if (-not $parentPath -and -not $readError) {
        if ($severity -eq 'Info') { $severity = 'Warning' }
        $reasons.Add('No parent declared: the file is not a usable differencing disk.')
    }
    if ($parentExists -and -not $parentInChain -and $severity -eq 'Info') {
        $severity = 'Warning'
        $reasons.Add('Parent exists but is itself outside every active chain.')
    }
    if ($locked -eq $true) {
        $reasons.Add('File is currently locked by another process (backup job or VMMS).')
    }
    if ($reasons.Count -eq 0) {
        $reasons.Add('Not referenced by any VM; parent is a live disk. Typical leftover of an interrupted merge.')
    }

    return [pscustomobject]@{
        Path                 = $path
        FileName             = $File.Name
        Folder               = $File.DirectoryName
        SizeBytes            = [int64]$File.Length
        Size                 = Format-Size ([int64]$File.Length)
        LastWriteTime        = $File.LastWriteTime
        ParentPath           = $parentPath
        ParentExists         = $parentExists
        ParentInActiveChain  = $parentInChain
        ParentOwnerVMs       = ($parentOwners -join ', ')
        DiskIdentifier       = $diskId
        ParentDiskIdentifier = $parentId
        ParentIdMismatch     = $idMismatch
        VhdType              = $vhdType
        Locked               = $locked
        ReadError            = $readError
        Severity             = $severity
        Reason               = ($reasons -join ' ')
    }
}

# =====================================================================
# VM-level health
# =====================================================================

function Get-VmHealthReport {
    param(
        [Parameter(Mandatory)]$Vm,
        [Parameter(Mandatory)][object[]]$Chain
    )

    $vmChains = @($Chain | Where-Object { $_.VMName -eq $Vm.Name })

    $snapshots = @()
    try { $snapshots = @(Get-VMSnapshot -VM $Vm -ErrorAction Stop) }
    catch { Write-Verbose "Cannot list checkpoints of '$($Vm.Name)': $($_.Exception.Message)" }

    $recovery = @($snapshots | Where-Object {
            $_.PSObject.Properties['SnapshotType'] -and $_.SnapshotType -eq 'Recovery'
        })

    $ghost = [System.Collections.Generic.List[string]]::new()
    foreach ($snap in $snapshots) {
        try {
            foreach ($d in @(Get-VMHardDiskDrive -VMSnapshot $snap -ErrorAction Stop)) {
                if ($d.Path -and -not (Test-Path -LiteralPath $d.Path -PathType Leaf)) {
                    $ghost.Add("$($snap.Name) -> $($d.Path)")
                    break
                }
            }
        } catch {
            Write-Verbose "Cannot read disks of checkpoint '$($snap.Name)': $($_.Exception.Message)"
        }
    }

    $broken = @($vmChains | Where-Object { $_.Status -ne 'Ok' })

    $status = if ($Vm.PSObject.Properties['Status']) { [string]$Vm.Status } else { '' }
    $ops = ''
    if ($Vm.PSObject.Properties['Operations'] -and $Vm.Operations) {
        $ops = (@($Vm.Operations) | ForEach-Object { $_.ToString() }) -join ','
    }
    $busy = ($status -match 'Backing up|Merging') -or ($ops -match 'Backup|Merg|Export|Snapshot')

    $avhdInChain = @($vmChains | ForEach-Object { $_.Links } |
            Where-Object { Test-AvhdPath $_.Path } |
            Select-Object -ExpandProperty Path -Unique)

    return [pscustomobject]@{
        VMName                = $Vm.Name
        VMId                  = if ($Vm.PSObject.Properties['Id']) { [string]$Vm.Id } else { '' }
        State                 = if ($Vm.PSObject.Properties['State']) { [string]$Vm.State } else { '' }
        Status                = $status
        Operations            = $ops
        BusyWithBackupOrMerge = $busy
        DiskCount             = $vmChains.Count
        MaxChainDepth         = if ($vmChains.Count) { ($vmChains | Measure-Object -Property Depth -Maximum).Maximum } else { 0 }
        AvhdInChainCount      = $avhdInChain.Count
        CheckpointCount       = $snapshots.Count
        RecoveryCheckpoints   = ($recovery | ForEach-Object {
                $age = [math]::Round(((Get-Date) - $_.CreationTime).TotalHours, 1)
                "$($_.Name) (age ${age}h)"
            }) -join '; '
        GhostCheckpoints      = ($ghost -join '; ')
        BrokenChains          = ($broken | ForEach-Object { "$($_.RootPath) [$($_.Status)]" }) -join '; '
        Healthy               = ($broken.Count -eq 0 -and $ghost.Count -eq 0)
    }
}

# =====================================================================
# Reporting
# =====================================================================

function Write-ConsoleReport {
    param([Parameter(Mandatory)]$Result)

    Write-Section 'VM health'
    if ($Result.VMHealth.Count -eq 0) {
        Write-Caution 'No VM analysed.'
    } else {
        $Result.VMHealth |
            Select-Object VMName, State, Status, DiskCount, MaxChainDepth,
                          AvhdInChainCount, CheckpointCount, Healthy |
            Format-Table -AutoSize | Out-String -Width 4096 | Write-Line

        foreach ($h in $Result.VMHealth) {
            if ($h.BusyWithBackupOrMerge) {
                Write-Caution "'$($h.VMName)' is busy right now (Status '$($h.Status)', Ops '$($h.Operations)'). Chain data may be a moving target."
            }
            if ($h.RecoveryCheckpoints) {
                Write-Caution "'$($h.VMName)' has recovery checkpoint(s): $($h.RecoveryCheckpoints)"
            }
            if ($h.GhostCheckpoints) {
                Write-Problem "'$($h.VMName)' has checkpoint(s) whose disk files are missing: $($h.GhostCheckpoints)"
            }
            if ($h.BrokenChains) {
                Write-Problem "'$($h.VMName)' has broken chain(s): $($h.BrokenChains)"
            }
        }
    }

    Write-Section 'Scan scope'
    Write-Note "Folders scanned          : $($Result.SearchRoots.Count)"
    foreach ($r in $Result.SearchRoots) { Write-Line "   - $r" }
    Write-Note "AVHD/AVHDX found         : $($Result.Summary.TotalAvhdFiles)"
    Write-Note "Referenced by a VM       : $($Result.Summary.InChainCount)"
    Write-Note "Not referenced (orphans) : $($Result.Summary.OrphanCount)"

    Write-Section 'Orphan analysis'
    if ($Result.Orphans.Count -eq 0) {
        Write-Ok 'No orphaned AVHD/AVHDX found in the scanned folders.'
    } else {
        Write-Caution ("{0} orphaned file(s), {1} total." -f $Result.Orphans.Count, $Result.Summary.OrphanSize)
        $Result.Orphans |
            Select-Object @{ n = 'File'; e = { $_.FileName } },
                          @{ n = 'Size'; e = { $_.Size } },
                          @{ n = 'Modified'; e = { $_.LastWriteTime } },
                          Severity,
                          @{ n = 'ParentOK'; e = { $_.ParentExists } },
                          @{ n = 'ParentLive'; e = { $_.ParentInActiveChain } },
                          @{ n = 'Locked'; e = { $_.Locked } } |
            Format-Table -AutoSize | Out-String -Width 4096 | Write-Line

        foreach ($o in $Result.Orphans) {
            switch ($o.Severity) {
                'Error'   { Write-Problem "$($o.Path)"; Write-Line "          $($o.Reason)" -Color Red }
                'Warning' { Write-Caution "$($o.Path)"; Write-Line "          $($o.Reason)" -Color Yellow }
                default   { Write-Note    "$($o.Path)"; Write-Line "          $($o.Reason)" -Color Gray }
            }
            if ($o.ParentPath) { Write-Line "          parent: $($o.ParentPath)" -Color DarkGray }
        }

        Write-Line ''
        Write-Caution 'This tool is read-only: nothing above has been merged, moved or deleted.'
        Write-Note 'Review each file manually. Before acting: power the VM off, stop backup jobs, and take a copy of the whole chain.'
    }

    Write-Section 'Summary'
    Write-Line ("Analysed VMs        : {0}" -f $Result.Summary.VMCount)
    Write-Line ("Orphaned files      : {0} ({1})" -f $Result.Summary.OrphanCount, $Result.Summary.OrphanSize)
    Write-Line ("  of which critical : {0}" -f $Result.Summary.OrphanErrorCount)
    Write-Line ("Broken chains       : {0}" -f $Result.Summary.BrokenChainCount)
    Write-Line ("Ghost checkpoints   : {0}" -f $Result.Summary.GhostCheckpointCount)
    if ($Result.Summary.OrphanCount -eq 0 -and $Result.Summary.BrokenChainCount -eq 0 -and
        $Result.Summary.GhostCheckpointCount -eq 0) {
        Write-Ok 'Nothing to report: all chains are consistent.'
    }
}

function ConvertTo-HtmlReport {
    param([Parameter(Mandatory)]$Result)

    function Get-Encoded {
        param($Value)
        if ($null -eq $Value) { return '' }
        return [System.Net.WebUtility]::HtmlEncode([string]$Value)
    }

    function New-HtmlTable {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
            Justification = 'Builds an HTML string, changes no system state.')]
        param([object[]]$Row, [string[]]$Column, [string]$Empty = 'Nothing to show.')
        if (-not $Row -or $Row.Count -eq 0) { return "<p class='ok'>$(Get-Encoded $Empty)</p>" }
        $sb = [System.Text.StringBuilder]::new()
        [void]$sb.Append('<table><thead><tr>')
        foreach ($c in $Column) { [void]$sb.Append("<th>$(Get-Encoded $c)</th>") }
        [void]$sb.Append('</tr></thead><tbody>')
        foreach ($r in $Row) {
            $cls = ''
            if ($r.PSObject.Properties['Severity']) { $cls = " class='sev-$($r.Severity.ToLowerInvariant())'" }
            [void]$sb.Append("<tr$cls>")
            foreach ($c in $Column) {
                $val = if ($r.PSObject.Properties[$c]) { $r.$c } else { '' }
                [void]$sb.Append("<td>$(Get-Encoded $val)</td>")
            }
            [void]$sb.Append('</tr>')
        }
        [void]$sb.Append('</tbody></table>')
        return $sb.ToString()
    }

    $css = @'
body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#1b1b1b;background:#fafafa}
h1{font-size:22px;margin-bottom:4px}h2{font-size:17px;margin-top:28px;border-bottom:2px solid #ddd;padding-bottom:4px}
.meta{color:#666;font-size:12px;margin-bottom:18px}
table{border-collapse:collapse;width:100%;font-size:13px;background:#fff;box-shadow:0 1px 2px rgba(0,0,0,.08)}
th,td{border:1px solid #e0e0e0;padding:6px 9px;text-align:left;vertical-align:top;word-break:break-all}
th{background:#f0f0f0;font-weight:600}
tr:nth-child(even) td{background:#fbfbfb}
.sev-error td{background:#fdecea}.sev-warning td{background:#fff8e1}
.ok{color:#1b7f3b;font-weight:600}
.cards{display:flex;flex-wrap:wrap;gap:12px;margin:14px 0}
.card{background:#fff;border:1px solid #e0e0e0;border-radius:6px;padding:12px 18px;min-width:140px}
.card .n{font-size:24px;font-weight:700}.card .l{font-size:12px;color:#666}
.banner{background:#e8f0fe;border-left:4px solid #1a73e8;padding:10px 14px;margin:14px 0;font-size:13px}
ul{font-size:13px}
'@

    $s = $Result.Summary
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">')
    [void]$sb.AppendLine("<title>Hyper-V AVHDX analysis - $(Get-Encoded $Result.HostName)</title>")
    [void]$sb.AppendLine("<style>$css</style></head><body>")
    [void]$sb.AppendLine("<h1>Hyper-V orphaned AVHD/AVHDX analysis</h1>")
    [void]$sb.AppendLine("<div class='meta'>Host <b>$(Get-Encoded $Result.HostName)</b> &middot; generated $(Get-Encoded $Result.GeneratedAt) &middot; $(Get-Encoded $Result.Tool) v$(Get-Encoded $Result.Version)</div>")
    [void]$sb.AppendLine("<div class='banner'>This report is <b>read-only</b>. No disk was merged, moved or deleted. Verify every finding before acting.</div>")

    [void]$sb.AppendLine("<div class='cards'>")
    foreach ($card in @(
            @{ n = $s.VMCount; l = 'VMs analysed' },
            @{ n = $s.TotalAvhdFiles; l = 'AVHD/AVHDX found' },
            @{ n = $s.OrphanCount; l = 'Orphans' },
            @{ n = $s.OrphanSize; l = 'Orphan size' },
            @{ n = $s.OrphanErrorCount; l = 'Critical orphans' },
            @{ n = $s.BrokenChainCount; l = 'Broken chains' },
            @{ n = $s.GhostCheckpointCount; l = 'Ghost checkpoints' })) {
        [void]$sb.AppendLine("<div class='card'><div class='n'>$(Get-Encoded $card.n)</div><div class='l'>$(Get-Encoded $card.l)</div></div>")
    }
    [void]$sb.AppendLine('</div>')

    [void]$sb.AppendLine('<h2>VM health</h2>')
    [void]$sb.AppendLine((New-HtmlTable -Row $Result.VMHealth -Empty 'No VM analysed.' -Column @(
                'VMName', 'State', 'Status', 'DiskCount', 'MaxChainDepth', 'AvhdInChainCount',
                'CheckpointCount', 'RecoveryCheckpoints', 'GhostCheckpoints', 'BrokenChains', 'Healthy')))

    [void]$sb.AppendLine('<h2>Orphaned AVHD/AVHDX</h2>')
    [void]$sb.AppendLine((New-HtmlTable -Row $Result.Orphans -Empty 'No orphaned AVHD/AVHDX found.' -Column @(
                'Path', 'Size', 'LastWriteTime', 'Severity', 'ParentPath', 'ParentExists',
                'ParentInActiveChain', 'ParentOwnerVMs', 'ParentIdMismatch', 'Locked', 'Reason')))

    [void]$sb.AppendLine('<h2>Active disk chains</h2>')
    [void]$sb.AppendLine((New-HtmlTable -Row $Result.Chains -Empty 'No chain resolved.' -Column @(
                'VMName', 'Context', 'RootPath', 'Depth', 'Status', 'Detail')))

    [void]$sb.AppendLine('<h2>Folders scanned</h2><ul>')
    foreach ($r in $Result.SearchRoots) { [void]$sb.AppendLine("<li>$(Get-Encoded $r)</li>") }
    [void]$sb.AppendLine('</ul>')

    [void]$sb.AppendLine('</body></html>')
    return $sb.ToString()
}

function Export-AnalysisReport {
    param(
        [Parameter(Mandatory)]$Result,
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string[]]$Format
    )

    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
        try { New-Item -ItemType Directory -Path $Folder -Force -ErrorAction Stop | Out-Null }
        catch { Write-Problem "Cannot create report folder '$Folder': $($_.Exception.Message)"; return @() }
    }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $base  = Join-Path $Folder ("AvhdxAnalysis_{0}_{1}" -f $Result.HostName, $stamp)
    $written = [System.Collections.Generic.List[string]]::new()

    foreach ($f in ($Format | Sort-Object -Unique)) {
        try {
            switch ($f) {
                'CSV' {
                    $map = @{
                        'orphans.csv'  = $Result.Orphans
                        'vmhealth.csv' = $Result.VMHealth
                        'chains.csv'   = ($Result.Chains | Select-Object VMName, Context, RootPath, Depth, Status, Detail)
                    }
                    foreach ($name in $map.Keys) {
                        $target = "{0}_{1}" -f $base, $name
                        $rows = @($map[$name])
                        if ($rows.Count -eq 0) {
                            Set-Content -LiteralPath $target -Value '' -Encoding UTF8
                        } else {
                            $rows | Export-Csv -LiteralPath $target -NoTypeInformation -Encoding UTF8
                        }
                        $written.Add($target)
                    }
                }
                'JSON' {
                    $target = "$base.json"
                    $Result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $target -Encoding UTF8
                    $written.Add($target)
                }
                'HTML' {
                    $target = "$base.html"
                    ConvertTo-HtmlReport -Result $Result | Set-Content -LiteralPath $target -Encoding UTF8
                    $written.Add($target)
                }
            }
        } catch {
            Write-Problem "Cannot write the $f report: $($_.Exception.Message)"
        }
    }

    return , $written.ToArray()
}

# =====================================================================
# Analysis driver
# =====================================================================

function Invoke-AvhdAnalysis {
    param(
        [Parameter(Mandatory)][object[]]$TargetVm,
        [Parameter(Mandatory)][object[]]$AllHostVm,
        [string[]]$ExtraPath,
        [switch]$SkipLock
    )

    Write-Section 'Building host-wide disk index'
    Write-Note "VMs on this host: $($AllHostVm.Count) (all indexed, to protect disks owned by other VMs)"
    Write-Note "VMs selected for the report: $($TargetVm.Count)"

    $owner = New-DiskOwnerIndex -Vm $AllHostVm
    Write-Note "Distinct disk files referenced by a VM: $($owner.Index.Count)"

    Write-Section 'Scanning folders'
    $roots = Get-SearchRoot -Vm $AllHostVm -Index $owner.Index -Extra $ExtraPath
    foreach ($r in $roots) { Write-Note "scan: $r" }

    $files = Get-AvhdFile -Root $roots
    Write-Note "AVHD/AVHDX files found: $($files.Count)"

    $orphanList = [System.Collections.Generic.List[pscustomobject]]::new()
    $inChain = 0
    foreach ($f in $files) {
        $key = ConvertTo-ComparablePath $f.FullName
        if ($owner.Index.ContainsKey($key)) { $inChain++; continue }
        $orphanList.Add((Get-OrphanDiagnostic -File $f -Index $owner.Index -SkipLock:$SkipLock))
    }
    $orphans = @($orphanList | Sort-Object @{ e = {
                switch ($_.Severity) { 'Error' { 0 } 'Warning' { 1 } default { 2 } } } }, Path)

    $health = foreach ($v in $TargetVm) { Get-VmHealthReport -Vm $v -Chain $owner.Chains }
    $health = @($health)

    $orphanBytes = ($orphans | Measure-Object -Property SizeBytes -Sum).Sum
    if (-not $orphanBytes) { $orphanBytes = 0 }

    $ghostCount = 0
    foreach ($h in $health) { if ($h.GhostCheckpoints) { $ghostCount += @($h.GhostCheckpoints -split ';').Count } }

    return [pscustomobject]@{
        Tool        = $script:ToolName
        Version     = $script:ToolVersion
        HostName    = $env:COMPUTERNAME
        GeneratedAt = (Get-Date).ToString('s')
        SearchRoots = $roots
        VMHealth    = $health
        Orphans     = $orphans
        Chains      = @($owner.Chains | Select-Object VMName, Context, RootPath, Depth, Status, Detail)
        Summary     = [pscustomobject]@{
            VMCount              = $TargetVm.Count
            HostVMCount          = $AllHostVm.Count
            TotalAvhdFiles       = $files.Count
            InChainCount         = $inChain
            OrphanCount          = $orphans.Count
            OrphanErrorCount     = @($orphans | Where-Object { $_.Severity -eq 'Error' }).Count
            OrphanBytes          = [int64]$orphanBytes
            OrphanSize           = Format-Size ([int64]$orphanBytes)
            BrokenChainCount     = @($owner.Chains | Where-Object { $_.Status -ne 'Ok' }).Count
            GhostCheckpointCount = $ghostCount
        }
    }
}

# =====================================================================
# Preflight and VM selection
# =====================================================================

function Test-Prerequisite {
    if (-not (Get-Module -ListAvailable -Name Hyper-V)) {
        Write-Problem 'The Hyper-V PowerShell module is not installed on this machine.'
        Write-Note    'Install it with: Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-Management-PowerShell'
        return $false
    }
    try { Import-Module Hyper-V -ErrorAction Stop }
    catch { Write-Problem "Cannot load the Hyper-V module: $($_.Exception.Message)"; return $false }

    try {
        $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-Caution 'Not running elevated. Hyper-V cmdlets may fail unless you are a member of "Hyper-V Administrators".'
        }
    } catch {
        Write-Verbose "Elevation check failed: $($_.Exception.Message)"
    }
    return $true
}

function Select-TargetVm {
    param(
        [Parameter(Mandatory)][object[]]$AllVm,
        [string[]]$Name,
        [switch]$All,
        [switch]$Interactive
    )

    if ($All) { return , $AllVm }

    if ($Name) {
        $selected = [System.Collections.Generic.List[object]]::new()
        foreach ($pattern in $Name) {
            $matched = @($AllVm | Where-Object { $_.Name -like $pattern })
            if ($matched.Count -eq 0) { Write-Caution "No VM matches '$pattern'." }
            foreach ($m in $matched) {
                if (-not ($selected | Where-Object { $_.Id -eq $m.Id })) { $selected.Add($m) }
            }
        }
        return , $selected.ToArray()
    }

    if (-not $Interactive) { return , $AllVm }

    Write-Section 'VM selection'
    for ($i = 0; $i -lt $AllVm.Count; $i++) {
        Write-Line ('{0,3}) {1}  [{2}]' -f ($i + 1), $AllVm[$i].Name, $AllVm[$i].State)
    }
    Write-Line '  a) all VMs on this host'

    while ($true) {
        $raw = (Read-Host "`nSelect VM (number, comma separated list, or 'a')").Trim()
        if ($raw -match '^(a|all)$') { return , $AllVm }

        $picked = [System.Collections.Generic.List[object]]::new()
        $bad = $false
        foreach ($token in ($raw -split '[,\s]+' | Where-Object { $_ })) {
            $n = 0
            if (-not [int]::TryParse($token, [ref]$n) -or $n -lt 1 -or $n -gt $AllVm.Count) { $bad = $true; break }
            $vm = $AllVm[$n - 1]
            if (-not ($picked | Where-Object { $_.Id -eq $vm.Id })) { $picked.Add($vm) }
        }
        if (-not $bad -and $picked.Count -gt 0) { return , $picked.ToArray() }
        Write-Caution "Invalid selection. Enter numbers between 1 and $($AllVm.Count), or 'a'."
    }
}

function Select-ExportFormat {
    Write-Section 'Report export'
    Write-Line '  1) none (console only)'
    Write-Line '  2) HTML'
    Write-Line '  3) CSV'
    Write-Line '  4) JSON'
    Write-Line '  5) all three'
    switch (Read-Choice 'Choice (1-5)' @('1', '2', '3', '4', '5')) {
        '2' { return @('HTML') }
        '3' { return @('CSV') }
        '4' { return @('JSON') }
        '5' { return @('CSV', 'HTML', 'JSON') }
        default { return @() }
    }
}

# =====================================================================
# Entry point
# =====================================================================

function Invoke-Main {
    param(
        [string[]]$VMName,
        [switch]$AllVM,
        [string[]]$AdditionalSearchPath,
        [string[]]$ExportFormat,
        [string]$ReportPath,
        [switch]$SkipLockCheck,
        [switch]$NoTranscript,
        [bool]$Interactive
    )

    $script:ExitCode = 0
    $interactive = $Interactive

    $workDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).ProviderPath }
    $reportDir = if ($ReportPath) { $ReportPath } else { $workDir }

    $transcriptStarted = $false
    if (-not $NoTranscript) {
        $logPath = Join-Path $workDir ("AvhdxAnalysis_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        try {
            Start-Transcript -LiteralPath $logPath -ErrorAction Stop | Out-Null
            $transcriptStarted = $true
        } catch {
            Write-Caution "Cannot start the transcript log: $($_.Exception.Message)"
        }
    }

    try {
        Write-Line ''
        Write-Line "$script:ToolName v$script:ToolVersion - read-only Hyper-V AVHD/AVHDX analysis" -Color Cyan
        Write-Line 'This tool never merges, moves, renames or deletes anything.' -Color DarkGray

        if (-not (Test-Prerequisite)) { $script:ExitCode = 2; return }

        $allVm = @()
        try { $allVm = @(Get-VM -ErrorAction Stop | Sort-Object Name) }
        catch { Write-Problem "Cannot enumerate VMs: $($_.Exception.Message)"; $script:ExitCode = 2; return }

        if ($allVm.Count -eq 0) { Write-Caution 'No VM found on this host.'; return }

        $targets = Select-TargetVm -AllVm $allVm -Name $VMName -All:$AllVM -Interactive:$interactive
        if ($targets.Count -eq 0) { Write-Problem 'No VM selected.'; $script:ExitCode = 2; return }
        Write-Ok ("Selected: {0}" -f (($targets | ForEach-Object { $_.Name }) -join ', '))

        $formats = @($ExportFormat)
        if ($interactive) { $formats = Select-ExportFormat }

        $result = Invoke-AvhdAnalysis -TargetVm $targets -AllHostVm $allVm `
            -ExtraPath $AdditionalSearchPath -SkipLock:$SkipLockCheck

        Write-ConsoleReport -Result $result

        if ($formats.Count -gt 0) {
            Write-Section 'Reports'
            foreach ($file in (Export-AnalysisReport -Result $result -Folder $reportDir -Format $formats)) {
                Write-Ok "Written: $file"
            }
        }

        if ($result.Summary.OrphanErrorCount -gt 0 -or $result.Summary.BrokenChainCount -gt 0 -or
            $result.Summary.GhostCheckpointCount -gt 0) {
            $script:ExitCode = 1
        }

        # Emit the object so the script can be consumed from a pipeline.
        Write-Output $result
    } finally {
        if ($transcriptStarted) { try { Stop-Transcript | Out-Null } catch { Write-Verbose 'Transcript already stopped.' } }
    }
}

# Guard so the file can be dot-sourced by tests without running the analysis.
if ($MyInvocation.InvocationName -ne '.') {
    $script:ExitCode = 0
    Invoke-Main -VMName $VMName -AllVM:$AllVM -AdditionalSearchPath $AdditionalSearchPath `
        -ExportFormat $ExportFormat -ReportPath $ReportPath -SkipLockCheck:$SkipLockCheck `
        -NoTranscript:$NoTranscript -Interactive ($PSBoundParameters.Count -eq 0 -and -not $Quiet)
    exit $script:ExitCode
}
