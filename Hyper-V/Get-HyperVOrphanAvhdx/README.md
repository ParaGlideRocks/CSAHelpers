# Get-HyperVOrphanAvhdx

Analyzes AVHD/AVHDX differencing disks on a Hyper-V host, reports the ones no virtual machine references any more, and flags broken disk chains. The script is read-only: it never merges, deletes, moves, or renames a disk, never removes a checkpoint, never stops a virtual machine, and never restarts a service.

## Why the scan is host-wide

Defining an orphan as "not in the chain of the selected VM" is unsafe, because virtual machines commonly share one `Virtual Hard Disks` folder: another VM's live disk would be reported as an orphan.

This script indexes every VM on the host first (attached disks plus every checkpoint disk, each walked up to its root parent), then treats a file as orphaned only when no VM references it, even when the report covers a single VM.

## Requirements

- Windows with the Hyper-V PowerShell module
- Windows PowerShell 5.1 or PowerShell 7 or later
- Elevation, or membership of the local `Hyper-V Administrators` group

## Parameters

- `VMName`: Virtual machines to report on. Wildcards are supported.
- `AllVM`: Report on every virtual machine without prompting.
- `AdditionalSearchPath`: Extra folders to scan recursively, such as an old backup location.
- `ExportFormat`: Any combination of `CSV`, `HTML`, and `JSON`.
- `ReportPath`: Destination folder for the report files. Defaults to the script directory.
- `SkipLockCheck`: Skips the per-file lock probe.
- `NoTranscript`: Does not write the session transcript.
- `Quiet`: Suppresses console output and produces only the export files. Implies non-interactive.

Running the script without parameters starts an interactive picker for the virtual machines and the export formats.

## Usage

```powershell
.\Get-HyperVOrphanAvhdx.ps1
.\Get-HyperVOrphanAvhdx.ps1 -AllVM -ExportFormat HTML,CSV,JSON
.\Get-HyperVOrphanAvhdx.ps1 -VMName 'SQL*','WEB-01' -AdditionalSearchPath 'E:\Backups\Old'
.\Get-HyperVOrphanAvhdx.ps1 -AllVM -Quiet -NoTranscript -ExportFormat JSON -ReportPath C:\Reports
```

## Output

The script prints a console report and writes `AvhdxAnalysis_<timestamp>.log` as a transcript unless `NoTranscript` is used. Requested exports are written as `AvhdxAnalysis_<host>_<timestamp>.html`, `.json`, and `_orphans.csv`, `_vmhealth.csv`, `_chains.csv`.

For every orphaned file the report includes size, last write time, declared parent, whether that parent exists and belongs to a live chain, parent/child identifier mismatch (the condition behind error `0xC03A000E`), current lock state, and a severity of `Info`, `Warning`, or `Error`.

For every virtual machine the report includes state, chain depth, checkpoint count, recovery checkpoints left behind by host-level backups with their age, ghost checkpoints whose disk files are missing, and broken chains.

The full result object is also emitted to the pipeline:

```powershell
$report = .\Get-HyperVOrphanAvhdx.ps1 -AllVM -Quiet
$report.Orphans | Where-Object Severity -eq 'Error'
```

Exit codes: `0` when nothing critical is found, `1` when there are critical orphans, broken chains, or ghost checkpoints, and `2` when the analysis cannot run.

## Notes

The script only reports. Before acting on a finding, confirm no backup job is running, power the virtual machine off, and copy the whole chain: `Merge-VHD` rewrites the parent and deletes the source, which is not reversible.

Files reported as `Error` should not be merged blindly. A missing parent, an identifier mismatch, or an unreadable header means the chain cannot be trusted, and `Set-VHD -IgnoreIdMismatch` produces a corrupt disk when the parent is not genuinely the correct one.

Lock detection opens each candidate file for reading with no sharing and closes it immediately; nothing is written. Use `SkipLockCheck` to skip that probe entirely.
