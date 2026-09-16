# Collect-WindowsEvents

Collects Windows events from remote servers according to a JSON configuration and exports normalized results to JSON.

## Requirements

- PowerShell 5.0 or later
- Network access and permission to read the configured remote event logs

## Parameters

- `ConfigPath`: Configuration file path. Defaults to `EventCollection-Config.json` beside the script.
- `OutputPath`: Destination JSON path. Defaults to a timestamped file.
- `Credential`: Optional credential for remote event log access.
- `MaxEventsPerQuery`: Maximum events returned by each query. Defaults to `1000`.

## Configuration

Edit `EventCollection-Config.json` to define servers, event log names, event IDs, and either explicit start/end times or a relative `lastHours` window.

## Usage

```powershell
.\Collect-WindowsEvents.ps1
.\Collect-WindowsEvents.ps1 -Credential (Get-Credential) -MaxEventsPerQuery 500
```

## Output

The JSON contains collection metadata, a per-server summary, and normalized event records.