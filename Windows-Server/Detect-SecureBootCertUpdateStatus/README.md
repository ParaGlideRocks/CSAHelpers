# Detect-SecureBootCertUpdateStatus

Reports Secure Boot certificate update readiness and UEFI CA 2023 rollout status for a Windows device.

## Requirements

- Windows 10 or Windows 11 on a UEFI-capable device
- Permission to read the relevant registry and event log data

## Parameters

- `OutputPath`: Optional directory or network share for the JSON result. When omitted, JSON is written to standard output.

## Usage

```powershell
.\Detect-SecureBootCertUpdateStatus.ps1
.\Detect-SecureBootCertUpdateStatus.ps1 -OutputPath \\server\SecureBootLogs$
```

## Output

The result includes device, firmware, Secure Boot, certificate, registry, and related event data. Exit code `0` indicates no detected issue; exit code `1` indicates that attention may be required.

## Notes

The script is read-only and does not apply Secure Boot remediation.