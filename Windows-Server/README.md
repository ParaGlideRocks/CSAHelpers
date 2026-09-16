# Windows Server Scripts

This folder contains scripts focused on Windows Server diagnostics and event collection.

## Scripts

- [Collect-WindowsEvents](Collect-WindowsEvents/README.md): Collects configured events from remote Windows servers.
- [Detect-SecureBootCertUpdateStatus](Detect-SecureBootCertUpdateStatus/README.md): Reports UEFI CA 2023 and Secure Boot certificate status.
- [Find-OOBsGuid](Find-OOBsGuid/README.md): Finds Microsoft Update Catalog GUIDs for predefined out-of-band updates.

## Layout

Each script directory contains the PowerShell script and a README with requirements, parameters, examples, and output details. The `Collect-WindowsEvents` directory also contains its default JSON configuration.
