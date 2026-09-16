# Find-OOBsGuid

Searches the Microsoft Update Catalog for a predefined list of out-of-band KB updates and exports each matching update GUID and product scope.

## Requirements

- PowerShell 5.0 or later
- Internet access to `catalog.update.microsoft.com`

## Parameters

- `OutputPath`: Destination CSV path. Defaults to `OOB-GUIDs.csv` in the script directory.

## Usage

```powershell
.\Find-OOBsGuid.ps1
.\Find-OOBsGuid.ps1 -OutputPath C:\Temp\OOB-GUIDs.csv
```

## Output

The script displays results in the console and writes a CSV containing `KB`, `UpdateID`, and `Products`.

## Notes

The KB list is defined in the script. Catalog page parsing depends on the current Microsoft Update Catalog HTML structure.