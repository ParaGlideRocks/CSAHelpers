# Find-ADConflicts

Finds mail and proxy address collisions between Active Directory users and contacts that can disrupt Microsoft Entra ID synchronization.

## Requirements

- Active Directory PowerShell module (RSAT)
- Domain connectivity and permission to read directory objects

## Parameters

- `OutputPath`: Destination CSV path. Defaults to `ConflictReport.csv` in the script directory.

## Usage

```powershell
.\Find-ADConflicts.ps1
.\Find-ADConflicts.ps1 -OutputPath C:\Reports\conflicts.csv
```

## Output

The CSV identifies each conflict, severity, matched value, contact, and user. The script only reads Active Directory and does not remediate conflicts.