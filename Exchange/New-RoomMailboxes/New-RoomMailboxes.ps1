<#
.SYNOPSIS
    Provisioning massivo di room mailbox per Room Finder / Microsoft Places.

.DESCRIPTION
    Lo script legge l'anagrafica delle sale da un file CSV e crea/aggiorna le room
    mailbox su Exchange Online in tre fasi distinte:

        Fase 1 - creazione delle room mailbox mancanti (New-Mailbox)
        Fase 2 - attesa della propagazione e configurazione dei Place e del
                 calendar processing (Set-Mailbox, Set-Place, Set-CalendarProcessing)
        Fase 3 - creazione delle room list e inserimento delle sale

    L'esecuzione e' idempotente: una sala gia' presente viene aggiornata e non
    ricreata, quindi lo script puo' essere rilanciato sulle sale rimaste indietro
    senza duplicare quelle gia' create.

.PARAMETER CsvPath
    Percorso del file CSV di anagrafica (delimitatore ';').

.PARAMETER Domain
    Dominio SMTP da usare per gli indirizzi delle room mailbox.

.PARAMETER Prefix
    Prefisso opzionale per alias e nomi delle room list.

.PARAMETER HeaderRow
    Riga del CSV che contiene le intestazioni (default 1).

.PARAMETER MaxRoomsPerList
    Numero massimo di sale per room list (default 50, come raccomandato da Microsoft).

.PARAMETER SkipRoomLists
    Salta la fase 3 (creazione e popolamento delle room list).

.PARAMETER WhatIfProvisioning
    Esegue l'intero flusso in sola simulazione: nessuna modifica sul tenant.

.EXAMPLE
    .\New-RoomMailboxes.ps1 -CsvPath '.\sample-rooms.csv' -Domain 'contoso.com' -WhatIfProvisioning

.NOTES
    Version 1.0.0
    Richiede il modulo ExchangeOnlineManagement V3 o superiore: versioni
    precedenti possono fallire quando Set-Place imposta piu' proprieta' insieme.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath = ".\sample-rooms.csv",

    [Parameter(Mandatory = $false)]
    [string]$Domain,

    # Prefisso applicato ad alias e nomi delle room list, utile quando si genera
    # un nuovo insieme di sale che deve restare distinguibile da quelle esistenti.
    # Se l'anagrafica espone una colonna Prefix, quel valore ha la precedenza.
    [Parameter(Mandatory = $false)]
    [string]$Prefix = '',

    # Paese usato quando l'anagrafica non espone una colonna Country.
    [Parameter(Mandatory = $false)]
    [string]$DefaultCountry = 'IT',

    # Riga del file che contiene le intestazioni. Alcuni modelli Excel esportati in
    # CSV hanno titolo e righe vuote in testa: in quel caso indicare la riga giusta.
    [Parameter(Mandatory = $false)]
    [int]$HeaderRow = 1,

    [Parameter(Mandatory = $false)]
    [int]$PropagationWaitSeconds = 120,

    [Parameter(Mandatory = $false)]
    [int]$MaxAttempts = 6,

    [Parameter(Mandatory = $false)]
    [int]$RetryDelaySeconds = 30,

    [Parameter(Mandatory = $false)]
    [int]$MaxRoomsPerList = 50,

    [Parameter(Mandatory = $false)]
    [switch]$SkipRoomLists,

    [Parameter(Mandatory = $false)]
    [switch]$WhatIfProvisioning
)

$ErrorActionPreference = 'Stop'

Import-Module ExchangeOnlineManagement

$LogPath = Join-Path $PSScriptRoot "New-RoomMailboxes_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

#region Logging -----------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timestamp] [$Level] $Message"

    $color = switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARNING' { 'Yellow' }
        'ERROR'   { 'Red' }
        default   { 'Cyan' }
    }

    Write-Host $entry -ForegroundColor $color
    Add-Content -Path $LogPath -Value $entry
}

function Write-LogInfo    { param([string]$Message) Write-Log -Message $Message -Level 'INFO' }
function Write-LogSuccess { param([string]$Message) Write-Log -Message $Message -Level 'SUCCESS' }
function Write-LogWarning { param([string]$Message) Write-Log -Message $Message -Level 'WARNING' }
function Write-LogError   { param([string]$Message) Write-Log -Message $Message -Level 'ERROR' }

function Write-CommandWarnings {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Command,

        [AllowEmptyCollection()]
        [object[]]$Warnings
    )

    foreach ($warning in $Warnings) {
        Write-LogWarning "[$Command] $warning"
    }
}

#endregion

#region Helper ------------------------------------------------------------------

function ConvertTo-AliasToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $normalized = $Value.Normalize([Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder

    foreach ($char in $normalized.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($char) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($char)
        }
    }

    $token = $builder.ToString().Normalize([Text.NormalizationForm]::FormC).ToLowerInvariant()
    $token = $token -replace '[^a-z0-9]+', '-'

    return $token.Trim('-')
}

function Get-RowValue {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Row,

        [Parameter(Mandatory = $true)]
        [string[]]$FieldNames
    )

    foreach ($fieldName in $FieldNames) {
        $value = $Row.PSObject.Properties |
            Where-Object { $_.Name.Trim() -eq $fieldName } |
            Select-Object -First 1 -ExpandProperty Value

        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        return ([string]$value).Trim()
    }

    return $null
}

# Accetta sia i valori inglesi (yes/y/true/1) sia quelli italiani (si/si'/s)
# usati nelle anagrafiche compilate in italiano, dove lo stesso file puo' mescolare
# "si"/"no" e "Yes"/"No" a seconda della colonna.
function Test-Yes {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    return $Value.Trim() -match '^(yes|y|si|s|true|vero|1)$'
}

function ConvertTo-FloorNumber {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    switch -Regex ($Value.Trim()) {
        '^pt$'         { return 0 }
        '^p\s*(?<N>-?\d+)$' { return [int]$Matches.N }
        '^-?\d+$'      { return [int]$Value }
        default        { throw "Valore di piano non valido '$Value'. Attesi pt, pN oppure un intero." }
    }
}

# Alias deterministico: costruito dai campi strutturati dell'anagrafica
# (Prefix, City, Building, Floor, Name) e non dalla stringa DisplayName completa.
# Lo stesso record produce sempre lo stesso alias tra un'esecuzione e l'altra.
function New-RoomAlias {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [string[]]$Parts,

        [int]$MaxLength = 64
    )

    $tokens = foreach ($part in $Parts) {
        if (-not [string]::IsNullOrWhiteSpace($part)) {
            $token = ConvertTo-AliasToken -Value $part

            if (-not [string]::IsNullOrWhiteSpace($token)) {
                $token
            }
        }
    }

    if (-not $tokens -or $tokens.Count -eq 0) {
        throw "Impossibile generare un alias valido dai campi forniti."
    }

    $alias = ($tokens -join '-')

    if ($alias.Length -gt $MaxLength) {
        $alias = $alias.Substring(0, $MaxLength).Trim('-')
    }

    return $alias
}

function Resolve-UniqueValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [int]$MaxLength,

        [Parameter(Mandatory = $true)]
        [hashtable]$UsedValues
    )

    if ($Value.Length -gt $MaxLength) {
        $Value = $Value.Substring(0, $MaxLength).Trim('-').TrimEnd()
    }

    $base = $Value
    $sequence = 2

    while ($UsedValues.ContainsKey($Value.ToLowerInvariant())) {
        $suffix = "-$sequence"
        $available = $MaxLength - $suffix.Length

        if ($base.Length -gt $available) {
            $Value = $base.Substring(0, $available).Trim('-').TrimEnd() + $suffix
        }
        else {
            $Value = $base + $suffix
        }

        $sequence++
    }

    $UsedValues[$Value.ToLowerInvariant()] = $true

    return $Value
}

# Attende che l'oggetto sia effettivamente risolvibile su Exchange Online prima
# di configurarlo: e' la contromisura alla latenza di replica che nell'esecuzione
# del 27 luglio 2026 ha prodotto 26 errori PlaceNotFoundInDirectory.
function Wait-MailboxAvailable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,

        [int]$Attempts = 6,

        [int]$DelaySeconds = 30
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            $mailbox = Get-Mailbox -Identity $Identity -ErrorAction Stop

            if ($mailbox) {
                # Set-Place legge la directory dei Place, che si popola dopo la
                # mailbox: si verifica anche quella prima di dichiarare pronto.
                try {
                    $null = Get-Place -Identity $Identity -ErrorAction Stop
                    return $true
                }
                catch {
                    Write-LogInfo "Mailbox '$Identity' presente, directory Place non ancora pronta (tentativo $attempt di $Attempts)."
                }
            }
        }
        catch {
            Write-LogInfo "Mailbox '$Identity' non ancora risolvibile (tentativo $attempt di $Attempts)."
        }

        if ($attempt -lt $Attempts) {
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    return $false
}

#endregion

#region Mappa colonne -----------------------------------------------------------

#
# L'anagrafica di origine cambia spesso struttura tra una versione e l'altra. La mappa elenca,
# per ogni informazione, tutti i nomi di colonna incontrati finora: il primo che
# esiste e valorizzato vince. Cosi' lo stesso script legge sia il vecchio CSV
# (Country;City;Building;Floor;Prefix;Name;DisplayName;...) sia il modello
# aggiornato in italiano (Sede;Edificio;Piano;Nome Sala Riunioni;...), che non ha
# piu' le colonne Prefix, Country, Name e DisplayName.
#
$Field = @{
    Name            = @('Nome Sala Riunioni', 'DisplayName', 'Name')
    City            = @('Sede', 'City')
    Building        = @('Edificio', 'Building')
    Floor           = @('Piano', 'Floor')
    Capacity        = @('Capienza (numero posti a sedere)', 'Capienza', 'ResourceCapacity', 'Capacity')
    Audio           = @('Call conference', 'AudioDeviceName')
    Video           = @('Video conference', 'VideoDeviceName')
    Display         = @('Proiettore', 'DisplayDeviceName')
    Description     = @('Descrizione della Sala *(facoltativo)', 'Descrizione della Sala', 'DeviceDescription')
    RequiresApproval = @('Workflow approvativo', 'RequiresApproval')
    Approvers       = @('Approvatore (email)', 'Approvatore', 'Approvers', 'ApproverEmail')
    Country         = @('CountryOrRegion', 'Country')
    Prefix          = @('Prefix')
}

#endregion

#region Connessione -------------------------------------------------------------

Write-LogInfo "File di log: $LogPath"

if ($WhatIfProvisioning) {
    Write-LogWarning "Modalita' simulazione attiva: nessuna modifica verra' applicata al tenant."
}

$module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
    Sort-Object Version -Descending |
    Select-Object -First 1

if ($module -and $module.Version.Major -lt 3) {
    Write-LogWarning "Modulo ExchangeOnlineManagement $($module.Version) rilevato. Set-Place puo' fallire con piu' proprieta' insieme: si raccomanda la V3."
}

$existingSession = Get-ConnectionInformation -ErrorAction SilentlyContinue |
    Where-Object { $_.State -eq 'Connected' -and $_.Name -like 'ExchangeOnline*' }

if ($existingSession) {
    Write-LogInfo "Riutilizzo della sessione Exchange Online esistente ($($existingSession.UserPrincipalName))."
}
else {
    Write-LogInfo "Connessione a Exchange Online..."
    Connect-ExchangeOnline
    Write-LogSuccess "Connesso a Exchange Online."
}

#endregion

#region Lettura anagrafica ------------------------------------------------------

Write-LogInfo "Caricamento sale da: $CsvPath"

if (-not (Test-Path -LiteralPath $CsvPath)) {
    throw "File di anagrafica non trovato: $CsvPath"
}

if ($HeaderRow -gt 1) {
    Write-LogInfo "Intestazioni attese alla riga ${HeaderRow}: le righe precedenti vengono ignorate."
    $csvLines = Get-Content -LiteralPath $CsvPath | Select-Object -Skip ($HeaderRow - 1)
    $Rooms = $csvLines | ConvertFrom-Csv -Delimiter ';'
}
else {
    $Rooms = Import-Csv -Path $CsvPath -Delimiter ';'
}

$Rooms = $Rooms |
    Where-Object { -not [string]::IsNullOrWhiteSpace((Get-RowValue -Row $_ -FieldNames $Field.Name)) }

Write-LogInfo "$($Rooms.Count) sala/e caricate dal CSV."

Write-LogInfo "Preparazione dei dati e calcolo degli alias..."

$PreparedRooms = foreach ($Room in $Rooms) {
    $rawName     = Get-RowValue -Row $Room -FieldNames $Field.Name
    $city        = Get-RowValue -Row $Room -FieldNames $Field.City
    $building    = Get-RowValue -Row $Room -FieldNames $Field.Building
    $floorRaw    = Get-RowValue -Row $Room -FieldNames $Field.Floor
    $capacityRaw = Get-RowValue -Row $Room -FieldNames $Field.Capacity
    $country     = Get-RowValue -Row $Room -FieldNames $Field.Country
    $rowPrefix   = Get-RowValue -Row $Room -FieldNames $Field.Prefix

    if (-not $country) { $country = $DefaultCountry }

    # Il prefisso arriva dalla colonna Prefix se esiste (vecchio tracciato),
    # altrimenti dal parametro: nel modello aggiornato e' gia' dentro il nome sala.
    $effectivePrefix = if ($rowPrefix) { $rowPrefix } else { $Prefix }

    # Nome breve della sala per l'alias: si scarta il prefisso e le parti di
    # localizzazione gia' presenti nel nome ("ACME - Milano Nord - Sala Verdi"
    # diventa "Sala Verdi"), cosi' l'alias non ripete citta' ed edificio.
    $shortName = $rawName

    if ($shortName -like "$effectivePrefix*") {
        $shortName = $shortName.Substring($effectivePrefix.Length)
    }

    $segments = @($shortName -split '\s+-\s+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

    if ($segments.Count -gt 0) {
        $shortName = $segments[-1]
    }

    if ([string]::IsNullOrWhiteSpace($shortName)) {
        $shortName = $rawName
        Write-LogWarning "Nome sala non scomponibile: '$rawName'. Alias costruito sul nome completo."
    }

    $capacity = $null

    if ($capacityRaw) {
        $capacityMatch = [regex]::Match($capacityRaw, '\d+')

        if ($capacityMatch.Success) {
            $capacity = [int]$capacityMatch.Value

            if ($capacityRaw -notmatch '^\s*\d+\s*$') {
                Write-LogWarning "Capienza '$capacityRaw' per la sala '$rawName' interpretata come ${capacity}: da confermare con il cliente."
            }
        }
        else {
            Write-LogWarning "Valore di capienza non valido '$capacityRaw' per la sala '$rawName': ignorato."
        }
    }

    $floor = $null

    try {
        $floor = ConvertTo-FloorNumber $floorRaw
    }
    catch {
        Write-LogWarning "$($_.Exception.Message) Sala '$rawName': piano non impostato."
    }

    [pscustomobject]@{
        RoomName           = $rawName
        MailboxName        = $rawName
        ShortName          = $shortName
        Alias              = New-RoomAlias -Parts @($effectivePrefix, $city, $building, $floorRaw, $shortName)
        PrimarySmtpAddress = $null
        City               = $city
        Building           = $building
        Floor              = $floor
        FloorLabel         = $floorRaw
        CountryOrRegion    = $country
        Office             = (@($city, $building, $floorRaw) | Where-Object { $_ }) -join ' - '
        ResourceCapacity   = $capacity
        AudioDeviceName    = Get-RowValue -Row $Room -FieldNames $Field.Audio
        VideoDeviceName    = Get-RowValue -Row $Room -FieldNames $Field.Video
        DisplayDeviceName  = Get-RowValue -Row $Room -FieldNames $Field.Display
        DeviceDescription  = Get-RowValue -Row $Room -FieldNames $Field.Description
        RequiresApproval   = Test-Yes (Get-RowValue -Row $Room -FieldNames $Field.RequiresApproval)
        Approvers          = @((Get-RowValue -Row $Room -FieldNames $Field.Approvers) -split '[,;]') |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ }
        Created            = $false
        Configured         = $false
        Skipped            = $false
    }
}

#
# Il modello aggiornato non porta piu' il piano dentro il nome sala, quindi
# nomi identici su piani diversi non sono piu' distinguibili dall'utente.
# Si aggiunge il piano al nome visualizzato solo dove serve.
#
$duplicateNames = $PreparedRooms |
    Group-Object RoomName |
    Where-Object { $_.Count -gt 1 }

foreach ($duplicate in $duplicateNames) {
    Write-LogWarning "Nome sala ripetuto $($duplicate.Count) volte nell'anagrafica: '$($duplicate.Name)'. Verra' aggiunto il piano al nome visualizzato."

    foreach ($room in $duplicate.Group) {
        if ($room.FloorLabel) {
            $room.RoomName   = "$($room.RoomName) - ($($room.FloorLabel))"
            $room.MailboxName = $room.RoomName
        }
    }
}

Write-LogInfo "$($PreparedRooms.Count) sala/e preparate. Risoluzione di alias e nomi duplicati..."

$AssignedAliases = @{}
$AssignedNames   = @{}

foreach ($PreparedRoom in $PreparedRooms) {
    $PreparedRoom.Alias = Resolve-UniqueValue -Value $PreparedRoom.Alias -MaxLength 64 -UsedValues $AssignedAliases
    $PreparedRoom.MailboxName = Resolve-UniqueValue -Value $PreparedRoom.MailboxName -MaxLength 64 -UsedValues $AssignedNames

    if ($Domain) {
        $PreparedRoom.PrimarySmtpAddress = "$($PreparedRoom.Alias)@$Domain"
    }
}

$approvalCount = @($PreparedRooms | Where-Object { $_.RequiresApproval }).Count
Write-LogInfo "Sale con prenotazione soggetta ad approvazione: $approvalCount."

$missingApprovers = @($PreparedRooms | Where-Object { $_.RequiresApproval -and $_.Approvers.Count -eq 0 })

foreach ($room in $missingApprovers) {
    Write-LogWarning "La sala '$($room.RoomName)' richiede approvazione ma non ha delegati indicati: verra' configurata in sola richiesta senza delegati."
}

#endregion

#region Controlli preliminari --------------------------------------------------

Write-LogInfo "=== CONTROLLI PRELIMINARI SULL'ANAGRAFICA ==="

# Capacity e' una delle tre proprieta' essenziali per il Room Finder insieme a
# City e Floor: senza di essa la sala non e' filtrabile per numero di posti.
$noCapacity = @($PreparedRooms | Where-Object { $null -eq $_.ResourceCapacity })

if ($noCapacity.Count -gt 0) {
    Write-LogWarning "$($noCapacity.Count) sala/e senza capienza: non saranno filtrabili per numero di posti in Room Finder."
}

$noCity = @($PreparedRooms | Where-Object { -not $_.City })

if ($noCity.Count -gt 0) {
    Write-LogWarning "$($noCity.Count) sala/e senza sede: non compariranno nel filtro per citta' e restano fuori dalle room list."
}

#
# Varianti di scrittura dell'edificio (maiuscole/minuscole o spazi) spezzerebbero
# le room list in gruppi distinti: si segnalano e si normalizza la forma piu'
# frequente.
#
$buildingVariants = $PreparedRooms |
    Where-Object { $_.Building -and $_.City } |
    Group-Object { ("$([string]$_.City)|$([string]$_.Building)").ToLowerInvariant().Trim() } |
    Where-Object { (@($_.Group.Building | Select-Object -Unique)).Count -gt 1 }

foreach ($variant in $buildingVariants) {
    $forms = @($variant.Group.Building | Group-Object | Sort-Object Count -Descending)
    $canonical = $forms[0].Name
    $others = ($forms | Select-Object -Skip 1 | ForEach-Object { "'$($_.Name)'" }) -join ', '

    Write-LogWarning "Edificio scritto in piu' modi nella sede '$($variant.Group[0].City)': $others normalizzati in '$canonical'."

    foreach ($room in $variant.Group) {
        $room.Building = $canonical
    }
}

$duplicateAliasCandidates = $PreparedRooms |
    Group-Object Alias |
    Where-Object { $_.Count -gt 1 }

foreach ($candidate in $duplicateAliasCandidates) {
    Write-LogWarning "Alias identico per $($candidate.Count) sale ('$($candidate.Name)'): possibile doppione in anagrafica da verificare con il cliente."
}

#endregion

#region FASE 1 - Creazione delle room mailbox -----------------------------------

Write-LogInfo "=== FASE 1: creazione delle room mailbox ($($PreparedRooms.Count) sale) ==="

foreach ($PreparedRoom in $PreparedRooms) {
    $roomAlias = $PreparedRoom.Alias

    try {
        $existing = Get-Mailbox -Identity $roomAlias -ErrorAction SilentlyContinue

        if ($existing) {
            Write-LogInfo "Gia' presente, verra' aggiornata: $($PreparedRoom.RoomName) ($roomAlias)."
            $PreparedRoom.Created = $true
            continue
        }

        $newMailboxParams = @{
            Room        = $true
            Name        = $PreparedRoom.MailboxName
            Alias       = $roomAlias
            DisplayName = $PreparedRoom.RoomName
        }

        if ($null -ne $PreparedRoom.ResourceCapacity) {
            $newMailboxParams.ResourceCapacity = $PreparedRoom.ResourceCapacity
        }

        if ($PreparedRoom.Office) {
            $newMailboxParams.Office = $PreparedRoom.Office
        }

        if ($PreparedRoom.PrimarySmtpAddress) {
            $newMailboxParams.PrimarySmtpAddress = $PreparedRoom.PrimarySmtpAddress
        }

        Write-LogInfo "Creazione room mailbox: $($PreparedRoom.RoomName) ($roomAlias)..."

        if ($WhatIfProvisioning) {
            Write-LogInfo "[simulazione] New-Mailbox -Room -Alias $roomAlias"
            $PreparedRoom.Created = $true
            continue
        }

        $mailboxWarnings = @()
        New-Mailbox @newMailboxParams -WarningVariable mailboxWarnings -WarningAction SilentlyContinue | Out-Null
        Write-CommandWarnings -Command 'New-Mailbox' -Warnings $mailboxWarnings

        $PreparedRoom.Created = $true
        Write-LogSuccess "Creata: $($PreparedRoom.RoomName) ($roomAlias)"
    }
    catch {
        Write-LogError "Creazione fallita: $($PreparedRoom.RoomName) ($roomAlias)"
        Write-LogError $_.Exception.Message
    }
}

$createdCount = @($PreparedRooms | Where-Object { $_.Created }).Count
Write-LogInfo "Fase 1 completata: $createdCount sala/e disponibili per la configurazione."

#endregion

#region FASE 2 - Configurazione Place e calendar processing ---------------------

Write-LogInfo "=== FASE 2: configurazione attributi e regole di prenotazione ==="

if (-not $WhatIfProvisioning -and $PropagationWaitSeconds -gt 0) {
    Write-LogInfo "Attesa di $PropagationWaitSeconds secondi per la propagazione degli oggetti in directory..."
    Start-Sleep -Seconds $PropagationWaitSeconds
}

$Retry = [System.Collections.Generic.List[object]]::new()

foreach ($PreparedRoom in $PreparedRooms) {
    if (-not $PreparedRoom.Created) {
        $PreparedRoom.Skipped = $true
        continue
    }

    $roomAlias = $PreparedRoom.Alias

    try {
        if ($WhatIfProvisioning) {
            Write-LogInfo "[simulazione] Set-Place / Set-CalendarProcessing su $roomAlias"
            $PreparedRoom.Configured = $true
            continue
        }

        if (-not (Wait-MailboxAvailable -Identity $roomAlias -Attempts $MaxAttempts -DelaySeconds $RetryDelaySeconds)) {
            Write-LogWarning "Oggetto non ancora risolvibile dopo $MaxAttempts tentativi: $($PreparedRoom.RoomName) ($roomAlias). Messa in coda di rilavorazione."
            $Retry.Add($PreparedRoom)
            continue
        }

        Set-Mailbox -Identity $roomAlias -DisplayName $PreparedRoom.RoomName -ErrorAction Stop

        #
        # Attributi Place (City, Building, Floor, FloorLabel, Country, dispositivi)
        #
        $setPlaceParams = @{ Identity = $roomAlias }

        foreach ($property in 'City', 'Building', 'Floor', 'FloorLabel', 'CountryOrRegion') {
            $value = $PreparedRoom.$property

            if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                $setPlaceParams[$property] = $value
            }
        }

        if ($null -ne $PreparedRoom.ResourceCapacity) {
            $setPlaceParams.Capacity = $PreparedRoom.ResourceCapacity
        }

        foreach ($property in 'AudioDeviceName', 'VideoDeviceName', 'DisplayDeviceName') {
            if (Test-Yes $PreparedRoom.$property) {
                $setPlaceParams[$property] = 'Yes'
            }
        }

        if ($setPlaceParams.Count -gt 1) {
            $placeWarnings = @()
            Set-Place @setPlaceParams -WarningVariable placeWarnings -WarningAction SilentlyContinue -ErrorAction Stop
            Write-CommandWarnings -Command 'Set-Place' -Warnings $placeWarnings

            if ($placeWarnings.Count -eq 0) {
                Write-LogInfo "Attributi Place configurati per $($PreparedRoom.RoomName)."
            }
        }

        #
        # Regole di prenotazione
        #
        $requiresApproval = $PreparedRoom.RequiresApproval -and $PreparedRoom.Approvers.Count -gt 0

        $calendarParams = @{
            Identity           = $roomAlias
            AutomateProcessing = 'AutoAccept'
            AllBookInPolicy    = -not $PreparedRoom.RequiresApproval
            AllRequestInPolicy = [bool]$PreparedRoom.RequiresApproval
        }

        if ($requiresApproval) {
            $calendarParams.ResourceDelegates = $PreparedRoom.Approvers
        }

        $calendarWarnings = @()
        Set-CalendarProcessing @calendarParams -WarningVariable calendarWarnings -WarningAction SilentlyContinue
        Write-CommandWarnings -Command 'Set-CalendarProcessing' -Warnings $calendarWarnings

        if ($PreparedRoom.RequiresApproval) {
            $delegates = if ($requiresApproval) { $PreparedRoom.Approvers -join ', ' } else { 'nessun delegato indicato' }
            Write-LogInfo "Prenotazione soggetta ad approvazione per $($PreparedRoom.RoomName) ($delegates)."
        }

        $PreparedRoom.Configured = $true
        Write-LogSuccess "Configurata: $($PreparedRoom.RoomName) ($roomAlias)"
    }
    catch {
        Write-LogError "Configurazione fallita: $($PreparedRoom.RoomName) ($roomAlias)"
        Write-LogError $_.Exception.Message
        $Retry.Add($PreparedRoom)
    }
}

#
# Coda di rilavorazione: una sala non pronta al primo passaggio non viene scartata.
#
if ($Retry.Count -gt 0 -and -not $WhatIfProvisioning) {
    Write-LogWarning "$($Retry.Count) sala/e in coda di rilavorazione. Nuovo tentativo tra $RetryDelaySeconds secondi..."
    Start-Sleep -Seconds $RetryDelaySeconds

    foreach ($PreparedRoom in $Retry) {
        $roomAlias = $PreparedRoom.Alias

        try {
            if (-not (Wait-MailboxAvailable -Identity $roomAlias -Attempts $MaxAttempts -DelaySeconds $RetryDelaySeconds)) {
                Write-LogError "Non risolvibile anche in rilavorazione: $($PreparedRoom.RoomName) ($roomAlias)."
                continue
            }

            $setPlaceParams = @{ Identity = $roomAlias }

            foreach ($property in 'City', 'Building', 'Floor', 'FloorLabel', 'CountryOrRegion') {
                $value = $PreparedRoom.$property

                if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
                    $setPlaceParams[$property] = $value
                }
            }

            if ($null -ne $PreparedRoom.ResourceCapacity) {
                $setPlaceParams.Capacity = $PreparedRoom.ResourceCapacity
            }

            foreach ($property in 'AudioDeviceName', 'VideoDeviceName', 'DisplayDeviceName') {
                if (Test-Yes $PreparedRoom.$property) {
                    $setPlaceParams[$property] = 'Yes'
                }
            }

            Set-Place @setPlaceParams -ErrorAction Stop

            $requiresApproval = $PreparedRoom.RequiresApproval -and $PreparedRoom.Approvers.Count -gt 0

            $calendarParams = @{
                Identity           = $roomAlias
                AutomateProcessing = 'AutoAccept'
                AllBookInPolicy    = -not $PreparedRoom.RequiresApproval
                AllRequestInPolicy = [bool]$PreparedRoom.RequiresApproval
            }

            if ($requiresApproval) {
                $calendarParams.ResourceDelegates = $PreparedRoom.Approvers
            }

            Set-CalendarProcessing @calendarParams -WarningAction SilentlyContinue

            $PreparedRoom.Configured = $true
            Write-LogSuccess "Configurata in rilavorazione: $($PreparedRoom.RoomName) ($roomAlias)"
        }
        catch {
            Write-LogError "Configurazione fallita anche in rilavorazione: $($PreparedRoom.RoomName) ($roomAlias)"
            Write-LogError $_.Exception.Message
        }
    }
}

#endregion

#region FASE 3 - Room list ------------------------------------------------------

if ($SkipRoomLists) {
    Write-LogInfo "=== FASE 3 saltata su richiesta (-SkipRoomLists) ==="
}
else {
    Write-LogInfo "=== FASE 3: creazione delle room list e inserimento delle sale ==="

    #
    # Room Finder usa le room list come valori del filtro Building e ogni lista
    # deve contenere sale di una sola citta'. Si raggruppa quindi per
    # citta' + edificio, spezzando le liste oltre $MaxRoomsPerList membri.
    #
    $groups = $PreparedRooms |
        Where-Object { $_.Configured -and $_.City } |
        Group-Object { ("$([string]$_.City)|$([string]$_.Building)").ToLowerInvariant().Trim() }

    foreach ($group in $groups) {
        # Le etichette si prendono dal primo record del gruppo, non dalla chiave
        # normalizzata, cosi' la room list conserva la grafia dell'anagrafica.
        $city = $group.Group[0].City
        $building = $group.Group[0].Building

        $chunkIndex = 0
        $chunks = @()
        $current = @()

        foreach ($room in $group.Group) {
            $current += $room

            if ($current.Count -ge $MaxRoomsPerList) {
                $chunks += , $current
                $current = @()
            }
        }

        if ($current.Count -gt 0) {
            $chunks += , $current
        }

        foreach ($chunk in $chunks) {
            $chunkIndex++

            $listLabelParts = @($Prefix, $city)

            if ($building) {
                $listLabelParts += $building
            }

            if ($chunks.Count -gt 1) {
                $listLabelParts += "$chunkIndex"
            }

            $listDisplayName = ($listLabelParts -join ' - ')
            $listAlias = New-RoomAlias -Parts $listLabelParts

            try {
                $existingList = Get-DistributionGroup -Identity $listAlias -ErrorAction SilentlyContinue

                if (-not $existingList) {
                    if ($WhatIfProvisioning) {
                        Write-LogInfo "[simulazione] New-DistributionGroup -RoomList -Alias $listAlias ($($chunk.Count) sale)"
                    }
                    else {
                        $newListParams = @{
                            Name        = $listDisplayName
                            DisplayName = $listDisplayName
                            Alias       = $listAlias
                            RoomList    = $true
                        }

                        if ($Domain) {
                            $newListParams.PrimarySmtpAddress = "$listAlias@$Domain"
                        }

                        New-DistributionGroup @newListParams | Out-Null
                        Write-LogSuccess "Room list creata: $listDisplayName ($listAlias)"
                    }
                }
                else {
                    Write-LogInfo "Room list gia' presente: $listDisplayName ($listAlias)."
                }

                if ($WhatIfProvisioning) {
                    continue
                }

                $currentMembers = @()

                if ($existingList) {
                    $currentMembers = @(Get-DistributionGroupMember -Identity $listAlias -ResultSize Unlimited |
                        Select-Object -ExpandProperty Alias)
                }

                foreach ($room in $chunk) {
                    if ($currentMembers -contains $room.Alias) {
                        continue
                    }

                    try {
                        Add-DistributionGroupMember -Identity $listAlias -Member $room.Alias -ErrorAction Stop
                    }
                    catch {
                        Write-LogWarning "Inserimento in room list fallito per $($room.RoomName): $($_.Exception.Message)"
                    }
                }

                Write-LogInfo "Room list $listDisplayName allineata ($($chunk.Count) sale)."
            }
            catch {
                Write-LogError "Gestione della room list fallita: $listDisplayName"
                Write-LogError $_.Exception.Message
            }
        }
    }
}

#endregion

#region Riepilogo ---------------------------------------------------------------

$total      = $PreparedRooms.Count
$configured = @($PreparedRooms | Where-Object { $_.Configured }).Count
$failed     = @($PreparedRooms | Where-Object { -not $_.Configured }).Count

Write-LogInfo "=== RIEPILOGO ==="
$approvalConfigured = @($PreparedRooms | Where-Object { $_.Configured -and $_.RequiresApproval }).Count

Write-LogInfo "Sale in anagrafica:        $total"
Write-LogInfo "Sale configurate:          $configured"
Write-LogInfo "Di cui ad approvazione:    $approvalConfigured"

if ($failed -gt 0) {
    Write-LogWarning "Sale non configurate:      $failed (rilanciare lo script per completarle)"
}
else {
    Write-LogSuccess "Nessuna sala rimasta incompleta."
}

$reportPath = Join-Path $PSScriptRoot "New-RoomMailboxes_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

$PreparedRooms |
    Select-Object RoomName, ShortName, Alias, City, Building, Floor, FloorLabel, ResourceCapacity,
                  RequiresApproval, @{ Name = 'Approvers'; Expression = { $_.Approvers -join ',' } },
                  Created, Configured |
    Export-Csv -Path $reportPath -NoTypeInformation -Delimiter ';' -Encoding UTF8

Write-LogInfo "Report di esecuzione salvato in: $reportPath"
Write-LogInfo "Nota: le modifiche possono richiedere da 24 a 48 ore prima di essere visibili in Room Finder."

Write-LogInfo "Disconnessione da Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false
Write-LogSuccess "Operazione completata."

#endregion
