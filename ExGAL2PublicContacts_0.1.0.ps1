#requires -Version 5.1
<#
========================================================================
 ExGAL2PublicContacts
 Version: 0.1.0

 Read-Only development version.
 Reads Exchange UserMailbox objects, AD account status and existing
 Public Folder Contacts. NO contact is created, changed or deleted.

 Project initiator: Steffen Pelzetter
 Developed with: OpenAI ChatGPT
 Target: Exchange 2019 / Exchange 2019 SE, PowerShell 5.1
========================================================================
#>

[CmdletBinding()]
param([switch]$VerboseLog)

# ---------------------------- Configuration ----------------------------

$Script:Version          = '0.1.0'
$Script:EwsDll           = 'D:\ExchangeServer\Bin\Microsoft.Exchange.WebServices.dll'
$Script:EwsUrl           = 'https://localhost/EWS/Exchange.asmx'
$Script:PublicFolderPath = 'Kontakte\Verwaltung'
$Script:LogFolder        = 'C:\Logs\ExGAL2PublicContacts'

# -------------------------- Global variables ---------------------------

$Script:LogFile = $null
$Script:StartTime = Get-Date
$Script:Service = $null
$Script:Statistics = [ordered]@{
    ExchangeUsers=0; ActiveUsers=0; DisabledUsers=0; UsersWithoutSmtp=0
    PublicContacts=0; ContactsWithSmtp=0; ContactsWithoutSmtp=0
    Matched=0; Missing=0; DisabledMatches=0; DuplicateSmtp=0; Errors=0
}

# ------------------------------- Logging -------------------------------

function Initialize-Log {
    if (!(Test-Path -LiteralPath $Script:LogFolder)) {
        New-Item -Path $Script:LogFolder -ItemType Directory -Force | Out-Null
    }
    $Script:LogFile = Join-Path $Script:LogFolder (
        'ExGAL2PublicContacts_{0:yyyyMMdd}.log' -f (Get-Date))
}

function Write-Log {
    param(
        [ValidateSet('INFO','DEBUG','OK','NEW','UPDATE','DELETE','WARN','ERROR')]
        [string]$Level,
        [string]$Message
    )
    $line = '{0} [{1,-6}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'),
        $Level, $Message
    Add-Content -LiteralPath $Script:LogFile -Value $line -Encoding UTF8

    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'OK'    { Write-Host $line -ForegroundColor Green }
        'NEW'   { Write-Host $line -ForegroundColor Cyan }
        'UPDATE'{ Write-Host $line -ForegroundColor Cyan }
        'DELETE'{ Write-Host $line -ForegroundColor Magenta }
        default { Write-Host $line }
    }
}

function Write-Header {
    Write-Log INFO  ('=' * 72)
    Write-Log INFO  "ExGAL2PublicContacts Version $Script:Version"
    Write-Log INFO  'READ-ONLY TESTLAUF - KEINE AENDERUNGEN'
    Write-Log INFO  ('=' * 72)
    Write-Log INFO  "Computer       : $env:COMPUTERNAME"
    Write-Log INFO  "PowerShell     : $($PSVersionTable.PSVersion)"
    Write-Log INFO  "EWS DLL        : $Script:EwsDll"
    Write-Log INFO  "EWS URL        : $Script:EwsUrl"
    Write-Log INFO  "Public Folder  : $Script:PublicFolderPath"
    Write-Log INFO  "Logfile        : $Script:LogFile"
}

# ------------------------------- Helpers -------------------------------

function Normalize-SmtpAddress {
    param([AllowNull()][string]$Address)
    if ([string]::IsNullOrWhiteSpace($Address)) { return $null }
    return $Address.Trim().ToLowerInvariant()
}

function Get-MailboxSmtpAddress {
    param($Mailbox)
    if ($null -eq $Mailbox.PrimarySmtpAddress) { return $null }
    return Normalize-SmtpAddress $Mailbox.PrimarySmtpAddress.ToString()
}

# ------------------------------- EWS -----------------------------------

function Connect-Ews {
    Write-Log INFO 'Lade EWS Managed API...'

    if (!(Test-Path -LiteralPath $Script:EwsDll)) {
        throw "EWS DLL nicht gefunden: $Script:EwsDll"
    }

    try { Add-Type -Path $Script:EwsDll -ErrorAction Stop }
    catch {
        if (!('Microsoft.Exchange.WebServices.Data.ExchangeService' -as [type])) { throw }
    }

    Write-Log OK 'EWS Managed API geladen.'
    Write-Log INFO 'Erstelle EWS-Verbindung mit Windows Authentication...'

    $version = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
    $Script:Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($version)
    $Script:Service.UseDefaultCredentials = $true
    $Script:Service.Url = $Script:EwsUrl

    # Nur fuer den lokalen Entwicklungs-/Testlauf. Spaeter entfernen.
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {
        param($sender,$certificate,$chain,$sslPolicyErrors)
        $true
    }

    Write-Log OK 'EWS-Verbindung erstellt.'
}

function Get-PublicFolderChild {
    param(
        [Microsoft.Exchange.WebServices.Data.Folder]$ParentFolder,
        [string]$Name
    )

    $view = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
    $view.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
    $view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

    foreach ($folder in $ParentFolder.FindFolders($view).Folders) {
        if ($folder.DisplayName -ieq $Name) { return $folder }
    }
    return $null
}

function Get-PublicFolderByPath {
    param([string]$Path)

    $parts = $Path -split '\\' | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
    if (!$parts) { throw 'Public-Folder-Pfad ist leer.' }

    Write-Log INFO 'Oeffne Public Folder Root...'
    $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(
        $Script:Service,
        [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)
    Write-Log OK 'Public Folder Root gefunden.'

    foreach ($part in $parts) {
        Write-Log INFO "Suche Public-Folder: $part"
        $next = Get-PublicFolderChild $folder $part
        if ($null -eq $next) {
            throw "Public-Folder '$part' wurde unter '$($folder.DisplayName)' nicht gefunden."
        }
        Write-Log OK "Gefunden: $($next.DisplayName) [$($next.FolderClass)]"
        $folder = $next
    }

    if ($folder.FolderClass -ne 'IPF.Contact') {
        Write-Log WARN "Zielordner ist '$($folder.FolderClass)', erwartet wurde IPF.Contact."
    }
    return $folder
}

# -------------------------- Public Contacts ----------------------------

function Get-PublicContacts {
    param([Microsoft.Exchange.WebServices.Data.Folder]$Folder)

    Write-Log INFO 'Lese vorhandene Public Contacts...'
    $contacts = New-Object System.Collections.Generic.List[object]

    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)
    $view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
    $view.PropertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass)
    $view.PropertySet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)

    $offset = 0
    do {
        $view.Offset = $offset
        $result = $Folder.FindItems($view)

        foreach ($item in $result.Items) {
            if ($item.ItemClass -ne 'IPM.Contact') {
                Write-Log DEBUG "Ueberspringe '$($item.Subject)' [$($item.ItemClass)]."
                continue
            }

            try {
                $props = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
                    [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $contact = [Microsoft.Exchange.WebServices.Data.Contact]::Bind(
                    $Script:Service,$item.Id,$props)
                $contacts.Add($contact)
                $Script:Statistics.PublicContacts++
            }
            catch {
                $Script:Statistics.Errors++
                Write-Log ERROR "Kontakt '$($item.Subject)' konnte nicht gelesen werden: $($_.Exception.Message)"
            }
        }

        $offset += $result.Items.Count
    } while ($result.MoreAvailable)

    Write-Log OK "$($Script:Statistics.PublicContacts) Public Contacts gelesen."
    return $contacts
}

function Get-ContactSmtpAddress {
    param([Microsoft.Exchange.WebServices.Data.Contact]$Contact)

    try {
        $email = $Contact.EmailAddresses[
            [Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1]
        if ($null -ne $email) { return Normalize-SmtpAddress $email.Address }
    }
    catch {
        Write-Log WARN "E-Mail-Adresse von '$($Contact.DisplayName)' konnte nicht gelesen werden: $($_.Exception.Message)"
    }
    return $null
}

function New-PublicContactIndex {
    param([System.Collections.Generic.List[object]]$Contacts)

    Write-Log INFO 'Erstelle Kontaktindex nach SMTP-Adresse...'
    $index = @{}

    foreach ($contact in $Contacts) {
        $smtp = Get-ContactSmtpAddress $contact

        if ([string]::IsNullOrWhiteSpace($smtp)) {
            $Script:Statistics.ContactsWithoutSmtp++
            Write-Log WARN "Kontakt ohne SMTP-Adresse: '$($contact.DisplayName)'"
            continue
        }

        $Script:Statistics.ContactsWithSmtp++

        if ($index.ContainsKey($smtp)) {
            $Script:Statistics.DuplicateSmtp++
            Write-Log WARN "Doppelte SMTP-Adresse im Public Folder: $smtp"
            continue
        }

        $index[$smtp] = $contact
    }

    Write-Log OK "$($index.Count) eindeutige SMTP-Adressen indiziert."
    return $index
}

# -------------------------- AD / Exchange ------------------------------

function Get-ExchangeUsers {
    Write-Log INFO 'Lese Exchange UserMailbox...'

    if (!(Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
        throw 'Get-Mailbox nicht verfuegbar. Bitte Exchange Management Shell verwenden.'
    }
    if (!(Get-Command Get-ADUser -ErrorAction SilentlyContinue)) {
        throw 'Get-ADUser nicht verfuegbar. ActiveDirectory-Modul fehlt.'
    }

    $mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited
    $users = New-Object System.Collections.Generic.List[object]

    foreach ($mailbox in $mailboxes) {
        $Script:Statistics.ExchangeUsers++
        $smtp = Get-MailboxSmtpAddress $mailbox

        if ([string]::IsNullOrWhiteSpace($smtp)) {
            $Script:Statistics.UsersWithoutSmtp++
            Write-Log WARN "Mailbox ohne primäre SMTP-Adresse: $($mailbox.Identity)"
            continue
        }

        try {
            $ad = Get-ADUser -Identity $mailbox.DistinguishedName -Properties `
                Enabled,ObjectGUID,Department,Title,Office,Mobile,TelephoneNumber,Fax,`
                StreetAddress,PostalCode,City,State,Country,Company,GivenName,Surname,`
                DisplayName,WebPage -ErrorAction Stop

            $user = [PSCustomObject]@{
                Identity=$mailbox.Identity.ToString()
                DistinguishedName=$mailbox.DistinguishedName
                SamAccountName=$ad.SamAccountName
                ObjectGUID=$ad.ObjectGUID
                Enabled=[bool]$ad.Enabled
                PrimarySmtpAddress=$smtp
                DisplayName=$ad.DisplayName
                GivenName=$ad.GivenName
                Surname=$ad.Surname
                Company=$ad.Company
                Department=$ad.Department
                Title=$ad.Title
                Office=$ad.Office
                TelephoneNumber=$ad.TelephoneNumber
                Mobile=$ad.Mobile
                Fax=$ad.Fax
                StreetAddress=$ad.StreetAddress
                PostalCode=$ad.PostalCode
                City=$ad.City
                State=$ad.State
                Country=$ad.Country
                WebPage=$ad.WebPage
            }

            $users.Add($user)

            if ($user.Enabled) {
                $Script:Statistics.ActiveUsers++
            } else {
                $Script:Statistics.DisabledUsers++
                Write-Log WARN "AD-Konto deaktiviert: $($user.DisplayName) <$smtp>"
            }
        }
        catch {
            $Script:Statistics.Errors++
            Write-Log ERROR "AD-Daten fuer '$($mailbox.Identity)' nicht lesbar: $($_.Exception.Message)"
        }
    }

    Write-Log OK "$($Script:Statistics.ExchangeUsers) UserMailbox gelesen."
    Write-Log INFO "Aktiv: $($Script:Statistics.ActiveUsers) / Deaktiviert: $($Script:Statistics.DisabledUsers)"
    return $users
}

# ---------------------------- Comparison -------------------------------

function Compare-UsersWithContacts {
    param(
        [System.Collections.Generic.List[object]]$Users,
        [hashtable]$ContactIndex
    )

    Write-Log INFO 'Vergleiche Benutzer mit Public Contacts...'

    foreach ($user in $Users) {
        $smtp = $user.PrimarySmtpAddress

        if (!$user.Enabled) {
            if ($ContactIndex.ContainsKey($smtp)) {
                $Script:Statistics.DisabledMatches++
                Write-Log DELETE "[READ-ONLY] Würde entfernen: $($user.DisplayName) <$smtp> - AD-Konto deaktiviert."
            }
            continue
        }

        if ($ContactIndex.ContainsKey($smtp)) {
            $Script:Statistics.Matched++
            $contact = $ContactIndex[$smtp]
            Write-Log OK "Zuordnung OK: $($user.DisplayName) <$smtp> -> '$($contact.DisplayName)'"
        } else {
            $Script:Statistics.Missing++
            Write-Log NEW "[READ-ONLY] Würde anlegen: $($user.DisplayName) <$smtp>"
        }
    }
}

# ----------------------------- Statistics ------------------------------

function Show-Statistics {
    $duration = (Get-Date) - $Script:StartTime
    Write-Log INFO ('=' * 72)
    Write-Log INFO 'Zusammenfassung'
    Write-Log INFO ('-' * 72)
    Write-Log INFO "Exchange UserMailbox     : $($Script:Statistics.ExchangeUsers)"
    Write-Log INFO "Aktive Benutzer          : $($Script:Statistics.ActiveUsers)"
    Write-Log INFO "Deaktivierte Benutzer    : $($Script:Statistics.DisabledUsers)"
    Write-Log INFO "Benutzer ohne SMTP       : $($Script:Statistics.UsersWithoutSmtp)"
    Write-Log INFO "Public Contacts          : $($Script:Statistics.PublicContacts)"
    Write-Log INFO "Contacts mit SMTP        : $($Script:Statistics.ContactsWithSmtp)"
    Write-Log INFO "Contacts ohne SMTP       : $($Script:Statistics.ContactsWithoutSmtp)"
    Write-Log INFO "Zuordnung OK             : $($Script:Statistics.Matched)"
    Write-Log INFO "Fehlende Kontakte        : $($Script:Statistics.Missing)"
    Write-Log INFO "Deaktivierte Kontakte    : $($Script:Statistics.DisabledMatches)"
    Write-Log INFO "Doppelte SMTP-Adressen   : $($Script:Statistics.DuplicateSmtp)"
    Write-Log INFO "Fehler                   : $($Script:Statistics.Errors)"
    Write-Log INFO ("Laufzeit                 : {0:hh\:mm\:ss}" -f $duration)
    Write-Log INFO ('=' * 72)
}

# -------------------------------- Main ---------------------------------

try {
    Initialize-Log
    Write-Header
    if ($VerboseLog) { Write-Log DEBUG 'Verbose logging aktiviert.' }

    Connect-Ews
    $targetFolder = Get-PublicFolderByPath $Script:PublicFolderPath
    Write-Log INFO "Zielordner: '$($targetFolder.DisplayName)' / $($targetFolder.FolderClass)"

    $contacts = Get-PublicContacts $targetFolder
    $contactIndex = New-PublicContactIndex $contacts

    $users = Get-ExchangeUsers
    Compare-UsersWithContacts $users $contactIndex

    Show-Statistics
    Write-Log OK 'Read-Only-Testlauf erfolgreich beendet.'
}
catch {
    if ($Script:LogFile) {
        Write-Log ERROR $_.Exception.Message
        Write-Log ERROR $_.ScriptStackTrace
        Show-Statistics
    } else {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host $_.ScriptStackTrace -ForegroundColor Red
    }
    exit 1
}
