# Setze die Konsolen-Ausgabekodierung auf UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Abfrage der SharePoint Site URL und des Bibliotheksnamens
$siteUrl = Read-Host "Bitte geben Sie die SharePoint Site URL ein (z.B.: https://contoso.sharepoint.com/sites/ITTeam)"
$libraryName = Read-Host "Bitte geben Sie den Namen der Bibliothek ein (z.B.: Dokumente)"

# Aufbau der Verbindung zu Microsoft Graph mit den benÃ¶tigten Scopes
Connect-MgGraph -Scopes "Sites.Read.All", "Files.Read.All"

# Abrufen der Organisationsdaten, um die Tenant-ID zu erhalten
$org = Get-MgOrganization | Select-Object -First 1
if (-not $org) {
    Write-Error "Organisationsdaten konnten nicht abgerufen werden."
    exit
}
$tenantId = $org.Id

# Aufbereitung der SharePoint Site URL zur Abfrage des Site-Objekts
try {
    $uri = [System.Uri]$siteUrl
} catch {
    Write-Error "Die eingegebene URL '$siteUrl' ist ungÃ¼ltig."
    exit
}
$hostname = $uri.Host                     # z.B.: contoso.sharepoint.com
$sitePath = $uri.AbsolutePath.Trim('/')    # z.B.: sites/ITTeam

# Erstellen des zusammengesetzten Site-ID-Formats
$siteIdFormatted = "${hostname}:/$sitePath"

# Abrufen des Site-Objekts Ã¼ber Microsoft Graph
$site = Get-MgSite -SiteId $siteIdFormatted
if (-not $site) {
    Write-Error "Site mit der URL '$siteUrl' wurde nicht gefunden."
    exit
}

# Abrufen aller Listen (Dokumentbibliotheken sind Listen)
try {
    $lists = Get-MgSiteList -SiteId $site.Id -Property "displayName,sharepointIds"
} catch {
    Write-Error "Fehler beim Abrufen der Listen: $_"
    exit
}

if (-not $lists) {
    Write-Error "Keine Listen (Bibliotheken) in der Site '$siteUrl' gefunden."
    exit
}

# Auswahl der Bibliothek anhand des angegebenen Namens (exakte Ãœbereinstimmung)
$library = $lists | Where-Object { $_.DisplayName -eq $libraryName }
if (-not $library) {
    Write-Error "Die Bibliothek '$libraryName' wurde in der Site '$siteUrl' nicht gefunden."
    exit
}

# Falls sharepointIds nicht direkt in der Liste vorhanden sind, erfolgt ein erneuter Abruf der Detailinformationen
if (-not $library.sharepointIds) {
    Write-Verbose "sharepointIds wurden in der Liste nicht gefunden. Erneuter Abruf der Bibliotheksdetails..."
    $libraryDetails = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($site.Id)/lists/$($library.Id)?`$select=displayName,sharepointIds"
    $spIds = $libraryDetails.sharepointIds
} else {
    $spIds = $library.sharepointIds
}

if (-not $spIds) {
    Write-Error "sharepointIds wurden fÃ¼r die Bibliothek '$libraryName' nicht gefunden."
    exit
}

# Aufbau des finalen Identifier-Strings im gewÃ¼nschten Format:
# tenantId={Tenant-ID}&siteId={Site-ID}&webId={Web-ID}&listId={List-ID}&webUrl={SharePointSiteURL}&version=1
$sharePointLibraryId = "tenantId=$tenantId" + `
    "&siteId={" + $spIds.siteId + "}" + `
    "&webId={" + $spIds.webId + "}" + `
    "&listId=" + $spIds.listId + `
    "&webUrl=" + $site.WebUrl + `
    "&version=1"

# Formattierte Ausgabe
Write-Host "Die SharePoint Library ID lautet:" -ForegroundColor Green
Write-Host $sharePointLibraryId -ForegroundColor Cyan
