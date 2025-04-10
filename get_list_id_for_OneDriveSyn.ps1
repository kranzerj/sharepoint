
# Microsoft Graph verbinden
Connect-MgGraph -Scopes "Sites.Read.All"

# Eingaben abfragen
$sitePath = Read-Host "Gib den Pfad zur SharePoint-Seite ein (z. B. /sites/ITTS)"
$libName = Read-Host "Gib den Namen der Dokumentbibliothek ein (z. B. Dokumente)"
$webUrl = "https://xxxxxx.sharepoint.com$sitePath"

# Tenant-ID holen
$org = Get-MgOrganization
$tenantId = $org.Id

# Site abrufen
$site = Get-MgSite -SiteId "xxxxxx.sharepoint.com:${sitePath}:"

if (-not $site) {
    Write-Error "Seite nicht gefunden. Bitte überprüfe den Pfad."
    exit
}

# SiteId und WebId extrahieren
$siteId = $site.SiteId
$webId = $site.Root.WebId

# Listen abrufen und passende finden
$lists = Get-MgSiteList -SiteId $site.Id
$library = $lists | Where-Object { $_.DisplayName -eq $libName }

if (-not $library) {
    Write-Error "Bibliothek '$libName' nicht gefunden."
    exit
}

$listId = $library.Id

# Finale URL erzeugen
$finalUrl = "tenantId=$tenantId&siteId=$siteId&webId=$webId&listId=$listId&webUrl=$webUrl&version=1"

Write-Host "`n✅ Die OneDrive-Link-URL lautet:" -ForegroundColor Green
Write-Host $finalUrl -ForegroundColor Cyan
