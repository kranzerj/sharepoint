# Setze die Konsolen-Ausgabekodierung auf UTF-8 (für korrekte Anzeige von Unicode-Zeichen)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$siteUrl     = Read-Host "enter site URL (e.g. https://contoso.sharepoint.com/sites/ITTeam)"
$libraryName = Read-Host "enter name of library (e.g. Dokuments)"
$folderPath  = Read-Host "enter Path / Folder at libary (z.B.: 'FolderA')"

# 2. Verbindung zu Microsoft Graph herstellen
Connect-MgGraph -Scopes "Sites.Read.All", "Files.Read.All"

# 3. Tenant‑ID ermitteln
$org      = Get-MgOrganization | Select-Object -First 1
$tenantId = $org.Id

# 4. Site‑Objekt abrufen
$uri             = [System.Uri]$siteUrl
$hostname        = $uri.Host
$sitePath        = $uri.AbsolutePath.Trim('/')
$siteIdFormatted = "${hostname}:/$sitePath"
$site            = Get-MgSite -SiteId $siteIdFormatted

# 5. Bibliothek (List) abrufen und sharepointIds extrahieren
$lists   = Get-MgSiteList -SiteId $site.Id -Property "displayName,sharepointIds"
$library = $lists | Where-Object DisplayName -EQ $libraryName

if (-not $library) {
    Write-Error "Bibliothek '$libraryName' nicht gefunden."
    exit
}

if ($library.sharepointIds) {
    $spIds = $library.sharepointIds
} else {
    $spIds = (Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/sites/$($site.Id)/lists/$($library.Id)?`$select=sharepointIds"
    ).sharepointIds
}

# 6. Drive (Document Library) ermitteln
$drive      = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$($site.Id)/drive"
$driveId    = $drive.id

# 7. Ordner‑DriveItem holen und sharepointIds auslesen
#    -> wir holen nur die id‑Fazette, daher $select=sharepointIds
$encodedPath      = [uri]::EscapeDataString($folderPath)
$folderRequestUri = "https://graph.microsoft.com/v1.0/drives/$($driveId)/root:/$($encodedPath):/?`$select=sharepointIds"
$folder           = Invoke-MgGraphRequest -Method GET -Uri $folderRequestUri
$folderSpIds      = $folder.sharepointIds

if (-not $folderSpIds) {
    Write-Error "Für Ordner '$folderPath' keine sharepointIds gefunden."
    exit
}

# 8. Sync‑URL zusammensetzen
$finalUrl = "tenantId=$tenantId" + `
    "&siteId={" + $spIds.siteId          + "}" + `
    "&webId={"  + $spIds.webId           + "}" + `
    "&listId="  + $spIds.listId               + `
    "&folderId=" + $folderSpIds.listItemUniqueId + `
    "&webUrl="  + $site.WebUrl                + `
    "&version=1"

# 9. Ergebnis formatiert ausgeben
Write-Host "`n✅ Die OneDrive‑Sync‑URL lautet:" -ForegroundColor Green
Write-Host $finalUrl                                 -ForegroundColor Cyan

