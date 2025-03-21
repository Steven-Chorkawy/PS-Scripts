# Do not include these sites in the report.
$EXCLUDED_SITE_URLS = @(
    "https://claringtonnet-my.sharepoint.com/", 
    "https://claringtonnet.sharepoint.com/portals/hub",
    "https://claringtonnet.sharepoint.com/portals/Community",
    "https://claringtonnet.sharepoint.com/search",
    "https://claringtonnet.sharepoint.com/sites/ClaringtonAppCatalog",
    "https://claringtonnet.sharepoint.com/portals/SharePoint-Invoice-Demo"
)

Clear-Host
# Connect to SharePoint tenant and retrieve list of sites
Connect-PnPOnline -Url "https://claringtonnet-admin.sharepoint.com" -Interactive

# Get all SharePoint sites exluding selected sites.
$ALL_SITES = Get-PnPTenantSite | Where-Object { $_.Url -notin $EXCLUDED_SITE_URLS }

Write-Host "'$($ALL_SITES.Count)' Sites Found"

for ($allSiteIndex = 0; $allSiteIndex -lt $ALL_SITES.Count; $allSiteIndex++) {
    $currentSite = $ALL_SITES[$allSiteIndex]
    Write-Host "`n`n------------------------------------------------------------------------------------------------------------------"
    Write-Host "$($allSiteIndex+1)/$($ALL_SITES.Count) Current Site: '$($currentSite.Title)' - $($currentSite.Status) - $($currentSite.Url)"
    # Connect to current site
    Connect-PnPOnline -Url $currentSite.Url -Interactive
    
    $documentLibraries = Get-PnPList -Includes DefaultView | Where-Object { $_.Hidden -eq $false -and $_.Title -ne "Form Templates" -and $_.Title -ne "Events" -and $_.Title -ne "Site Assets" -and $_.Title -ne "Style Library" -and $_.Title -ne "Site Pages" -and $_.EntityTypeName -ne "Shared_x0020_Documents" -and $_.EntityTypeName -ne "EventsList" }

    for ($libraryIndex = 0; $libraryIndex -lt $documentLibraries.Count; $libraryIndex++) {
        $currentLibrary = $documentLibraries[$libraryIndex]
        Write-Host "`t$($libraryIndex+1)/$($documentLibraries.Count) - $($currentLibrary.Title) - $($currentLibrary.ItemCount) Items"
        $libraryContentTypes = Get-PnPContentType -List $currentLibrary.Id
        for ($contentTypeIndex = 0; $contentTypeIndex -lt $libraryContentTypes.Count; $contentTypeIndex++) {
            $currentContentType = $libraryContentTypes[$contentTypeIndex]
            Write-Host "`t`t$($contentTypeIndex+1)/$($libraryContentTypes.Count) - $($currentContentType.Name) - $($currentContentType.Group)"
            if ($currentContentType.Name -eq "Folder" -and $currentContentType.Group -eq "Folder Content Types" -and $libraryContentTypes.Count -ge 2) {
                Write-Host "`t`tDeleting Folder Content Type..." -ForegroundColor DarkYellow
                Remove-PnPContentTypeFromList -List $currentLibrary.Id -ContentType "Folder"
            }
        }
    }
}