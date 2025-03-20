<#
Ticket: https://clarington.freshservice.com/a/tickets/46072?current_tab=details

The purpose of this script is to crawl all of our SharePoint sites and generate a report of all the content types that are being used.
#>

# ! Start of Script.
Clear-Host

$FileDateTime = Get-Date -Format FileDateTime
$DEFAULT_EXPORT_FOLDER_PATH = "C:\Users\sc13\OneDrive - clarington.net\Desktop\"
$DEFAULT_EXPORT_FOLDER_NAME = "ContentTypeAuditExport-$($FileDateTime)"
$DEFAULT_FILE_NAME = "SharePointContentTypeAudit-$($FileDateTime).csv"
$DEFAULT_EXPORT_FOLDER = "$($DEFAULT_EXPORT_FOLDER_PATH)$($DEFAULT_EXPORT_FOLDER_NAME)"
$EXPORT_FOLDER = Read-Host "Enter Path to create $($DEFAULT_EXPORT_FOLDER_NAME) or press enter to use default path"

if ($EXPORT_FOLDER -eq "") {
    Write-Host "No EXPORT_FOLDER path provided... Using Default Path $($DEFAULT_EXPORT_FOLDER)"
    $EXPORT_FOLDER = $DEFAULT_EXPORT_FOLDER
}

# Make a new folder for export results.
New-Item -Path $DEFAULT_EXPORT_FOLDER_PATH -Name $DEFAULT_EXPORT_FOLDER_NAME -ItemType "directory"

# Do not include these sites in the report.
$EXCLUDED_SITE_URLS = @(
    "https://claringtonnet-my.sharepoint.com/", 
    "https://claringtonnet.sharepoint.com/portals/hub",
    "https://claringtonnet.sharepoint.com/portals/Community",
    "https://claringtonnet.sharepoint.com/search",
    "https://claringtonnet.sharepoint.com/sites/ClaringtonAppCatalog",
    "https://claringtonnet.sharepoint.com/portals/SharePoint-Invoice-Demo"
)

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

    #Get all document libraries - Exclude Hidden Libraries
    $documentLibraries = Get-PnPList -Includes DefaultView | Where-Object { $_.Hidden -eq $false -and $_.Title -ne "Form Templates" -and $_.Title -ne "Site Assets" -and $_.Title -ne "Style Library" -and $_.Title -ne "Site Pages" -and $_.EntityTypeName -ne "Shared_x0020_Documents" -and $_.EntityTypeName -ne "EventsList" }

    for ($libraryIndex = 0; $libraryIndex -lt $documentLibraries.Count; $libraryIndex++) {
        $currentLibrary = $documentLibraries[$libraryIndex]
        $currentLibraryDefaultViewUrl = "$($currentSite.Url)$($currentLibrary.DefaultViewUrl.TrimStart("/"))"
        Write-Host "`t$($libraryIndex+1)/$($documentLibraries.Count) - $($currentLibrary.Title) - $($currentLibraryDefaultViewUrl) - $($currentLibrary.ItemCount) Items"

        $libraryContentTypes = Get-PnPContentType -List $currentLibrary.Id -Includes @("Parent", "Parent.Parent")
        for ($contentTypeIndex = 0; $contentTypeIndex -lt $libraryContentTypes.Count; $contentTypeIndex++) {
            $currentContentType = $libraryContentTypes[$contentTypeIndex]
            Write-Host "`t`t$($contentTypeIndex+1)/$($libraryContentTypes.Count) - $($currentContentType.Name) - $($currentContentType.Group)"
            Write-Host "`t`tParent CT: $($currentContentType.Parent.Name) - $($currentContentType.Parent.Group)"

            # This should be the Organizational Content Type, Document, Document Set, Folder, or List item.
            Write-Host "`t`tParents Parent CT: $($currentContentType.Parent.Parent.Name) - $($currentContentType.Parent.Parent.Group)"
            Write-Host "`n"

            $EXPORT_OBJECT = @{
                "Site Title"                = $currentSite.Title;
                "Site URL"                  = $currentSite.Url;
                "Site Status"               = $currentSite.Status
                "Library Title"             = $currentLibrary.Title;
                "Library URL"               = $currentLibraryDefaultViewUrl;
                "Library Item Count"        = $currentLibrary.ItemCount;
                "Content Type"              = $currentContentType.Name;
                "Content Type Group"        = $currentContentType.Group;
                "Parent Content Type"       = $currentContentType.Parent.Parent.Name;
                "Parent Content Type Group" = $currentContentType.Parent.Parent.Group;
            }

            $EXPORT_OBJECT | Select-Object "Site Title", "Site URL", "Site Status", "Library Title", "Library URL", "Library Item Count", "Content Type", "Content Type Group", "Parent Content Type", "Parent Content Type Group" | Export-Csv -Path "$($DEFAULT_EXPORT_FOLDER)\$($DEFAULT_FILE_NAME)" -Append -NoTypeInformation
        }
    }
}
# ! End of Script.