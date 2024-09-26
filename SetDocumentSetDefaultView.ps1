<#
#
#   START OF SCRIPT.
#
#>
Clear-Host

# Connect to SharePoint tenant and retrieve list of sites
Connect-PnPOnline -Url "https://claringtonnet-admin.sharepoint.com" -Interactive
$sites = Get-PnPTenantSite

# Prompt user to select a site URL from the list
$selectedSite = $sites | Out-GridView -Title "Select site URL" -OutputMode Single

# Connect to selected site
Connect-PnPOnline -Url $selectedSite.Url -Interactive

$Context = Get-PnPContext
$DOCUMENT_SET_ID = "0x0120D520"
$contentTypeIndexCounter = 0

#Get all document libraries - Exclude Hidden Libraries
$DocumentLibraries = Get-PnPList -Includes Views, ContentTypes, Title | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false -and $_.Title -ne "Form Templates" -and $_.Title -ne "Site Assets" -and $_.Title -ne "Style Library" } #Or $_.BaseType -eq "DocumentLibrary"

$selectedLibraries = $DocumentLibraries | Out-GridView -Title "Select Libraries" -OutputMode Multiple

foreach ($library in $selectedLibraries) {
    $allViews = Get-PnPView -List $library.Title

    foreach ($contentType in $library.ContentTypes) {
        $Context.Load($contentType.Parent)
        $Context.ExecuteQuery()
        $parentContentType = $contentType.Parent
        if ($parentContentType.Id.ToString().StartsWith($DOCUMENT_SET_ID)) {
            # Write-Host "$($parentContentType.Name) -> $($parentContentType.Id)"
            $selectedView = $allViews | Out-GridView -Title "Select default view for $($library.Title) -> $($parentContentType.Name)" -OutputMode Single
            # Write-Host "Updating $($selectedView.Title)"
            # Set-PnPView -List $library.Title -Identity $selectedView.Id -Values @{ContentTypeId = $contentType.Id; DefaultViewForContentType = $true }

            $innerXML = "<wpv:WelcomePageView xmlns:wpv=`"http://schemas.microsoft.com/office/documentsets/welcomepageview`" ViewId=`"$($selectedView.Id)`" />"

            # $library.ContentTypes[$contentTypeIndexCounter].SchemaXml = $innerXML
            # $library.Update()
            # $Context.Load($library)
            # $Context.ExecuteQuery()

            Set-PnPContentType -List $library.Title -Identity $contentType.Id -Values @{SchemaXml = $innerXML }
        }
        $contentTypeIndexCounter++
    }
}

Write-Host "The End."
<#
#
#   END OF SCRIPT.
#
#>