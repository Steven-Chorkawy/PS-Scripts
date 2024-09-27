Clear-Host
$excelRows = Import-Excel -Path "C:\Users\sc13\OneDrive - clarington.net\Desktop\SharePoint Site Config Template.xlsx"

foreach ($row in $excelRows) {
    Write-Host "Connecting to $($row.SiteUrl)"
    Connect-PnPOnline -Url $row.SiteUrl -Interactive
    $docSet_ParentContentType = Get-PnPCompatibleHubContentTypes -WebUrl $row.SiteUrl | Where-Object { $_.Group -eq "Organizational Content Types" -and $_.ParentName -eq "EDRM Document Set" -and $_.Name -match $row.DocumentSet_ParentContentType }
    $doc_ParentContentType = Get-PnPCompatibleHubContentTypes -WebUrl $row.SiteUrl | Where-Object { $_.Group -eq "Organizational Content Types" -and $_.ParentName -eq "EDRM Document" -and $_.Name -match $row.DocumentSet_ParentContentType }

    Add-PnPContentTypesFromContentTypeHub -ContentTypes @($docSet_ParentContentType.Id, $doc_ParentContentType.Id)
}

Write-Host "Site Content Types"
Get-PnPContentType | Where-Object { $_.Group -eq "Organizational Content Types" }

#######################
#   This stuff works :) 
# Clear-Host

# $siteURL = "https://claringtonnet.sharepoint.com/sites/TemplateforCommitteeSites"

# Connect-PnPOnline -Url $siteURL -Interactive

# # Get content types present in content type hub site that are possible to be added to the current site.
# $hubContentTypes = Get-PnPCompatibleHubContentTypes -WebUrl $siteURL | Where-Object { $_.Group -eq "Organizational Content Types" -and $_.Name -match "AA.05.01" }

# $hubContentTypeIDs = @()

# foreach ($contentType in $hubContentTypes) {
#     $hubContentTypeIDs += $contentType.Id
# }

# Add-PnPContentTypesFromContentTypeHub -ContentTypes $hubContentTypeIDs

# Get Content Types currently on site. 
#Get-PnPContentType