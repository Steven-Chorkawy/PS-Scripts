# Title:    GetAllLibraryContentTypes.ps1
# Author:   Steven Chorkawy
# Date:     02/23/2022
# Modified: 02/23/2022

Clear-Host

# Prompt the user to see what environment they would like to work in. 
$tenantEnvironmentUrl = @("https://claringtonnet-admin.sharepoint.com/", "https://claringtonnetdev-admin.sharepoint.com/") | Out-GridView -OutputMode Single -Title "Select Prod or Dev Tenant"

#Connect to SharePoint Online Admin Site. 
# Note that the -Interactive flag is now prefered over the -UseWebLogin flag. https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html#-interactive
Connect-PnPOnline -Url $tenantEnvironmentUrl -Interactive

# Gets a list of all SP sites.
$AllSites = Get-PnPTenantSite
Write-Host $AllSites.Count " Sites Found..."

for ($siteIndex = 0; $siteIndex -lt $AllSites.Count; $siteIndex++) 
{
    $site = $AllSites[$siteIndex]
    Write-Host "Index: " $siteIndex " | " $site.Url


    Connect-PnPOnline -Url $site.Url -Interactive

    $DocumentLibraries = Get-PnPList | Where-Object {($_.BaseTemplate -eq 101 -and $_.Hidden -eq $false -and $_.Title -ne "Documents" -and $_.Title -ne "Form Templates" -and $_.Title -ne "Site Assets" -and $_.Title -ne "Style Library")} #Or $_.BaseType -eq "DocumentLibrary"
    $DocumentLibraries | select Title, Url, ItemCount
    Write-Host $DocumentLibraries.Count " Libraries Found"
    Read-Host

    for($libraryIndex = 0; $libraryIndex -lt $DocumentLibraries.Count; $libraryIndex++)
    {
        $library = $DocumentLibraries[$libraryIndex]
        
        Write-Host ($library| Format-Table | Out-String)
        Write-Host ($library.ContentTypes.Count)
        #$ListContentTypes = Get-PnPContentType -List $library.ContentTypes
    }
}