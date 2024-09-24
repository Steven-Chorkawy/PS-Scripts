#Import-Module PnP.PowerShell 

<# 
    See related issue here: https://sharepoint.stackexchange.com/questions/302734/spfx-form-customizer-deployment.

    Sometimes the NewFormClientSideComponentId, EditFormClientSideComponentId, and DisplayFormClientSideComponentId properties cannot be found. 
    Apparently PnP.PowerShell version 1.11.0 works.
#>

Clear-Host
$siteURL="https://claringtonnet.sharepoint.com/sites/Clerk"
$listName = "Death Registration"
# Name of ID of Content Type.  https://pnp.github.io/powershell/cmdlets/Get-PnPContentType.html#-identity
$contentTypeName = "0x0100DA81DCD717B72D499724C0023271F50C00D36BB2C588D851418CF8CEDFA4ED7036"
# This GUID is found in the NAME.manifest.json file of the project. 
$app_manifest_id = "bc5454d2-bea9-4a5d-88f5-5e2a21dc2b98"


Connect-PnPOnline -Url $siteURL -Interactive
$clientContext = Get-PnPContext
$contentType = Get-PnPContentType -List $listName -Identity $contentTypeName

$contentType | Format-List


# This GUID is found in the NAME.manifest.json file of the project. 
$contentType.NewFormClientSideComponentId = $app_manifest_id;
#$contentType.EditFormClientSideComponentId = "";
$contentType.DisplayFormClientSideComponentId = $app_manifest_id;

$contentType.Update($false)
$clientContext.ExecuteQuery()
