Clear-Host
$PnPID = "08a2a4b7-9f6f-46cb-87e7-2ed02a66fc22"
$siteUrl = Read-Host "Enter URL of the site to save as a template (e.g., https://Claringtonnet.sharepoint.com/sites/yoursite)"
$Path = Read-Host "Enter the path where you'd like to save the site template (e.g., C:\Temp\MoCCommittee.pnp)"

Connect-PnPOnline -url $siteUrl -ClientId $PnPID -Interactive
Get-PnPSiteTemplate -out $Path -IncludeAllTermGroups 
#-ExcludeHandlers SiteSecurity