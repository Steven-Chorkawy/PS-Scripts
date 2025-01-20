#Connect to PnP Online
Connect-PnPOnline -Url "https://claringtonnet-admin.sharepoint.com" -Interactive
   
$allSites = Get-PnPTenantSite | Where-Object { $_.Status -eq "Active" -and ($_.Template -eq "SITEPAGEPUBLISHING#0" -or $_.Template -eq "STS#3") } | Select-Object -Property Title, Url, Status, Owner, Template, IsTeamsConnected, IsTeamsChannelConnected

$allSites | Export-Csv -Path "C:\Users\sc13\OneDrive - clarington.net\Desktop\Excel Export\Disable_Members_Edit_Membership.csv"

#Enable Members can Sharing
$Web = Get-PnPWeb
$Web.MembersCanShare = $true
$Web.Update()
Invoke-PnPQuery
 
#Update the Members Group - Members can share Files and Folders. But ONLY Site owners can share site!
$MembersGroup = Get-PnPGroup -AssociatedMemberGroup
Set-PnPGroup -Identity $MembersGroup -AllowMembersEditMembership:$false

foreach ($site in $allSites) {
    Write-Host $site.Title
    Connect-PnPOnline -Url $site.Url -Interactive
    $Web = Get-PnPWeb
    Write-Host $Web.Url
}

Write-Host "The End."