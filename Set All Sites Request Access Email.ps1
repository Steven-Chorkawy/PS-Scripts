Clear-Host

#Set global Parameters
$AdminSiteUrl = "https://claringtonnet-admin.sharepoint.com"
$AccessReqEmail = "helpdesk@clarington.net"
$Counter = 1

Connect-PnPOnline -Url $AdminSiteUrl -Interactive

#Get All Site collections data and export to CSV file
$AllSites = Get-PnPTenantSite -Detailed | Select Title, URL

Write-Host "Number of Sites Found: $($AllSites.Count)"

 ForEach ($Site in $AllSites)
 {
    Connect-PnPOnline -Url $Site.Url -Interactive
    
    Write-Host "Connected to Site $($Counter)/$($AllSites.Count)... $(Get-PnPConnection | Select Url)"
    
    # Add my admin account as a site owner.
    Set-PnPTenantSite -Identity $Site.Url -Owners "scadmin@clarington.net"
    $CurrentReqAccessEmail = Get-PnPRequestAccessEmails

    if($CurrentReqAccessEmail -eq $AccessReqEmail) 
    {
        Write-Host "Access Req Email Already Set." -ForegroundColor Green
    }
    else
    {
        Set-PnPRequestAccessEmails -Emails $AccessReqEmail
    }
 
    $Counter = $Counter + 1
 }

 Disconnect-PnPOnline