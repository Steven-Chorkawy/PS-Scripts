Clear-Host

# Connect to SharePoint tenant and retrieve list of sites
Connect-PnPOnline -Url "https://claringtonnet-admin.sharepoint.com" -Interactive
$sites = Get-PnPTenantSite

# Prompt user to select a site URL from the list
$selectedSite = $sites | Out-GridView -Title "Select site URL" -OutputMode Single

# Connect to selected site
Connect-PnPOnline -Url $selectedSite.Url -Interactive
Write-Host "Connect-PnPOnline -> $($selectedSite.Url)"

$allCT = Get-PnPContentType
$contentTypes = Get-PnPContentType | Where-Object { $_.Group -eq "Organizational Content Types" }
$aaCT = Get-PnPContentType | Where-Object { $_.Name -like "Document" }

$aaCT | Format-Table


if ($contentTypes.Length -eq 0) {
    Write-Host "No Organizational Content Types found... Checking for other CT."
    $allCT | Format-Table
    Write-Host "Content Types Found: $($allCT.Length)" -ForegroundColor Yellow
}
Write-Host "The End."