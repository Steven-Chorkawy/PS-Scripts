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

#Get all document libraries - Exclude Hidden Libraries
$DocumentLibraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false -and $_.Title -ne "Form Templates" -and $_.Title -ne "Site Assets" -and $_.Title -ne "Style Library" } #Or $_.BaseType -eq "DocumentLibrary"

$selectedLibraries = $DocumentLibraries | Out-GridView -Title "Select Libraries" -OutputMode Multiple

foreach ($library in $selectedLibraries) {
    $allViews = Get-PnPView -List $library.Title
    foreach ($view in $allViews) {
        if ($view.Title -ne "All Documents" -and $view.BaseViewId -eq "1") {
            Write-Host "Attempting to remove $($library.Title) > $($view.Title)."
            # This will prompt the user before deleting the view.
            Remove-PnPView -List $library.Title -Identity $view.Title -Force
            Write-Host "$($library.Title) > $($view.Title) has been removed."
        }
    }
}

<#
#
#   END OF SCRIPT.
#
#>