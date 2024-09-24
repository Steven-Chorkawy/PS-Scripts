# https://www.sharepointdiary.com/2016/05/sharepoint-online-powershell-to-create-list-view.html

# Clear host
Clear-Host

# Load SharePoint PnP module
Import-Module SharePointPnPPowerShellOnline -ErrorAction SilentlyContinue

# Connect to SharePoint tenant and retrieve list of sites
Connect-PnPOnline -Url "https://claringtonnet-admin.sharepoint.com" -Interactive
$sites = Get-PnPTenantSite

# Prompt user to select a site URL from the list
$selectedSite = $sites | Out-GridView -Title "Select site URL" -OutputMode Single

# Connect to selected site
Connect-PnPOnline -Url $selectedSite.Url -Interactive

# Get a list of all libraries and lists on the sharepoint site.
$allLists = Get-PnPList | Where-Object { $_.Hidden -eq $false }

$selectedLists = $allLists | Out-GridView -Title "Select List/ Library" -OutputMode Multiple

foreach ($list in $selectedLists) {



    #Config Variables

    # Prompt user for view name.
    $ViewName= Read-Host "Enter View Name"

    # Url of the selected site.
    $SiteURL = $selectedSite.Url

    # Title of the list of the new view.
    $ListName= $list.Title

    # Add Doc Icon and Doc Title by default.
    $ViewFields = @("DocIcon", "FileLeafRef")

    #Get list of all fields. 
    $ListFields = Get-PnPField -List $ListName

    # Prompt the user to select additional fields.
    $SelectedFields = $ListFields | Out-GridView -Title "Select Fields.  Doc Icon and File Name will automatically be added." -OutputMode Multiple

    # Prompt user to select a single field to group by. 
    # 01/12/2024 - As of now I cannot group by two columns without crashing a SharePoint site.
    $SelectedGroupByField = $ListFields | Out-GridView -Title "Select a field to Group By." -OutputMode Single

    # Add each selected field to the array of fields the view will display.
    foreach($field in $SelectedFields) {
        $ViewFields += $field.Title
    }

    $Query = "<Query><GroupBy Collapse='TRUE'><FieldRef Name='$($SelectedGroupByField)' /></GroupBy></Query>"
    
    Try {
 
        #sharepoint online pnp powershell create view
        Add-PnPView -List $ListName -Title $ViewName -Fields $ViewFields -Query $Query -ErrorAction Stop
        #Add-PnPView -List $ListName -Title $ViewName -ViewType Html -Fields $ViewFields -ErrorAction Stop
        Write-host "View '$ViewName' Created Successfully!" -f Green
    }
    catch {
        write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
    }
}



