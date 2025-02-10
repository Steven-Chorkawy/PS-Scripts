# Update SharePoint items from a CSV file.

Clear-Host
$data = Import-Csv -Path "C:\Users\sc13\Downloads\Master Cellphone List - csv.csv"
$siteUrl = "https://claringtonnet.sharepoint.com/sites/InfoTech"
$listName = "Hardware Inventory"

Connect-PnPOnline -Url $siteUrl -Interactive

ForEach ($item in $data) {
    Write-Host ""
    Write-Host "----------------------------------------------------------------------"
    Write-Host "$($item.'Owner') - $($item.'Model') - $($item.'IMEI')"
    Write-Host "`tGetting List Item ID."
    $spListItem = Get-PnPListItem -List $listName -Query "<View><Query><Where><Eq><FieldRef Name='IMEI'/><Value Type='Text'>$($item.'IMEI')</Value></Eq></Where></Query></View>" | Select-Object Id

    if ([string]$item.'Contract Start' -as [DateTime]) {
        Write-Host "`tUpdating List Item."
        Set-PnPListItem -List $listName -Identity $spListItem.Id -Values @{"PurchaseDate" = [DateTime]$item.'Contract Start' }
    }
    else {
        Write-Host "`tInvalid Date!  Skipping row. $($item.'Contract Start')"
    }
}
