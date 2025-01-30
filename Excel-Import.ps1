Clear-Host
#$Header = 'Device Name', 'Depot', 'Status', 'IMEI / MEID', 'Serial Number', 'Model Name', 'Agent Version', 'OS Version', 'MAC Address', 'Firmware Version', 'ICCID Information', 'Last Updated', 'Last Location (UTC-05:00)',	'Owner'
$data = Import-Csv -Path "C:\Users\sc13\Downloads\Master Cellphone List - csv.csv"
$siteUrl = "https://claringtonnet.sharepoint.com/sites/InfoTech"
#$siteUrl = "https://claringtonnet.sharepoint.com/sites/SCAdminTest"

Connect-PnPOnline -Url $siteUrl -Interactive


ForEach ($item in $data) {
    Write-Host ""
    Write-Host "----------------------------------------------------------------------"
    Write-Host "$($item.'Owner') - $($item.'Model') - MOC Org|Department|$($item.'Department')"

    Write-Host "Check for user: $($item.'Owner')"
    $user = Get-PnPUser | ? Title -eq $item.'Owner'

    if ($user) {
        Add-PnPListItem -List "Hardware Inventory" -ContentType "Mobile Device Item" -Values @{
            "Title"             = "$($item.'Owner') - $($item.'Model')";
            # "Status" = $item.'Status';
            # "SIMCards" = $item.'ICCID Information';
            "Model"             = $item.'Model';
            "MobilePhoneNumber" = $item.'Phone Number';
            #"MACAddress" = $item.'MAC Address';
            "IMEI"              = $item.'IMEI';
            "CurrentOwner"      = $item.'Owner';
            'AssetType'         = 'Smartphone';
            'Carrier'           = $item.'Provider';
            'Location'          = 'Mobile Device';
            "Department"        = "MOC Org|Department|$($item.'Department')";
        }
    }
    else {
        Add-PnPListItem -List "Hardware Inventory" -ContentType "Mobile Device Item" -Values @{
            "Title"             = "$($item.'Owner') - $($item.'Model')";
            # "Status" = $item.'Status';
            # "SIMCards" = $item.'ICCID Information';
            "Model"             = $item.'Model';
            "MobilePhoneNumber" = $item.'Phone Number';
            #"MACAddress" = $item.'MAC Address';
            "IMEI"              = $item.'IMEI';
            #"CurrentOwner"      = $item.'Owner';
            "Notes"             = "Last Owner: $($item.'Owner')";
            'AssetType'         = 'Smartphone';
            'Carrier'           = $item.'Provider';
            'Location'          = 'Mobile Device';
            "Department"        = "MOC Org|Department|$($item.'Department')";
        }
    }
}