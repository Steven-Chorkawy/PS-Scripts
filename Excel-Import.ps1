Clear-Host
$Header = 'Device Name','Depot', 'Status', 'IMEI / MEID','Serial Number', 'Model Name', 'Agent Version', 'OS Version', 'MAC Address', 'Firmware Version', 'ICCID Information', 'Last Updated', 'Last Location (UTC-05:00)',	'Owner'
$data = Import-Csv -Path "C:\Users\sc13\Downloads\CSV_FOCUS Devices.csv"

Connect-PnPOnline -Url "https://claringtonnet.sharepoint.com/sites/InfoTech" -Interactive


ForEach($item in $data) {
    Write-Host $item.'Device Name'

    Add-PnPListItem -List "Hardware Inventory" -ContentType "Mobile Device Item" -Values @{
        "Title" = $item.'Device Name';
        "Status" = $item.'Status';
        "SIMCards" = $item.'ICCID Information';
        "Model" = $item.'Model Name';
        "MACAddress" = $item.'MAC Address';
        "IMEI" = $item.'IMEI / MEID';
        "CurrentOwner" = $item.'Owner'
    }
}