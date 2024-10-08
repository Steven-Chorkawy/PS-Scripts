Clear-Host

Connect-PnPOnline -Url "https://claringtonnet.sharepoint.com/sites/DocSet_Test_3" -Interactive

$ct = Get-PnPContentType -List "Administration" -Includes Parent, Fields | Where-Object { $_.Name -like "*Case" }


Write-Host $ct.Fields[10]