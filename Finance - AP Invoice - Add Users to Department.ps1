Clear-Host

Connect-PnPOnline -Url "https://claringtonnet.sharepoint.com/sites/Finance" -Interactive

$BATCH_SIZE = 200

$items = Get-PnPListItem -List "Invoices" -PageSize $BATCH_SIZE | Where-Object { $_["Department"] -contains "Legislative Services" }

Write-Host "Items Found..."
Write-Host $items.Count