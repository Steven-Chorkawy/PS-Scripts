Clear-Host

# Set the SharePoint site URL and the name of the output CSV file
$siteUrl = "https://claringtonnet.sharepoint.com/sites/Finance"
$outputFile = "C:\Users\sc13\OneDrive - clarington.net\Desktop\Excel Export\RecycleBinItems.csv"

# Connect to the SharePoint site
Connect-PnPOnline -Url $siteUrl -Interactive

# Get all items in the site's recycle bin
$recycleBinItems = Get-PnPRecycleBinItem

# Export the recycle bin items to a CSV file
$recycleBinItems | Export-Csv -Path $outputFile -NoTypeInformation

# Disconnect from the SharePoint site
Disconnect-PnPOnline