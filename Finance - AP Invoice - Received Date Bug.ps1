# Clear the console
CLS

# Define the site and list URLs
$siteUrl = "https://claringtonnet.sharepoint.com/sites/Finance"
$listName = "Invoices"

# Connect to the SharePoint site using the interactive login method
Connect-PnPOnline -Url $siteUrl -Interactive

# Set the batch size (number of items to retrieve per batch)
$batchSize = 200

# Get the total number of items in the list
Write-Host "1..."
$totalItemCount = Get-PnPProperty -ClientObject (Get-PnPList -Identity $listName -Includes ItemCount) -Property ItemCount

# Initialize the page number for the first batch
$page = 1

# Initialize an empty array to hold the items
$items = @()

Write-Host "Starting loop..."

# Loop through the pages of items
do {
    # Get the items in the current page
    # !! This is currently returning 0 items.
    $pageItems = Get-PnPListItem -List $listName -Fields "Received_x0020_Date", "Modified" -PageSize $batchSize |
        Where-Object { $_["Received_x0020_Date"] -eq $_["Modified"] }

    # Add the page items to the $items array
    $items += $pageItems

    # Display the current page number in the console
    Write-Host "Page $page of $($totalItemCount / $batchSize) retrieved.  PageItems Count $($pageItems.Count).  Batch Size $($batchSize)"

    # Increment the page number for the next page
    $page++

    # Sleep for a few seconds to avoid hitting SharePoint throttling limits
    Start-Sleep -Seconds 5
} while ($pageItems.Count -eq $batchSize)

Write-Host "End of Loop..."

# Define the output file path
$outputFilePath = "C:\Users\sc13\OneDrive - clarington.net\Desktop\Excel Export\APInvoiceFix.csv"

# Export the results to a CSV file
$items | Select-Object Id, Title, @{n='Link';e={'{0}/Lists/{1}/DispForm.aspx?ID={2}' -f $siteUrl, $listName, $_.Id}}, Received_x0020_Date, Modified | Export-Csv -Path $outputFilePath -NoTypeInformation
