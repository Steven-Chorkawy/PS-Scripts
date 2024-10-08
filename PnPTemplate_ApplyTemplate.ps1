Clear-Host
$PnPID = "08a2a4b7-9f6f-46cb-87e7-2ed02a66fc22"
$siteUrl = Read-Host "Enter URL of the site to deploy the template (e.g., https://Claringtonnet.sharepoint.com/sites/yoursite)"
$Path = Read-Host "Enter the path to your site template (e.g., C:\Temp\MoCCommittee.pnp)"

Connect-PnPOnline -url $siteUrl -ClientId $PnPID -Interactive
Invoke-PnPSiteTemplate -Path $Path -ClearNavigation

# remove document content type
$Libraries = Get-PnPList | Where-Object {
    $_.Hidden -eq $false -and
    $_.Title -notlike "*Site Assets*" -and
    $_.Title -notlike "*Site Pages*" -and   
    $_.Title -ne "Documents" -and                    
    $_.Title -notlike "*Style Library*" -and
    $_.Title -notlike "*Form Templates*"
}

# Iterate through the document libraries and add them to the $Results array
foreach ($Library in $Libraries) {
    Write-Host "`nUpdating '$($Library.Title)'..."

    Remove-PnPContentTypeFromList -List $Library -ContentType "Document" -ErrorAction SilentlyContinue

    # Delete this field because Document Sets will have their own description field.
    Remove-PnPField -List $Library -Identity "_ExtendedDescription" -Force
    
    # Force Document Sets to use the modern form.
    $docSetCT = Get-PnPContentType -List $Library | Where-Object { $_.Name -like "*Case" }
    # Nulling this out makes Document Sets use the modern form.
    # Source: https://pnp.github.io/script-samples/spo-document-sets-modern-new-form/README.html?tabs=pnpps
    $docSetCT.NewFormClientSideComponentId = $null
    $docSetCT.Update($false)
}

# Required to force document sets to update.
Invoke-PnPQuery