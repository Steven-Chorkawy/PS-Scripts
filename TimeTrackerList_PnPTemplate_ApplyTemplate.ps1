Clear-Host
$PnPID = "08a2a4b7-9f6f-46cb-87e7-2ed02a66fc22"
$siteUrl = Read-Host "Enter URL of the site to deploy the template (e.g., https://Claringtonnet.sharepoint.com/sites/yoursite)"
$Path = Read-Host "Enter the path to your site template (e.g., C:\Temp\MoCCommittee.pnp)"

# Display name of the Timesheet list content type.
$GlobalContentTypeName = "Global Timesheet Item"
$GlobalContentTypeID = "0x0100AFFE98C4E58991409E853546CA3A0172"

# Content type name at the site level.
$LocalContentTypeName = "Timesheet Item"

# Display name of the main Timesheet list.
$TimesheetList = "Timesheet" 

$FieldOrder = @("Task", "Employee", "WorkDate", "HoursWorked", "OT", "MM_Department", "Project", "Details")


Connect-PnPOnline -url $siteUrl -ClientId $PnPID -Interactive

# Disable "NoScript" on the site before applying Invoke-PnPSiteTemplate.
# https://github.com/pnp/powershell/discussions/4014#discussioncomment-9774445
# https://clarington.freshservice.com/a/tickets/44698?current_tab=details&focus_conversation=8089164582
Set-PnPTenantSite -Url $siteUrl -DenyAddAndCustomizePages:$false

Add-PnPContentTypesFromContentTypeHub -ContentTypes $GlobalContentTypeID

Invoke-PnPSiteTemplate -Path $Path -ClearNavigation

# Set the Site content type to editable.
Set-PnPContentType -Identity $GlobalContentTypeID -ReadOnly $false

# Set the list content type to editable.
Set-PnPContentType -Identity $LocalContentTypeName -List $TimesheetList -ReadOnly $false

# Add the WorkDate field to the content type.  This field will not copy from the CTHub.
$workDateField = Get-PnPField -List $TimesheetList -Identity "WorkDate"
Add-PnPFieldToContentType -ContentType $LocalContentTypeName -Field $workDateField

# TODO: Add this field to a site level content type.  Not the CT Hub content type.
# Query and add the lookup field to the list content type.  This field will not copy from the CTHub.
$projectField = Get-PnPField -List $TimesheetList -Identity "Project"
Add-PnPFieldToContentType -ContentType $LocalContentTypeName -Field $projectField

# TODO: Add this field to a site level content type.  Not the CT Hub content type.
# Query and add the lookup field to the list content type.  This field will not copy from the CTHub.
$taskField = Get-PnPField -List $TimesheetList -Identity "Task"
Add-PnPFieldToContentType -ContentType $LocalContentTypeName -Field $taskField

# Remove the default Item content type from the list.
Remove-PnPContentTypeFromList -List $TimesheetList -ContentType "Item"

Set-PnPField -List $TimesheetList -Identity "Title" -Values @{Required = $false; }

#Get the default content type from list
$ContentType = Get-PnPContentType -Identity $LocalContentTypeName -List $TimesheetList
  
#Update column Order in default content type
$FieldLinks = Get-PnPProperty -ClientObject $ContentType[0] -Property "FieldLinks"
$FieldLinks.Reorder($FieldOrder)
$ContentType.Update($False)

$titleField = Get-PnPField -Identity "Title" -List $TimesheetList
$titleField.Hidden = $true
$titleField.SetShowInDisplayForm($false)
$titleField.SetShowInEditForm($false)
$titleField.SetShowInNewForm($false)
$titleField.Update()

Invoke-PnPQuery