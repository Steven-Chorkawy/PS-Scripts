<#
The purpose of this script is to apply common metadata columns to a list or library to track and record Teams Approvals.
#>

# ! Start of CONST.
$COLUMNS_NAMES = @("Approval Status", "Approvers", "Approved By", "Approval Comments", "Approval Summary", "Approval Date" )
# ! End of CONST.

# ! Start of Functions.
Function CreateColumns {
    param(
        [String]$ColumnName
    )

    Write-Host "`tCreating new $($ColumnName) column." -ForegroundColor Magenta

    switch ($ColumnName) {
        "Approval Comments" { 
            Add-PnPField -DisplayName "Approval Comments" -InternalName "ApprovalComments" -Type Note -Group "Custom Columns" | Out-Null
        }
        "Approval Date" {
            $approvalDateFieldXML = "<Field Type='DateTime' Name='ApprovalDate' ID='$([GUID]::NewGuid())' DisplayName='Approval Date' Required ='FALSE' Format='DateOnly' FriendlyDisplayFormat='Disabled'></Field>"
            Add-PnPFieldFromXml -FieldXml $approvalDateFieldXML | Out-Null
        }
        "Approval Status" {
            Add-PnPField -DisplayName "Approval Status" -InternalName "ApprovalStatus" -Type Choice -Group "Custom Columns" -Choices "New", "Awaiting Approval", "Approved", "Denied", "Cancelled" | Out-Null
        }
        "Approval Summary" {
            Add-PnPField -DisplayName "Approval Summary" -InternalName "ApprovalSummary" -Type Note -Group "Custom Columns" | Out-Null
        }
        "Approved By" {
            $approvedByFieldXML = "<Field Type='User' Name='ApprovedBy' ID='$([GUID]::NewGuid())' DisplayName='Approved By' Required ='FALSE' UserSelectionMode='PeopleOnly' Mult='TRUE'></Field>"
            Add-PnPFieldFromXml -FieldXml $approvedByFieldXML | Out-Null
        }
        "Approvers" {
            $approversFieldXML = "<Field Type='User' Name='Approvers' ID='$([GUID]::NewGuid())' DisplayName='Approvers' Required ='FALSE' UserSelectionMode='PeopleOnly' Mult='TRUE'></Field>"
            Add-PnPFieldFromXml -FieldXml $approversFieldXML | Out-Null
        }
        Default {}
    }

}

Function CheckForSiteColumns {
    foreach ($columnName in $COLUMNS_NAMES) {
        try {
            $column = Get-PnPField -Identity $columnName | Where-Object { $_.Hidden -eq $false }
            if ($column) {
                Write-Host "$($columnName) already exists." -ForegroundColor Green
            }
            else {
                CreateColumns -ColumnName $columnName
            }
        }
        catch {
            CreateColumns -ColumnName $columnName
        }
    }
}
# ! End of Functions.

# ! Start of Script.
Clear-Host
# Connect to SharePoint tenant and retrieve list of sites
Connect-PnPOnline -Url "https://claringtonnet-admin.sharepoint.com" -Interactive
$tenantSites = Get-PnPTenantSite

$selectedSite = $tenantSites | Out-GridView -Title "Select site URL" -OutputMode Single

# Connect to selected site
Connect-PnPOnline -Url $selectedSite.Url -Interactive

$Context = Get-PnPContext

#Get all document libraries - Exclude Hidden Libraries
$documentLibraries = Get-PnPList -Includes DefaultView | Where-Object { $_.Hidden -eq $false -and $_.Title -ne "Form Templates" -and $_.Title -ne "Site Assets" -and $_.Title -ne "Style Library" } #Or $_.BaseType -eq "DocumentLibrary"

$selectedLibraries = $DocumentLibraries | Out-GridView -Title "Select Libraries" -OutputMode Multiple

# Check if Site columns exist.
CheckForSiteColumns

foreach ($library in $selectedLibraries) {
    Write-Host "`n===============================================================" -ForegroundColor Cyan
    Write-Host "$($selectedSite.Title) -> $($library.Title)"

    $listContentTypes = Get-PnPContentType -List $library.Title
    $listViews = Get-PnPView -List $library.Title

    foreach ($column in $COLUMNS_NAMES) {
        Write-Host "`tAdding $($column) to all Content Types."
        Add-PnPField -List $library.Title -Field $column
        foreach ($contentType in $listContentTypes) {   
            Add-PnPFieldToContentType -Field $column -ContentType $contentType -UpdateChildren $false
        }

        Write-Host "`tAdding $($column) to all views..."
        foreach ($view in $listViews) {
            if ($view.ViewFields -notcontains $column) {
                $view.ViewFields.Add($column)
                $view.Update()
                $Context.ExecuteQuery()
            }
        }
    }
}

Write-Host "`n`nThe End..."
# ! End of Script.