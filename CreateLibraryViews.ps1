Function Update-AllDocumentsViewColumns {
    param(
        [string]$libraryTitle
    )
    $selectedColumns = Get-PnPField -List $library.Title | Where-Object { $_.Hidden -eq $false -and $_.CanBeDeleted -eq $true } 

    # Order by DocIcon as per Brandi. 
    $BrandiViewQuery = "<OrderBy><FieldRef Name='DocIcon' Ascending='TRUE'/><FieldRef Name='LinkFilename' Ascending='TRUE'/></OrderBy>"
    #Get the Client Context
    $Context = Get-PnPContext
    $allDocumentsView = Get-PnPView -List $libraryTitle -Identity "All Documents" -Includes ViewType, ViewFields, Aggregations, Paged, ViewQuery, RowLimit
    $allDocumentsView.ViewQuery = $BrandiViewQuery
    $allDocumentsView.Update()
    $Context.ExecuteQuery()
    $fieldArray = @()

    foreach ($column in $allDocumentsView.ViewFields) {
        $fieldArray += $column
    }

    foreach ($column in $selectedColumns) {
        if ($fieldArray -notcontains $column.InternalName) {
            $fieldArray += $column.InternalName
        }
    }

    Set-PnPView -List $libraryTitle -Identity "All Documents" -Fields $fieldArray    
}

Function Create-CustomChoiceViews {
    param(
        [string]$libraryTitle
    )

    Write-Host "Getting Choice fields for $($libraryTitle)"
    $fields = Get-PnPField -List $library.Title | Where-Object { $_.TypeAsString -eq "Choice" }
    foreach ($field in $fields) {
        Create-GroupByOneColumnView -libraryTitle $libraryTitle -fieldName $field.Title
    }
}

Function Create-GroupByOneColumnView {
    param(
        [string]$libraryTitle,
        [string]$fieldName
    )

    $newViewName = "Group by $($fieldName)"
    Write-Host "Attempting to create a $($newViewName) for $($libraryTitle)"
    # If topicField is null we cannot create the view.
    $newField = Get-PnPField -Identity $fieldName -List $libraryTitle -ErrorAction SilentlyContinue
    # If newView is NOT null the view already exists and we do not need to create another one.
    $newView = Get-PnPView -Identity $newViewName -List $libraryTitle -ErrorAction SilentlyContinue
    $allDocumentsView = Get-PnPView -List $libraryTitle -Identity "All Documents" -Includes ViewType, ViewFields, Aggregations, Paged, ViewQuery, RowLimit

    If ($newView) {       
        Write-Host "$($newViewName) already exists in $($libraryTitle)... Skipping this step." -ForegroundColor Yellow
    }
    else {
        # View does not exist.  Create proceed to create the view.
        If ($newField) {
            Write-Host "$($fieldName) Field Found..." -ForegroundColor Green
            # Update the view properties of the All Documents view.
            $allDocumentsView.ViewFields.add($fieldName)
            $allDocumentsView.ViewQuery = "$($allDocumentsView.ViewQuery)<GroupBy Collapse='TRUE'  GroupLimit='30'><FieldRef Name='$($newField.InternalName)' Ascending='TRUE' /></GroupBy>"

            #Get Properties of the source View
            $ViewProperties = @{
                "List"         = $libraryTitle
                "Title"        = $newViewName
                "Paged"        = $allDocumentsView.Paged
                "Personal"     = $allDocumentsView.PersonalView
                "Query"        = $allDocumentsView.ViewQuery
                "RowLimit"     = $allDocumentsView.RowLimit
                "SetAsDefault" = $false
                "Fields"       = @($allDocumentsView.ViewFields)
                "ViewType"     = $allDocumentsView.ViewType
                "Aggregations" = $allDocumentsView.Aggregations
            }

            Add-PnPView @ViewProperties
            Write-Host "$($newViewName) View has been created!" -ForegroundColor Green
        }
        else {
            Write-Host "ERROR! $($fieldName) field not found in $($libraryTitle) library!" -ForegroundColor Red
            # All checks are good.  Create the view.
        }
    }
}

<#
.SYNOPSIS
Create a view that groups by two columns.

.DESCRIPTION
Create a view that groups by two columns.

.PARAMETER libraryTitle
Display name of a SharePoint library.

.PARAMETER fieldOneName
Display name of the column used to group by first.

.PARAMETER fieldTwoName
Display name of the column used to group by second.

.EXAMPLE
Create-GroupByTwoColumnView -libraryTitle $library.Title -fieldOneName "Document Type" -fieldTwoName "Topic"

.NOTES
General notes
#>
Function Create-GroupByTwoColumnView {
    param(
        [string]$libraryTitle,
        [string]$fieldOneName,
        [string]$fieldTwoName
    )

    $newViewName = "Group by $($fieldOneName) & $($fieldTwoName)"
    Write-Host "Attempting to create a $($newViewName) for $($libraryTitle)"
    # If topicField is null we cannot create the view.
    $firstField = Get-PnPField -Identity $fieldOneName -List $libraryTitle -ErrorAction SilentlyContinue
    $secondField = Get-PnPField -Identity $fieldTwoName -List $libraryTitle -ErrorAction SilentlyContinue

    # If newView is NOT null the view already exists and we do not need to create another one.
    $newView = Get-PnPView -Identity $newViewName -List $libraryTitle -ErrorAction SilentlyContinue
    $allDocumentsView = Get-PnPView -List $libraryTitle -Identity "All Documents" -Includes ViewType, ViewFields, Aggregations, Paged, ViewQuery, RowLimit

    If ($newView) {       
        Write-Host "$($newViewName) already exists in $($libraryTitle)... Skipping this step." -ForegroundColor Yellow
    }
    else {
        # View does not exist.  Create proceed to create the view.
        If ($firstField) {
            Write-Host "$($fieldOneName) Field Found..." -ForegroundColor Green
            If ($secondField) {
                Write-Host "$($fieldTwoName) Field Found..." -ForegroundColor Green
                # Update the view properties of the All Documents view.
                $allDocumentsView.ViewQuery = "$($allDocumentsView.ViewQuery)<GroupBy Collapse='TRUE'  GroupLimit='30'><FieldRef Name='$($firstField.InternalName)' /> <FieldRef Name='$($secondField.InternalName)' /></GroupBy>"

                #Get Properties of the source View
                $ViewProperties = @{
                    "List"         = $libraryTitle
                    "Title"        = $newViewName
                    "Paged"        = $allDocumentsView.Paged
                    "Personal"     = $allDocumentsView.PersonalView
                    "Query"        = $allDocumentsView.ViewQuery
                    "RowLimit"     = $allDocumentsView.RowLimit
                    "SetAsDefault" = $false
                    "Fields"       = @($allDocumentsView.ViewFields)
                    "ViewType"     = $allDocumentsView.ViewType
                    "Aggregations" = $allDocumentsView.Aggregations
                }

                Add-PnPView @ViewProperties
                Write-Host "$($newViewName) View has been created!" -ForegroundColor Green
            }
            else {
                Write-Host "ERROR! $($fieldTwoName) field not found in $($libraryTitle) library!" -ForegroundColor Red
            }
        }
        else {
            Write-Host "ERROR! $($fieldOneName) field not found in $($libraryTitle) library!" -ForegroundColor Red
        }
    }
}

Clear-Host

# Connect to SharePoint tenant and retrieve list of sites
Connect-PnPOnline -Url "https://claringtonnet-admin.sharepoint.com" -Interactive
$sites = Get-PnPTenantSite

# Prompt user to select a site URL from the list
$selectedSite = $sites | Out-GridView -Title "Select site URL" -OutputMode Single

# Connect to selected site
Connect-PnPOnline -Url $selectedSite.Url -Interactive
# $context = Get-PnPContext

#Get all document libraries - Exclude Hidden Libraries
$DocumentLibraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false -and $_.Title -ne "Form Templates" -and $_.Title -ne "Site Assets" -and $_.Title -ne "Style Library" } #Or $_.BaseType -eq "DocumentLibrary"

$selectedLibraries = $DocumentLibraries | Out-GridView -Title "Select Libraries" -OutputMode Multiple

# Get a list of all metadata columns from library.
foreach ($library in $selectedLibraries) {
    Write-Host "`n===============================================================" -ForegroundColor Cyan
    Write-Host "Library -> $($library.Title)"

    Update-AllDocumentsViewColumns -libraryTitle $library.Title

    Create-CustomChoiceViews -libraryTitle $library.Title

    # We want to create a group by view for each of these columns.
    # Filtering for "$_.Hidden -eq $false -and $_.CanBeDeleted -eq $true" seems to return custom fields only.  
    #This query might need to be updated in the future.
    $customColumns = Get-PnPField -List $library.Title | Where-Object { $_.Hidden -eq $false -and $_.CanBeDeleted -eq $true } 
    foreach ($column in $customColumns) {
        Create-GroupByOneColumnView -libraryTitle $library.Title -fieldName $column.Title
    }

    # Always try to create these 2x group by views.
    Create-GroupByTwoColumnView -libraryTitle $library.Title -fieldOneName "Topic" -fieldTwoName "Year"
    Create-GroupByTwoColumnView -libraryTitle $library.Title -fieldOneName "Year" -fieldTwoName "Topic"
    Create-GroupByTwoColumnView -libraryTitle $library.Title -fieldOneName "Document Type" -fieldTwoName "Topic"
    Create-GroupByTwoColumnView -libraryTitle $library.Title -fieldOneName "Topic" -fieldTwoName "Document Type"
}