Function Update-AllDocumentsViewColumns {
    param(
        [string]$libraryTitle
    )
    $selectedColumns = Get-PnPField -List $library.Title | Where-Object { $_.Hidden -eq $false -and $_.CanBeDeleted -eq $true } | Out-GridView -Title "Select columns to add to $($libraryTitle) -> All Documents View" -OutputMode Multiple
    $allDocumentsView = Get-PnPView -List $libraryTitle -Identity "All Documents" -Includes ViewType, ViewFields, Aggregations, Paged, ViewQuery, RowLimit
    $fieldArray = @()

    foreach ($column in $allDocumentsView.ViewFields) {
        $fieldArray += $column    
    }

    foreach ($column in $selectedColumns) {
        $fieldArray += $column.InternalName
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

Function Create-GroupByTwoColumnView {
    param(
        [string]$libraryTitle,
        [string]$fieldOneName,
        [string]$fieldTwoName
    )

    $newViewName = "Group by $($fieldOneName) & $($fieldTwoName)"
    Write-Host "`n========================================================="
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
$DocumentLibraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false } #Or $_.BaseType -eq "DocumentLibrary"

$selectedLibraries = $DocumentLibraries | Out-GridView -Title "Select Libraries" -OutputMode Multiple

# Get a list of all metadata columns from library.
foreach ($library in $selectedLibraries) {
    Write-Host "Library -> $($library.Title)"

    Update-AllDocumentsViewColumns -libraryTitle $library.Title

    Create-CustomChoiceViews -libraryTitle $library.Title

    Create-GroupByOneColumnView -libraryTitle $library.Title -fieldName "Topic"
    Create-GroupByOneColumnView -libraryTitle $library.Title -fieldName "Year"
    Create-GroupByOneColumnView -libraryTitle $library.Title -fieldName "Status"
    Create-GroupByOneColumnView -libraryTitle $library.Title -fieldName "Document Type"
    Create-GroupByOneColumnView -libraryTitle $library.Title -fieldName "Division"
    Create-GroupByOneColumnView -libraryTitle $library.Title -fieldName "Department"
    Create-GroupByOneColumnView -libraryTitle $library.Title -fieldName "Location"

    Create-GroupByTwoColumnView -libraryTitle $library.Title -fieldOneName "Topic" -fieldTwoName "Year"
    Create-GroupByTwoColumnView -libraryTitle $library.Title -fieldOneName "Year" -fieldTwoName "Topic"
    Create-GroupByTwoColumnView -libraryTitle $library.Title -fieldOneName "Document Type" -fieldTwoName "Topic"
    Create-GroupByTwoColumnView -libraryTitle $library.Title -fieldOneName "Topic" -fieldTwoName "Document Type"
    
}