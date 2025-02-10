$ALL_METADATA_COLUMNS_FOUND = $true
$DOCUMENT_SET_FEATURE_GUID = "3bae86a2-776d-499d-9db8-fa4cdc7884f8" # Hard coded value from SharePoint.
$DOCUMENT_SET_FEATURE_NAME = "DocumentSet"

Function MyConnect {
    param (
        [string]$Url
    )
    $getConn = Get-PnPWeb | Out-Null
    if ($getConn.Url -eq $Url) {
        return
    }
    Connect-PnPOnline -Url $Url -Interactive
    Write-Host "Connecting to $($Url)" -ForegroundColor DarkMagenta
}

Function WriteNewTitle {
    param (
        [string]$Title
    )
    Write-Host "`n========================================================================================================" -ForegroundColor Magenta
    Write-Host $Title -ForegroundColor Magenta
    Write-Host "========================================================================================================" -ForegroundColor Magenta
}

Function GetDefaultDocumentContentTypes {
    param (
        [string]$DocumentName,
        [string]$DocumentSetName
    )
    $document_ParentContentType = Get-PnPContentType | Where-Object { $_.Id -match "0x0101" -and $_.Name -eq $DocumentName }
    $documentSet_ParentContentType = Get-PnPContentType | Where-Object { $_.Id -match "0x0120D520" -and $_.Name -eq $DocumentSetName }
    return @{Document = $document_ParentContentType; DocumentSet = $documentSet_ParentContentType }
}

Function GetEDRMContentTypes {
    param (
        [string]$siteURL,
        [string]$DocumentName,
        [string]$DocumentSetName
    )
    $doc_ParentContentType = Get-PnPCompatibleHubContentTypes -WebUrl $SiteUrl | Where-Object { $_.Group -eq "Organizational Content Types" -and $_.ParentName -eq "EDRM Document" -and $_.Name -match $DocumentName }
    $docSet_ParentContentType = Get-PnPCompatibleHubContentTypes -WebUrl $SiteUrl | Where-Object { $_.Group -eq "Organizational Content Types" -and $_.ParentName -eq "EDRM Document Set" -and $_.Name -match $DocumentSetName }
    Add-PnPContentTypesFromContentTypeHub -ContentTypes @($docSet_ParentContentType.Id, $doc_ParentContentType.Id)
    $documentSet_ParentContentType = Get-PnPContentType | Where-Object { $_.Id -match "0x0120D520" -and $_.Name -match $DocumentSetName }
    $document_ParentContentType = Get-PnPContentType | Where-Object { $_.Id -match "0x0101" -and $_.Name -match $DocumentName }
    return @{Document = $document_ParentContentType; DocumentSet = $documentSet_ParentContentType }
}

Function ConvertColumnStringToArray {
    param (
        [string]$Columns
    )
    return $Columns.Split(", ")
}

Function CheckMetadataColumns {
    [OutputType([Boolean])]
    param (
        [string] $Columns
    )
    $columnArray = ConvertColumnStringToArray -Columns $Columns
    $output = $true
    foreach ($column in $columnArray) {
        try {
            # TODO: Trim the column name string before checking if it exists.
            Get-PnPField -Identity $column
        }
        catch {
            $output = CreateMetadataColumns -Column $column
        }
    }
    return $output
}

Function CreateMetadataColumns {
    [OutputType([Boolean])]
    param(
        [string] $Column
    )
    $output = $true
    switch ($Column) {
        "Topic" {
            Add-PnPField -Type Text -InternalName "Topic" -DisplayName "Topic" -Group "Custom Columns"
            Write-Host "`t'Topic' has been created" 
        }
        "Year" {
            Add-PnPField -Type Text -InternalName "Year" -DisplayName "Year" -Group "Custom Columns"
            Set-PnPField -Identity "Year" -Values @{ DefaultFormula = "=CONCATENATE(YEAR(Today))" }
            Write-Host "`t'Year' has been created"
        }
        "Document Type" {
            Add-PnPField -Type Choice -InternalName "DocumentType" -DisplayName "Document Type" -Group "Custom Columns" -Choices "Choice #1", "Choice #2", "Choice #3"
            Write-Host "`t'Document Type' has been created"
        }
        "Month" {
            Add-PnPField -Type Choice -InternalName "Month" -DisplayName "Month" -Group "Custom Columns" -Choices "01-Jan", "02-Feb", "03-Mar", "04-Apr", "05-May", "06-Jun", "07-Jul", "08-Aug", "09-Sep", "10-Oct", "11-Nov", "12-Dec"
            Write-Host "`t'Month' has been created"
        }
        "Department" {
            # Create a Department field with a Manged Metadata type.
            Add-PnPTaxonomyField -DisplayName "Department" -InternalName "MM_Department" -TermSetPath "MOC Org|Department" -Group "Custom Columns"
            Write-Host "`t'Department' has been created"
        }
        Default {
            Write-Host "`tPlease create column '$($Column)'" -ForegroundColor Red
            $output = $false
        }
    }
    return $output
}

# Function to format the All Documents view to sort by DocIcon (file type) in Ascending order.  This will sort by Folders first, then by documents. 
Function Update-AllDocuments-View {
    param(
        [string]$SiteURL,
        [string]$LibraryName
    )
    Write-Host "`tUpdate All Documents View for $($LibraryName)..."
    $ViewName = "All Documents"
    $ViewQuery = "<OrderBy><FieldRef Name='DocIcon' Ascending='TRUE'/><FieldRef Name='LinkFilename' Ascending='TRUE'/></OrderBy>"
    #Get the Client Context
    $Context = Get-PnPContext
    #Get the List View
    $View = Get-PnPView -Identity $ViewName -List $LibraryName
    #Update the view Query
    $View.ViewQuery = $ViewQuery
    $View.Update() 
    $Context.ExecuteQuery()
}

<#
.SYNOPSIS
Daisplay all custom metadata columns on the All Documents view.

.DESCRIPTION
1. Sort by DocIcon, then by File Name.
2. Add all custom metadata columns to the All Documents view.

.PARAMETER libraryTitle
Display name of the SharePoint library

.EXAMPLE
Update-AllDocumentsViewColumns -libraryTitle "My Library"

.NOTES
General notes
#>
Function Update-AllDocumentsViewColumns {
    param(
        [string]$libraryTitle
    )
    $selectedColumns = Get-PnPField -List $libraryDisplayName | Where-Object { $_.Hidden -eq $false -and $_.CanBeDeleted -eq $true -and $_.ReadOnlyField -eq $false -and $_.InternalName -ne "_ExtendedDescription" } 

    # Order by DocIcon as per Brandi. 
    $BrandiViewQuery = "<OrderBy><FieldRef Name='DocIcon' Ascending='TRUE'/><FieldRef Name='LinkFilename' Ascending='TRUE'/></OrderBy>"
    #Get the Client Context
    $Context = Get-PnPContext
    $allDocumentsView = Get-PnPView -List $libraryTitle -Identity "All Documents" -Includes ViewType, ViewFields, Aggregations, Paged, ViewQuery, RowLimit
    $allDocumentsView.ViewQuery = $BrandiViewQuery
    $allDocumentsView.Update()
    $Context.ExecuteQuery()

    # The first two columns in the view will always be file type and file name.
    $fieldArray = @("DocIcon", "LinkFilename")

    # The columns will be any custom columns.
    foreach ($column in $selectedColumns) {
        if ($fieldArray -notcontains $column.InternalName) {
            $fieldArray += $column.InternalName
        }
    }

    # The last two columns will always be Modified and Modified By
    $fieldArray += "Modified"
    $fieldArray += "Editor"

    Set-PnPView -List $libraryTitle -Identity "All Documents" -Fields $fieldArray | Out-Null
}

<#
.SYNOPSIS
For each Choice field in a library create a view that groups by that field.

.DESCRIPTION
For each Choice field in a library create a view that groups by that field.

.PARAMETER libraryTitle
Display name of the SharePoint library.

.EXAMPLE
Create-CustomChoiceViews -libraryTitle "My Library"

.NOTES
General notes
#>
Function Create-CustomChoiceViews {
    param(
        [string]$libraryTitle
    )

    $fields = Get-PnPField -List $libraryDisplayName | Where-Object { $_.TypeAsString -eq "Choice" }
    foreach ($field in $fields) {
        Create-GroupByOneColumnView -libraryTitle $libraryTitle -fieldName $field.Title
    }
}

<#
.SYNOPSIS
Create a view that groups by one columns.

.DESCRIPTION
Create a view that groups by one columns.

.PARAMETER libraryTitle
Display name of the SharePoint library.

.PARAMETER fieldName
Display name of the column used to group by.

.EXAMPLE
Create-GroupByOneColumnView -libraryTitle "My Title" -fieldName "My Column"

.NOTES
General notes
#>
Function Create-GroupByOneColumnView {
    param(
        [string]$libraryTitle,
        [string]$fieldName
    )

    $newViewName = "Group by $($fieldName)"
    # If topicField is null we cannot create the view.
    $newField = Get-PnPField -Identity $fieldName -List $libraryTitle -ErrorAction SilentlyContinue
    # If newView is NOT null the view already exists and we do not need to create another one.
    $newView = Get-PnPView -Identity $newViewName -List $libraryTitle -ErrorAction SilentlyContinue
    $allDocumentsView = Get-PnPView -List $libraryTitle -Identity "All Documents" -Includes ViewType, ViewFields, Aggregations, Paged, ViewQuery, RowLimit

    If ($newView) {       
        Write-Host "`t$($newViewName) already exists in $($libraryTitle)... Skipping this step." -ForegroundColor Yellow
    }
    else {
        # View does not exist.  Create proceed to create the view.
        If ($newField) {
            # Update the view properties of the All Documents view.
            $allDocumentsView.ViewFields.add($fieldName)
            $allDocumentsView.ViewQuery = "$($allDocumentsView.ViewQuery)<GroupBy Collapse='TRUE' GroupLimit='30'><FieldRef Name='$($newField.InternalName)' Ascending='TRUE' /></GroupBy>"

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
            Add-PnPView @ViewProperties | Out-Null
            Write-Host "`t$($newViewName) View has been created!" -ForegroundColor Green
        }
        else {
            Write-Host "`tWarning! $($fieldName) field not found in $($libraryTitle) library!" -ForegroundColor Red
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
Create-GroupByTwoColumnView -libraryTitle $libraryDisplayName -fieldOneName "Document Type" -fieldTwoName "Topic"

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
    # If topicField is null we cannot create the view.
    $firstField = Get-PnPField -Identity $fieldOneName -List $libraryTitle -ErrorAction SilentlyContinue
    $secondField = Get-PnPField -Identity $fieldTwoName -List $libraryTitle -ErrorAction SilentlyContinue

    # If newView is NOT null the view already exists and we do not need to create another one.
    $newView = Get-PnPView -Identity $newViewName -List $libraryTitle -ErrorAction SilentlyContinue
    $allDocumentsView = Get-PnPView -List $libraryTitle -Identity "All Documents" -Includes ViewType, ViewFields, Aggregations, Paged, ViewQuery, RowLimit

    If ($newView) {       
        Write-Host "`t$($newViewName) already exists in $($libraryTitle)... Skipping this step." -ForegroundColor Yellow
    }
    else {
        # View does not exist.  Create proceed to create the view.
        If ($firstField) {
            If ($secondField) {
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

                Add-PnPView @ViewProperties | Out-Null
                Write-Host "`t$($newViewName) View has been created!" -ForegroundColor Green
            }
            else {
                Write-Host "`tWarning! $($fieldTwoName) field not found in $($libraryTitle) library!" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "`tWarning! $($fieldOneName) field not found in $($libraryTitle) library!" -ForegroundColor Yellow
        }
    }
}

############################################ * START OF SCRIPT * ############################################
Clear-Host
$excelRows = ""
try {
    $path = Read-Host -Prompt "Enter Path to Excel Template"
    if ($path -eq "") {
        Write-Host "No Path Provided... Using Default Testing Path..." -ForegroundColor Yellow
        $path = "C:\Users\sc13\OneDrive - clarington.net\Desktop\SharePoint Site Config Template.xlsx"
    }
    $excelRows = Import-Excel -Path $path
}
catch {
    Write-Host "Failed to read excel file...  Please close the file before running this script..." -ForegroundColor Red
    Exit
}

foreach ($row in $excelRows) {
    WriteNewTitle -Title "New Row"
    $row | Format-List
    if ($null -eq $row.SiteUrl) {
        Write-Host "Invalid row found in excel.  Skipping current row..." -ForegroundColor Yellow
        continue
    }
    
    $ALL_METADATA_COLUMNS_FOUND = $true # Reset this variable at the start of each loop.
    $ALL_CONTENT_TYPES_FOUND = $true # Reset this variable at the start of each loop.

    # Attempt to connect to the current site URL.
    MyConnect -Url $row.SiteUrl

    # Check and create for the Document Set site collection feature.
    Write-Host "Checking if Document Set feature is enabled..."
    $AreDocumentSetsEnabled = Get-PnPFeature -Scope Site | Where-Object { $_.DisplayName -eq $DOCUMENT_SET_FEATURE_NAME -and $_.DefinitionId -eq $DOCUMENT_SET_FEATURE_GUID }
    if ($null -eq $AreDocumentSetsEnabled) {
        Write-Host "Enabling Document Set Feature..." -ForegroundColor Green
        Enable-PnPFeature -Identity $DOCUMENT_SET_FEATURE_GUID -Scope Site
    }

    # Always delete the default Department field and replace it with a Managed Metadata field.
    try {
        $DEPARTMENT_COLUMN = Get-PnPField -Identity "ol_Department" 
        if ($DEPARTMENT_COLUMN.TypeAsString -eq "Text") {
            Remove-PnPField -Identity "ol_Department" -Force     # Delete Department text field.
            CreateMetadataColumns -Column "Department" | Out-Null  # Create Department Managed Metadata field.    
        }
    }
    catch {
        <# No action required.  Catch here to supress error message.#>
    }

    # Check for metadata columns.  Common columns will be created.
    # If a column cannot be automatically created the current iteration will be skipped.
    WriteNewTitle -Title "Validating Metadata Columns for $($row.SiteUrl) > $($row.Library_DisplayName)"
    $documentColumnsFound = CheckMetadataColumns -Columns $row.Document_Columns
    $documentSetColumnsFound = CheckMetadataColumns -Columns $row.DocumentSet_Columns

    # Notify the user that the current iteration will be skipped.
    if ($documentColumnsFound -eq $false -or $documentSetColumnsFound -eq $false) {
        Write-Host "Some metadata columns cannot be found for $($row.SiteUrl) > $($row.Library_DisplayName)." -ForegroundColor Red
        Write-Host "Skipping Content Type Creation." -ForegroundColor Red
        $ALL_METADATA_COLUMNS_FOUND = $false # Setting this to false will skip later steps that require metadata columns.
        continue # This will skip the current interation of the foreach loop.
    }

    if ($ALL_METADATA_COLUMNS_FOUND) {
        $documentSet_ParentContentType = ""
        $document_ParentContentType = ""
    
        try {
            if ($row.Document_ParentContentType -eq "Document" -and $row.DocumentSet_ParentContentType -eq "Document Set") {
                # This handles default parent content types.
                $output = GetDefaultDocumentContentTypes -DocumentName $row.Document_ParentContentType -DocumentSetName $row.DocumentSet_ParentContentType
                $document_ParentContentType = $output["Document"]
                $documentSet_ParentContentType = $output["DocumentSet"]
            }
            else {
                # This handles custom parent content types.
                $output = GetEDRMContentTypes -siteURL $row.SiteUrl -DocumentName $row.Document_ParentContentType -DocumentSetName $row.DocumentSet_ParentContentType
                $document_ParentContentType = $output[1]["Document"]
                $documentSet_ParentContentType = $output[1]["DocumentSet"]
            }
        }
        catch {
            Write-Host "Failed to get parent Content Types" -ForegroundColor Red
            Write-Host $_
            $ALL_CONTENT_TYPES_FOUND = $false
        }

        # Try to create the Document and Document Set Content Type and add metadata columns.
        try {
            # Check if Content Type exists.
            $currentDocumentContentType = Get-PnPContentType -Identity $row.Document_ContentTypeName
        
            # Create content type if it does not exist.
            if (!$currentDocumentContentType) {
                WriteNewTitle -Title "Creating $($row.Document_ContentTypeName) Content Type for $($row.SiteUrl) > $($row.Library_DisplayName)"
                # Create the Document Content Type.
                Add-PnPContentType -Name $row.Document_ContentTypeName -Group "Custom Content Types" -ParentContentType $document_ParentContentType | Out-Null
            }
           
            if ($row.Document_Columns) {
                foreach ($column in ConvertColumnStringToArray -Columns $row.Document_Columns) {
                    # Add metadata columns to the Document Content Type.
                    Add-PnPFieldToContentType -ContentType $row.Document_ContentTypeName -Field $column
                    Write-Host "`tAdding '$($column)' column to '$($row.Document_ContentTypeName)' Content Type."
                }   
            }

            # Try to create the Document Set Content Type and add metadata columns.
            try {
                #Check if Content Type exists.
                $currentDocumentSetContentType = Get-PnPContentType -Identity $row.DocumentSet_ContentTypeName 

                # Create content type if it does not exist.
                if (!$currentDocumentSetContentType) {
                    WriteNewTitle -Title "Creating $($row.DocumentSet_ContentTypeName) Content Type for $($row.SiteUrl) > $($row.Library_DisplayName)"
                    # Create the Document Set Content Type.
                    Add-PnPContentType -Name $row.DocumentSet_ContentTypeName -Group "Custom Content Types" -ParentContentType $documentSet_ParentContentType | Out-Null                
                }

                if ($row.DocumentSet_Columns) {
                    foreach ($column in ConvertColumnStringToArray -Columns $row.DocumentSet_Columns) {
                        # Add metadata to Document Set content type.
                        Add-PnPFieldToContentType -ContentType $row.DocumentSet_ContentTypeName -Field $column
                        Write-Host "`tAdding '$($column)' column to '$($row.DocumentSet_ContentTypeName)' Content Type."
                        # Add default content type to the Document Set content type.
                        Add-PnPContentTypeToDocumentSet -ContentType $row.Document_ContentTypeName -DocumentSet $row.DocumentSet_ContentTypeName
                        Write-Host "`tAdding '$($row.Document_ContentTypeName)' to $($row.DocumentSet_ContentTypeName)."
                        # Remove the default Document content type from the Document Set content type.
                        Remove-PnPContentTypeFromDocumentSet -ContentType "Document" -DocumentSet $row.DocumentSet_ContentTypeName
                        Write-Host "`tRemoving default 'Document' from $($row.DocumentSet_ContentTypeName)"
                    }   
                }
            }
            catch {
                Write-Host "Failed to create '$($row.Document_ContentTypeName)' Content Type!" -ForegroundColor Red
                Write-Host $_
                $ALL_CONTENT_TYPES_FOUND = $false
            }
        }
        catch {
            Write-Host "Failed to create '$($row.Document_ContentTypeName)' Content Type!" -ForegroundColor Red
            Write-Host $_
            $ALL_CONTENT_TYPES_FOUND = $false
        }

        if ($ALL_CONTENT_TYPES_FOUND) {
            # Create Libraries.
            WriteNewTitle -Title "Creating '$($row.Library_DisplayName)' Library..."
            $libraryUrl = $row.Library_Name
            $libraryDisplayName = $row.Library_DisplayName
    
            # Create new library with settings
            New-PnPList -Title $libraryDisplayName -Url $libraryUrl -Template DocumentLibrary -EnableVersioning -OnQuickLaunch -EnableContentTypes | Out-Null
            Write-Host "`tLibrary Created."
    
            # Set MajorVersion limit to 100
            Set-PnPList -Identity $libraryUrl -MajorVersions 100 -EnableFolderCreation $false | Out-Null
            Write-Host "`tLibrary Updated."
    
            # Add Document Set Content Type
            Add-PnPContentTypeToList -List $libraryUrl -ContentType $row.DocumentSet_ContentTypeName
    
            # Add Document Content Type
            Add-PnPContentTypeToList -List $libraryUrl -ContentType $row.Document_ContentTypeName
            Write-Host "`tContent Types added."
    
            # Update the All Documents view to sort by Document Sets first. 
            # https://clarington.freshservice.com/a/tickets/40827?current_tab=details
            Update-AllDocuments-View -SiteURL $row.SiteUrl -LibraryName $libraryDisplayName

            # Remove Document and Folder Content Type
            Remove-PnPContentTypeFromList -List $libraryDisplayName -ContentType "Document"
            Remove-PnPContentTypeFromList -List $libraryDisplayName -ContentType "Folder"
            
            # Remove description field.
            Remove-PnPField -List $libraryUrl -Identity "_ExtendedDescription" -Force
            
            WriteNewTitle -Title "Creating Views for $($row.SiteUrl) > $($row.Library_DisplayName)"
            Update-AllDocumentsViewColumns -libraryTitle $libraryDisplayName
            Create-CustomChoiceViews -libraryTitle $libraryDisplayName
            $customColumns = Get-PnPField -List $libraryDisplayName | Where-Object { $_.Hidden -eq $false -and $_.CanBeDeleted -eq $true -and $_.ReadOnlyField -eq $false -and $_.InternalName -ne "DocumentSetDescription" -and $_.InternalName -ne "_ExtendedDescription" } 
            foreach ($column in $customColumns) {
                Create-GroupByOneColumnView -libraryTitle $libraryDisplayName -fieldName $column.Title
            }
            
            # Always try to create these 2x group by views.
            Create-GroupByTwoColumnView -libraryTitle $libraryDisplayName -fieldOneName "Topic" -fieldTwoName "Year"
            Create-GroupByTwoColumnView -libraryTitle $libraryDisplayName -fieldOneName "Year" -fieldTwoName "Topic"
            Create-GroupByTwoColumnView -libraryTitle $libraryDisplayName -fieldOneName "Year" -fieldTwoName "Month"
            Create-GroupByTwoColumnView -libraryTitle $libraryDisplayName -fieldOneName "Month" -fieldTwoName "Year"
            Create-GroupByTwoColumnView -libraryTitle $libraryDisplayName -fieldOneName "Document Type" -fieldTwoName "Topic"
            Create-GroupByTwoColumnView -libraryTitle $libraryDisplayName -fieldOneName "Topic" -fieldTwoName "Document Type"
        } 
        else {
            Write-Host "Cannot create library..."
        }
    }
}

WriteNewTitle -Title "`tDisconnect-PnPOnline"
Disconnect-PnPOnline
############################################ * END OF SCRIPT * ############################################