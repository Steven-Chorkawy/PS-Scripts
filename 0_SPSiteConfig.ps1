Function MyConnect {
    param (
        [string]$Url
    )
    $getConn = Get-PnPWeb
    if ($getConn.Url -eq $Url) {
        Write-Host "Skipping connection" -ForegroundColor DarkMagenta
        return
    }
    Write-Host "`n`n" -ForegroundColor DarkMagenta
    Write-Host "Connecting to $($Url)" -ForegroundColor DarkMagenta
    Connect-PnPOnline -Url $Url -Interactive
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
            $getFieldResult = Get-PnPField -Identity $column
        }
        catch {
            Write-Host "`tFAIL! Could not find '$($column)'!  Please create this column." -ForegroundColor Red
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
    Write-Host "Update All Documents View for $($LibraryName)..."
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


Clear-Host
$excelRows = ""
try {
    $path = Read-Host -Prompt "Enter Path to Excel Template"
    $excelRows = Import-Excel -Path $path
}
catch {
    Write-Host "Failed to read excel file...  Please close the file before running this script..." -ForegroundColor Red
    Exit
}

# Create Content Types.
Write-Host "`n`nCreating Content Types..."
foreach ($row in $excelRows) {
    Write-Host "`n`nConnecting to $($row.SiteUrl) > $($row.Library_DisplayName)"
    MyConnect -Url $row.SiteUrl

    Write-Host "Validating Metadata Columns..."
    $documentColumnsFound = CheckMetadataColumns -Columns $row.Document_Columns
    $documentSetColumnsFound = CheckMetadataColumns -Columns $row.DocumentSet_Columns

    if ($documentColumnsFound -eq $false -or $documentSetColumnsFound -eq $false) {
        Write-Host "Some metadata columns cannot be found for $($row.SiteUrl) > $($row.Library_DisplayName)." -ForegroundColor Red
        Write-Host "Skipping Content Type Creation." -ForegroundColor Red
        # This will skip the current interation of the foreach loop.
        continue
    }

    $documentSet_ParentContentType = ""
    $document_ParentContentType = ""
    
    if ($row.Document_ParentContentType -eq "Document" -and $row.DocumentSet_ParentContentType -eq "Document Set") {
        $output = GetDefaultDocumentContentTypes -DocumentName $row.Document_ParentContentType -DocumentSetName $row.DocumentSet_ParentContentType
        $document_ParentContentType = $output["Document"]
        $documentSet_ParentContentType = $output["DocumentSet"]
    }
    else {
        $output = GetEDRMContentTypes -siteURL $row.SiteUrl -DocumentName $row.Document_ParentContentType -DocumentSetName $row.DocumentSet_ParentContentType
        $document_ParentContentType = $output[1]["Document"]
        $documentSet_ParentContentType = $output[1]["DocumentSet"]
    }

    # Try to create the Document and Document Set Content Type and add metadata columns.
    try {
        Write-Host "Creating $($row.Document_ContentTypeName) Content Type..."
        $addDocRes = Add-PnPContentType -Name $row.Document_ContentTypeName -Group "Custom Content Types" -ParentContentType $document_ParentContentType
        foreach ($column in ConvertColumnStringToArray -Columns $row.Document_Columns) {
            Write-Host "`tAdding '$($column)' column to '$($row.Document_ContentTypeName)' Content Type."
            Add-PnPFieldToContentType -ContentType $row.Document_ContentTypeName -Field $column
        }

        # Try to create the Document Set Content Type and add metadata columns.
        try {
            Write-Host "Creating $($row.DocumentSet_ContentTypeName) Content Type..."
            $addDocSetRes = Add-PnPContentType -Name $row.DocumentSet_ContentTypeName -Group "Custom Content Types" -ParentContentType $documentSet_ParentContentType
            foreach ($column in ConvertColumnStringToArray -Columns $row.DocumentSet_Columns) {
                Write-Host "`tAdding '$($column)' column to '$($row.DocumentSet_ContentTypeName)' Content Type."
                Add-PnPFieldToContentType -ContentType $row.DocumentSet_ContentTypeName -Field $column
                Write-Host "`tAdding '$($row.Document_ContentTypeName)' to $($row.DocumentSet_ContentTypeName)."
                Add-PnPContentTypeToDocumentSet -ContentType $row.Document_ContentTypeName -DocumentSet $row.DocumentSet_ContentTypeName
                Write-Host "`tRemoving default 'Document' from $($row.DocumentSet_ContentTypeName)"
                Remove-PnPContentTypeFromDocumentSet -ContentType "Document" -DocumentSet $row.DocumentSet_ContentTypeName
            }
        }
        catch {
            Write-Host "Failed to create '$($row.Document_ContentTypeName)' Content Type!" -ForegroundColor Red
            Write-Host $_
        }
    }
    catch {
        Write-Host "Failed to create '$($row.Document_ContentTypeName)' Content Type!" -ForegroundColor Red
        Write-Host $_
    }
}

# Create Libraries.
Write-Host "`n`nCreating Libraries..."
foreach ($row in $excelRows) {
    Write-Host "`n`nConnecting to $($row.SiteUrl) > $($row.Library_DisplayName)"
    MyConnect -Url $row.SiteUrl
    $libraryUrl = $row.Library_Name
    $libraryDisplayName = $row.Library_DisplayName
    # Create new library with settings
    New-PnPList -Title $libraryDisplayName -Url $libraryUrl -Template DocumentLibrary -EnableVersioning -OnQuickLaunch -EnableContentTypes

    # Set MajorVersion limit to 100
    Set-PnPList -Identity $libraryUrl -MajorVersions 100 -EnableFolderCreation $false

    # Get library
    $library = Get-PnPList -Identity $libraryUrl -Includes ContentTypes
    # Add Document Set Content Type
    Add-PnPContentTypeToList -List $libraryUrl -ContentType $row.DocumentSet_ContentTypeName
    # Add Document Content Type
    Add-PnPContentTypeToList -List $libraryUrl -ContentType $row.Document_ContentTypeName

    # Update the All Documents view to sort by Document Sets first. 
    # https://clarington.freshservice.com/a/tickets/40827?current_tab=details
    Update-AllDocuments-View -SiteURL $row.SiteUrl -LibraryName $libraryDisplayName

    # Remove Document and Folder Content Type
    Remove-PnPContentType -List $libraryUrl -ContentType "Document"
    Remove-PnPContentType -List $libraryUrl -ContentType "Folder"

    # Remove description field.
    Remove-PnPField -List $libraryUrl -Identity "_ExtendedDescription" -Force
}