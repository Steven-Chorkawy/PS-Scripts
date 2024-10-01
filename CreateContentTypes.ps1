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

Clear-Host
$excelRows = ""
try {
    $excelRows = Import-Excel -Path "C:\Users\sc13\OneDrive - clarington.net\Desktop\SharePoint Site Config Template.xlsx"
}
catch {
    Write-Host "Failed to read excel file...  Please close the file before running this script..." -ForegroundColor Red
    Exit
}

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