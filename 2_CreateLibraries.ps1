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

# Clear host
Clear-Host

# Load SharePoint PnP module
Import-Module SharePointPnPPowerShellOnline -ErrorAction SilentlyContinue

# Connect to SharePoint tenant and retrieve list of sites
Connect-PnPOnline -Url "https://claringtonnet-admin.sharepoint.com" -Interactive
$sites = Get-PnPTenantSite

# Prompt user to select a site URL from the list
$selectedSite = $sites | Out-GridView -Title "Select site URL" -OutputMode Single


do {
    # Prompt user to enter library URL
    $libraryUrl = Read-Host "Enter library URL"

    # Prompt user to enter library display name
    $libraryDisplayName = Read-Host "Enter library display name"

    # Connect to selected site
    Connect-PnPOnline -Url $selectedSite.Url -Interactive

    # Create new library with settings
    New-PnPList -Title $libraryDisplayName -Url $libraryUrl -Template DocumentLibrary -EnableVersioning -OnQuickLaunch -EnableContentTypes

    # Set MajorVersion limit to 100
    Set-PnPList -Identity $libraryUrl -MajorVersions 100 -EnableFolderCreation $false

    # Get library
    $library = Get-PnPList -Identity $libraryUrl -Includes ContentTypes

    # Prompt user to select content types
    $contentTypes = Get-PnPContentType | Out-GridView -Title "Select content types for library" -OutputMode Multiple

    # Add selected content types to library
    foreach ($contentType in $contentTypes) {
        Add-PnPContentTypeToList -List $libraryUrl -ContentType $contentType.Name
    }

    # Retrieve the updated library object
    $library = Get-PnPList -Identity $libraryUrl -Includes ContentTypes

    # Update the All Documents view to sort by Document Sets first. 
    # https://clarington.freshservice.com/a/tickets/40827?current_tab=details
    Update-AllDocuments-View -SiteURL $selectedSite.Url -LibraryName $libraryDisplayName

    # Prompt user to select content types to remove
    $selectedContentTypesToRemove = $library.ContentTypes | Out-GridView -Title "Select content types to remove" -OutputMode Multiple

    # Remove selected content types from the library
    # After adding a custom content type to a library it's common to want to remove the default Document and Folder content types.
    foreach ($contentType in $selectedContentTypesToRemove) {
        Remove-PnPContentTypeFromList -List $libraryUrl -ContentType $contentType.Name
    }

    # Remove the default Description field.
    $extendedDescriptionField = Get-PnPField -List $libraryUrl -Identity "_ExtendedDescription" -ErrorAction SilentlyContinue
    if($extendedDescriptionField) {
        Write-Host "Removing _ExtendedDescription field..."
        Remove-PnPField -List $libraryUrl -Identity "_ExtendedDescription"
    } 
    else {
        Write-Host "Could not remove _ExtendedDescription field..." -ForegroundColor Red
    }

    # Ask user if they want to create a new library for the same site
    $createNewLibrary = Read-Host "Do you want to create a new library for the same site? (y/n)"
} while($createNewLibrary -eq "y")