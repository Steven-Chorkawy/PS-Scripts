# Copied from: https://www.sharepointdiary.com/2014/02/get-created-by-created-on-details-of-sharepoint-view.html

#Config Variables
$SiteURL = "https://claringtonnet.sharepoint.com/sites/StevensTestSite"
$ListName= "Death Register"
$ViewName = "All Items"
 
Try {
    #Connect to PnP Online
    Connect-PnPOnline -Url $SiteURL -Interactive
 
    #Get the Web
    $Web  = Get-PnPWeb
 
    #Get the List Views from the list
    $ListView  =  Get-PnPView -List $ListName -Identity $ViewName
 
    #Get the View File
    $ViewFile = $Web.GetFileByServerRelativeUrl($ListView.ServerRelativeUrl)    
    $Properties = Get-PnPProperty -ClientObject $ViewFile -Property Author, ModifiedBy, TimeCreated, TimeLastModified
 
    Write-Host "Created By: " $ViewFile.Author.Email
    Write-Host "Created on: " $ViewFile.TimeCreated

    # !!! These fields are not accurate or they are not updated fast enough.
    Write-Host "Modified By: " $ViewFile.ModifiedBy.Email
    Write-Host "Modified On: " $ViewFile.TimeLastModified

    Disconnect-PnPOnline
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}