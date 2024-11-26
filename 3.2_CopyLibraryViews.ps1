#https://www.sharepointdiary.com/2017/08/sharepoint-online-copy-list-views-using-powershell.html#h-pnp-powershell-to-clone-a-view-in-sharepoint-online

Clear-Host
#Parameters
$SiteURL = Read-Host "Enter Site URL"
$SourceListName = Read-Host "Enter Source List Name"
$SourceViewName = Read-Host "Enter Source View Name"
$TargetLists = @()
 
Try {
    #Connect to PnP Online
    Connect-PnPOnline -Url $SiteURL -Interactive
    #Get all document libraries - Exclude Hidden Libraries
    $DocumentLibraries = Get-PnPList -Includes DefaultView | Where-Object { $_.Hidden -eq $false -and $_.Title -ne "Form Templates" -and $_.Title -ne "Site Assets" -and $_.Title -ne "Style Library" } #Or $_.BaseType -eq "DocumentLibrary"

    $TargetLists = $DocumentLibraries | Out-GridView -Title "Select Libraries" -OutputMode Multiple
    
    #Get the Source View
    $SourceView = Get-PnPView -List $SourceListName -Identity $SourceViewName -Includes ViewQuery, ViewFields

    #Get the Client Context
    $Context = Get-PnPContext

    foreach ($list in $TargetLists) {
        Write-Host "Setting $($list.Title) - $($SourceViewName)"
        $TargetView = Get-PnPView -List $list.Title -Identity $SourceViewName -Includes ViewQuery
        $TargetView.ViewQuery = $SourceView.ViewQuery
        $TargetView.Update()
        $Context.ExecuteQuery()
        $setViewResult = Set-PnPView -List $list.Title -Identity $SourceViewName -Fields @($SourceView.ViewFields)
    }   
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}