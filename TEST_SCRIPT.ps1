Clear-Host

Connect-PnPOnline -Url "https://claringtonnet.sharepoint.com/sites/StevensTestSite" -Interactive


$Context = Get-PnPContext
$DOCUMENT_SET_ID = "0x0120D520"
$listName = "PSTest"
$viewName = "Group by Status"
$library = Get-PnPList -Identity $listName -Includes Views, ContentTypes, Title, DefaultView
$contentTypes = $library.ContentTypes

$ct = $contentTypes | Where-Object { $_.Name -eq "Document Set" }
Set-PnPView -List $listName -Identity $viewName -Values @{JSLink = "hierarchytaskslist.js|customrendering.js"; AssociatedContentTypeId = $ct.Id.ToString(); ContentTypeId = $ct.Id; DefaultViewForContentType = $true }

$views = $library.Views | Where-Object { $_.BaseViewId -eq "1" }
Write-Host "The End."