Clear-Host
$PnPID = "08a2a4b7-9f6f-46cb-87e7-2ed02a66fc22"
$siteUrl = Read-Host "Enter URL of the site to save as a template (e.g., https://Claringtonnet.sharepoint.com/sites/yoursite)"
$Path = Read-Host "Enter the path where you'd like to save the site template (e.g., C:\Temp\MoCCommittee.pnp)"

Connect-PnPOnline -url $siteUrl -ClientId $PnPID -Interactive
Get-PnPSiteTemplate -out $Path -ListsToExtract "TimesheetTasks", "TimesheetProjects", "Timesheet" -ExcludeHandlers ApplicationLifecycleManagement, AuditSettings, ComposedLook, CustomActions, ExtensibilityProviders, Features, Files, ImageRenditions, Navigation, PageContents, Pages, PropertyBagEntries, Publishing, RegionalSettings, SearchSettings, SiteFooter, SiteHeader, SitePolicy, SiteSecurity, SiteSettings, SupportedUILanguages, SyntexModels, Tenant, TermGroups, Theme, WebApiPermissions, WebSettings, Workflows -Force
#-ExcludeHandlers SiteSecurityTime