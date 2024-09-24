#
# This script gets all the users from a given group and outputs the result in a csv file.
#

Clear-Host
Connect-PnPOnline -Url "https://claringtonnet.sharepoint.com/" -Interactive

$ADGroupName = "PrinterPlanningM608"

$UserProfiles = @()

# -Group takes ID or a name of a group
$GroupMembers = Get-PnPGroupMember -Group "SPCommServCSLTEditors"

$GroupMembers

#Connect-PnPOnline -Url "https://claringtonnet-admin.sharepoint.com/" -Interactive

for($i=0; $i -lt $GroupMembers.Count; $i++)
{  
    Write-Host $i "/" $GroupMembers.Count "|" $GroupMembers[$i].Email

    $profile = Get-PnPUserProfileProperty -Account $GroupMembers[$i].LoginName -Connection $adminConnection
    $enabled = Get-ADUser -Filter "Mail -eq '$($GroupMembers[$i].Email)'" | Select-Object -Property Enabled
    
    $Props = @{
        Enabled = $enabled.Enabled
        UserName = $profile.UserProfileProperties.UserName
        WorkEmail = $profile.UserProfileProperties.WorkEmail
        Title = $profile.UserProfileProperties.Title
        Department = $profile.UserProfileProperties.Department
    }
    $UserProfiles += New-Object -TypeName PsObject -Property $Props
}

$UserProfiles | Export-Csv -Path "C:\Users\sc13\Desktop\Group Export\$ADGroupName_Group Members.csv"