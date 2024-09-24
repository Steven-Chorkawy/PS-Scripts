Clear-Host

Connect-PnPOnline -Url "https://claringtonnet.sharepoint.com/" -Interactive

$USER_INPUT = Read-Host "Enter comma delimited list of AD Groups to get members"
$ADGroupNames = $USER_INPUT.Split(",")
$GroupMemberOutput = @()
$EXCEL_EXPORT_PATH = "C:\Users\sc13\OneDrive - clarington.net\Desktop\Excel Export\ADGroupMemberReport-$(Get-Date -Format "MM-dd-yyyy").csv"

for($i=0; $i -lt $ADGroupNames.Count; $i++)
{
    $groupName = $ADGroupNames[$i].Trim()
    Write-Host ($i+1) "/" $ADGroupNames.Count "|" $groupName
    $groupMembers = Get-PnPAzureADGroupMember -Identity $groupName
    
    foreach($member in $groupMembers) 
    {
        $props = @{
            DisplayName = $member.DisplayName
            Group = $groupName
        }
        $GroupMemberOutput += New-Object -TypeName PsObject -Property $props
    }
}

Write-Host "Exporting to CSV..."
Write-Host $EXCEL_EXPORT_PATH

$GroupMemberOutput | Export-Csv -Path $EXCEL_EXPORT_PATH