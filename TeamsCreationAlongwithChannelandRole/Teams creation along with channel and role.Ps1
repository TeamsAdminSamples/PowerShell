# This script is to create Teams along with channel and role by providing input data TeamName, ChannelName, Owner, Member and Visibility in TeamsCreationAlongWithChannelRole.csv file

$logfile = ".\TeamsCreationAlongWithChannelRolelog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now
Connect-MicrosoftTeams
$Groups= Import-Csv -path ".\TeamsCreationAlongWithChannelRole.csv"
foreach($Group in $Groups)
{
try{
$Teams = New-team -DisplayName $Group.TeamName  -Owner $group.Owner -Visibility $Group.Visibility
$groupid = $teams.GroupId
$teamchannel = New-TeamChannel -DisplayName $group.ChannelName -GroupId $groupid
$teamuser = Add-TeamUser -GroupId $groupid -User $group.Member -Role Member

if($Teams -ne $null)
{
            $file = New-Object psobject
            $file | add-member -MemberType NoteProperty -Name TeamName $Teams.DisplayName
            $file | add-member -MemberType NoteProperty -Name Teamid $Teams.groupid
            $file | add-member -MemberType NoteProperty -Name channelName $teamchannel.DisplayName
            $file | add-member -MemberType NoteProperty -Name channelId $teamchannel.ID
            $file | add-member -MemberType NoteProperty -Name TeamOwner $group.Owner
	    $file | add-member -MemberType NoteProperty -Name TeamMember $group.Member
            $file | export-csv -path ".\Output.csv"  -NoTypeInformation -Append
            }
	    }
catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
}

$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
#end of script
