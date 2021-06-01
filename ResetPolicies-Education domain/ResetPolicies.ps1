# This Script will create custom policies to set all access to teachers and restrict to students in an organization
$logfile = ".\ResetPolicieslog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now

If(Get-Module -ListAvailable -Name MicrosoftTeams) 
 { 
 Write-Host "MicrosoftTeams Already Installed" 
 } 
 else { 
 try { Install-Module -Name MicrosoftTeams
 Write-Host "Installed MicrosoftTeams"
 }
 catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
 } }     
 try{
Connect-MicrosoftTeams
}
 catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
 }  
PolicyChoice = Read-host = "'please provide Number to reset policies
1.Teams policies
2.Meetings Policies
3.Meeting Settings
4.Messaging policies
5.Assignment policy
6.OrgWide settings
7.Apps'"

if(PolicyChoice -eq "1")
{try{
#teams policies
Set-CsTeamsChannelsPolicy -Identity global -AllowPrivateChannelCreation $false
New-CsTeamsChannelsPolicy -Identity AllTeachers -AllowPrivateChannelCreation $true
}
catch{
  $_.Exception.Message | out-file -Filepath $logfile -append
    }
}
if(PolicyChoice -eq "2"){
try{
#Meetings Policies
Set-CsTeamsMeetingPolicy -Identity global -AllowMeetNow $false -AllowOutlookAddIn $false -AllowChannelMeetingScheduling $false -AllowPrivateMeetingScheduling $false -AllowTranscription $true -AllowCloudRecording $false  -AllowParticipantGiveRequestControl $false -AllowExternalParticipantGiveRequestControl $false -AllowPowerPointSharing $true -AllowWhiteboard $true -AllowSharedNotes $true -AllowAnonymousUsersToStartMeeting $false -AutoAdmittedUsers OrganizerOnly -LiveCaptionsEnabledType DisabledUserOverride -MeetingChatEnabledType enabled -AllowPrivateMeetNow $false
New-CsTeamsMeetingPolicy -Identity AllTeachers -AllowMeetNow $true -AllowOutlookAddIn $true -AllowChannelMeetingScheduling $true -AllowPrivateMeetingScheduling $true -AllowTranscription $true -AllowCloudRecording $false -AllowParticipantGiveRequestControl $true -AllowExternalParticipantGiveRequestControl $true -AllowPowerPointSharing $true -AllowWhiteboard $true -AllowSharedNotes $true -AllowAnonymousUsersToStartMeeting $false -AutoAdmittedUsers EveryoneInCompany -LiveCaptionsEnabledType DisabledUserOverride -MeetingChatEnabledType enabled -AllowPrivateMeetNow $true
}catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
}
if(PolicyChoice -eq "3"){
try{
#meeting Settings
set-CsTeamsMeetingConfiguration -DisableAnonymousJoin $true
}catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
}
if(PolicyChoice -eq "4"){
try{
#messaging policies
Set-CsTeamsMessagingPolicy -Identity global -AllowUserDeleteMessage $false -AllowUserEditMessage $false -ReadReceiptsEnabledType Everyone -AllowUserChat $false -AllowMemes $false -AllowPriorityMessages $false -AudioMessageEnabledType disabled -AllowRemoveUser $false -AllowSmartReply $false -ChannelsInChatListEnabledType EnabledUserOverride -AllowGiphy $false -AllowGiphyDisplay $false
new-CsTeamsMessagingPolicy -Identity AllTeachers -AllowUserDeleteMessage $true -AllowUserEditMessage $true -ReadReceiptsEnabledType Everyone -AllowUserChat $true -AllowMemes $false -AllowPriorityMessages $true -AudioMessageEnabledType ChatsAndChannels -AllowRemoveUser $true -AllowSmartReply $true -ChannelsInChatListEnabledType EnabledUserOverride -AllowGiphy $false -AllowGiphyDisplay $false -AllowOwnerDeleteMessage $true
}catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
}
if(PolicyChoice -eq "5"){
try{
#Assignment policy
Set-CsTeamsEducationAssignmentsAppPolicy -MakeCodeEnabledType disabled -TurnItInEnabledType enabled
}
catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
}
if(PolicyChoice -eq "6"){
try{
#OrgWide settings
Set-CsTeamsClientConfiguration -AllowEmailIntoChannel $false -AllowDropBox $false -AllowBox $false -AllowGoogleDrive $false -AllowShareFile $false -AllowGuestUser $false -AllowEgnyte $false -AllowOrganizationTab $false -AllowScopedPeopleSearchandAccess $true
}
catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
}
if(PolicyChoice -eq "7"){
try{
#apps
Set-CsTeamsAppPermissionPolicy -GlobalCatalogAppsType  AllowedAppList
}catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
}
else
{write-host "please run script again and choose option between 1-7"
}

$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
#end of script
