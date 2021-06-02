# This script will assign policy to user using teams module cmdlets

$logfile = ".\UserPolicyAssignmentlog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
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

$UserPrincipalNames = import-csv -path ".\PolicyAssignment.csv"
$UserPricipleNames = $PolicyAssignment.UserPricipleName
$count = $PolicyAssignment.Count
write-host "Running the script for users:" $count

foreach($UserPrincipalName in $UserPrincipalNames.UserPrincipalName)
{
Write-Host "To change the Applied Policy to user" $UserPrincipalName


        function Get-Result() {
        write-host           "1- TeamsAppSetupPolicy 
                      2- TeamsMeetingPolicy 
                      3- TeamsCallingPolicy
                      4- TeamsMessagingPolicy 
                      5- BroadcastMeetingPolicy
                      6- TeamsCallParkPolicy
                      7- CallerIdPolicy 
                      8- TeamsEmergencyCallingPolicy 
                      9- TeamsEmergencyCallRoutingPolicy
                      10-VoiceRoutingPolicy 
                      11-TeamsAppPermissionPolicy 
                      12-TeamsDailPlan"
$proceed = Read-host "Please provide the policy number to Grant and proceed further" 


if ($proceed -eq '1')
{
    try{
    Get-CsOnlineUser -Identity "$UserPrincipalName" | fl TeamsAppSetupPolicy 
    Get-CSTeamsAppsetuppolicy |fl Identity
    $PolicyName=Read-Host "Please provide the Policy Name"
    Grant-CsTeamsAppSetupPolicy -identity "$UserPrincipalName" -PolicyName  "$PolicyName"
    }
    catch{
   $_.Exception.Message | out-file -Filepath $logfile -append
    }
}


elseif ($proceed -eq '2')
{
    try{
    Get-CsOnlineUser -Identity "$UserPrincipalName" | fl TeamsMeetingPolicy 
    Get-CsTeamsMeetingPolicy  |fl Identity  
    Grant-CsTeamsMeetingPolicy -identity "$UserPrincipalName" -PolicyName "$PolicyName"
    }
    catch{
   $_.Exception.Message | out-file -Filepath $logfile -append
    }
}

elseif ($proceed -eq '3')
{
try{
    Get-CsOnlineUser -Identity "$UserPrincipalName" | fl TeamsCallingPolicy
    Get-CsTeamsCallingPolicy  |fl Identity
    $PolicyName= Read-Host "Please provide the Policy Name"
    Grant-CsTeamsCallingPolicy  -identity "$UserPrincipalName" -PolicyName  "$PolicyName"
    }
    catch{
   $_.Exception.Message | out-file -Filepath $logfile -append
    }
}

elseif ($proceed -eq '4')
{
    try{
    Get-CsOnlineUser -Identity "$UserPrincipalName" | fl TeamsMessagingPolicy 
    Get-CsTeamsMessagingPolicy |fl Identity
    $PolicyName = Read-Host "Please provide the Policy Name"
    Grant-CsTeamsMessagingPolicy -identity "$UserPrincipalName" -PolicyName  "$PolicyName"
    }
    catch{
   $_.Exception.Message | out-file -Filepath $logfile -append
    }
}


elseif ($proceed -eq '5')
{
    try{
    Get-CsOnlineUser -Identity "$UserPrincipalName" | fl BroadcastMeetingPolicy  
    Get-CsBroadcastMeetingPolicy   |fl Identity
    $PolicyName=Read-Host "Please provide the Policy Name"
    Grant-CsBroadcastMeetingPolicy  -identity "$UserPrincipalName" -PolicyName  "$PolicyName"
}
catch{
   $_.Exception.Message | out-file -Filepath $logfile -append
    }
}

elseif ($proceed -eq '6')
{
    try{
    Get-CsOnlineUser -Identity "$UserPrincipalName" | fl TeamsCallParkPolicy
    Get-CsTeamsCallParkPolicy|fl Identity
    $PolicyName=Read-Host "Please provide the Policy Name"
    Grant-CsTeamsCallParkPolicy   -identity "$UserPrincipalName" -PolicyName  "$PolicyName"
    }
    catch{
   $_.Exception.Message | out-file -Filepath $logfile -append
    }
}

elseif ($proceed -eq '7')
{
try{
    Get-CsOnlineUser -Identity "$UserPrincipalName" | fl CallerIdPolicy 
    Get-CsTeamsCallerIdPolicy|fl Identity
    $PolicyName=Read-Host "Please provide the Policy Name"
    Grant-CsTeamsCallerIdPolicy  -identity "$UserPrincipalName" -PolicyName  "$PolicyName"
    }
    catch{
   $_.Exception.Message | out-file -Filepath $logfile -append
    }
}
elseif ($proceed -eq '8')
{
    try{
    Get-CsOnlineUser -Identity "$UserPrincipalName" | fl TeamsEmergencyCallingPolicy 
    Get-CsTeamsEmergencyCallingPolicy    |fl Identity
    $PolicyName=Read-Host "Please provide the Policy Name"
    Grant-CsTeamsEmergencyCallingPolicy   -identity "$UserPrincipalName" -PolicyName  "$PolicyName"
    }
    catch{
   $_.Exception.Message | out-file -Filepath $logfile -append
    }
}
elseif ($proceed -eq '9')
{
    try{
    Get-CsOnlineUser -Identity "$UserPrincipalName" | fl TeamsEmergencyCallRoutingPolicy  
    Get-CsTeamsEmergencyCallRoutingPolicy    |fl Identity
    $PolicyName=Read-Host "Please provide the Policy Name"
    Grant-CsTeamsEmergencyCallRoutingPolicy -identity "$UserPrincipalName" -PolicyName  "$PolicyName"
    }
    catch{
   $_.Exception.Message | out-file -Filepath $logfile -append
    }
}
elseif ($proceed -eq '10')
{
    try{
    Get-CsOnlineUser -Identity "$UserPrincipalName" | fl VoiceRoutingPolicy 
    Get-CsVoiceRoutingPolicy    | fl Identity
    $PolicyName=Read-Host "Please provide the Policy Name"
    Grant-CsVoiceRoutingPolicy -identity "$UserPrincipalName" -PolicyName  "$PolicyName"
    }
    catch{
   $_.Exception.Message | out-file -Filepath $logfile -append
    }
}
elseif ($proceed -eq '11')
 {
    try{
    Get-CsOnlineUser -Identity "$UserPrincipalName" | fl TeamsAppPermissionPolicy  
    Get-CsTeamsAppPermissionPolicy   |fl Identity
    $PolicyName=Read-Host "Please provide the Policy Name"
    Grant-CsTeamsAppPermissionPolicy   -identity "$UserPrincipalName" -PolicyName  "$PolicyName"
    }
   catch{
   $_.Exception.Message | out-file -Filepath $logfile -append
    }
 }

elseif($proceed -eq '12')
  {
    try{
    Get-CsOnlineUser -Identity "$UserPrincipalName" | FL DialPlan  
    Get-CsDialPlan|fl Identity
    $PolicyName=Read-Host "Please provide the Policy Name"
    Grant-CsDialPlan -Identity "$UserPrincipalName" -PolicyName  "$PolicyName"
    }
    catch{
   $_.Exception.Message | out-file -Filepath $logfile -append
    }
   }           
}

do
{

$ProceedNext = Read-host "Do you want to apply policy,press Y to continue"
if ($ProceedNext -eq "Y" ) 
        { 
        Get-Result 
        } 
        else 
        { 
        break 
        } 
    }
    while($true); 
    }
        
$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan   
$resultTime.TotalSeconds 
#end of script
