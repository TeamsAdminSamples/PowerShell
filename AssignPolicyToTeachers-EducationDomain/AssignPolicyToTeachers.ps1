# This Script will search for All Teachers distribution list in tenant and assign policy types TeamsChannelsPolicy, TeamsMeetingPolicy, TeamsMessagingPolicy with policy name AllTeachers to All Teachers distribution list

$logfile = ".\log_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
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
 }}
 
 If(Get-Module -ListAvailable -Name AzureAD) 
 { 
 Write-Host "AzureAD Already Installed" 
 } 
 else { 
 try { Install-Module -Name AzureAD
 Write-Host "Installed AzureAD"
 }
 catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
 }}
 
 try{
 $credential= get-credential
 }catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
try{
Connect-MicrosoftTeams -Credential $credential
}catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
try{
Connect-AzureAD -Credential $credential
}
catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}


$group = Get-AzureADGroup -SearchString "all teachers"
$groupid = $group.objectid
try{
New-CsGroupPolicyAssignment -GroupId $groupid -PolicyType TeamsChannelsPolicy -PolicyName "AllTeachers" -Rank 1
New-CsGroupPolicyAssignment -GroupId $groupid -PolicyType TeamsMeetingPolicy -PolicyName "AllTeachers" -Rank 1
New-CsGroupPolicyAssignment -GroupId $groupid -PolicyType TeamsMessagingPolicy -PolicyName "AllTeachers" -Rank 1
}
catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}


$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
