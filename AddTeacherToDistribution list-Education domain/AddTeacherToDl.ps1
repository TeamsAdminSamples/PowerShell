#Script will search and filter teachers in a tenant using license parameter and adds to the All teachers distribution list
$logfile = ".\log_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now

If(Get-Module -ListAvailable -NameAzureAD) 
 { 
 Write-Host "AzureAD Already Installed" 
 } 
 else { 
 try { Install-Module -Name AzureAD
 Write-Host "Installed AzureAD"
 }
 catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
 }
 }
 try{
 $credential= get-credential
 Connect-AzureAD -Credential $credential
 }
 catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
$group = Get-AzureADGroup -SearchString "all teachers"
$groupid = $group.objectid
try{
$user = Get-AzureADUser -All $true  | where {$_.AssignedLicenses  -like "*94763226-9b3c-4e75-a931-5c89701abe66*"}
$userid = $user.objectid
foreach ($userid in $userid){
Add-AzureADGroupMember -ObjectId $groupid -RefObjectId $userid
}
Get-AzureADSubscribedSku | ft  *skupart*,*consu*
(Get-AzureADGroupMember -all $true  -ObjectId $groupid).count
}
 catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}

$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan

#end of script
