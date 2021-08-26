#This script will update All teachers distributation list based on teacher license 
$logfile = ".\TeachersGroupUpdate-Education domainlog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now
if(Get-Module -ListAvailable -Name AzureAD) 
 { 
 Write-Host "AzureAD Already Installed" 
 } 
 else { 
 try { 
  Write-Host "AzureAD is installing"
  Install-Module -Name AzureAD
 }
 catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
 }
 }
$credential = get-credential
Connect-AzureAD -Credential $credential
$group = Get-AzureADGroup -SearchString "all teachers"
$groupid = $group.objectid
$user = Get-AzureADUser -All $true  | where {$_.AssignedLicenses  -like "*94763226-9b3c-4e75-a931-5c89701abe66*"}
$userid = $user.objectid 
$memeber = get-AzureADGroupMember -All $true  -ObjectId  $groupid
$memberid = $memeber.objectid
foreach ($userid in $userid)
{
   if ($userid -notin $memberid)
    {
       try{
        Add-AzureADGroupMember -ObjectId $groupid -RefObjectId $userid
        }
        catch{
             $_.Exception.Message | out-file -Filepath $logfile -append
             }
    }
}
Get-AzureADSubscribedSku | ft  *skupart*,*consu*
(Get-AzureADGroupMember -all $true  -ObjectId $groupid).count
$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
#end of script
