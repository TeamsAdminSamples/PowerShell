# This script will create and assign addressbook policy
$logfile = ".\CreateAndAssignAddressbookPolicylog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now
if(Get-Module -ListAvailable -Name ExchangeOnlineManagement) 
 { 
 Write-Host "ExchangeOnlineManagement Already Installed" 
 } 
 else { 
 try { 
Install-Module -Name ExchangeOnlineManagement
 Write-Host "ExchangeOnlineManagement is installing"
 }
 catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
 }
 }
 #Creating and assigning address policy
try{
$cred = Connect-ExchangeOnline
$user=  Get-Mailbox -ResultSize unlimited
$school = $user.CustomAttribute10 |  Sort-Object | Get-Unique
$school=$school.Where({ $_ -ne "" })
$school=$school.Trim()
foreach ($school in $school)
{
New-AddressList -Name $school   -ConditionalCustomAttribute10 $school -IncludedRecipients "AllRecipients"
$addresslist = (Get-AddressList $school).name+' GAL'
New-GlobalAddressList -Name "$addresslist" -ConditionalCustomAttribute10  $school -IncludedRecipients "AllRecipients"
$GAL=  $school+' GAL'
New-OfflineAddressBook -Name $school -AddressLists "\$gal"
New-AddressBookPolicy -Name $school -AddressLists "\$school" -RoomList "\All Rooms" -OfflineAddressBook "$school" -GlobalAddressList "$gal"
Get-Mailbox | where{$_.customattribute10 -like "*$school*"} | Set-Mailbox -AddressBookPolicy $school
}
}
catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
#end of script
