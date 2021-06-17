#This script will assign custom Teams app setup policy to the user using PowerShell cmdlets
param(
[Parameter(Mandatory=$true)][System.String]$Policyname,
[Parameter(Mandatory=$true)][System.String]$user
)
$logfile = ".\TeamsAppSetupPolicySingleUserlog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now 
If(Get-Module -ListAvailable -Name SkypeOnlineConnector) 
 { 
 Write-Host "SkypeOnlineConnector Already Installed" 
 } 
 else { 
    try {
         Write-Host "Installing SkypeOnlineConnector" 
         Install-Module -Name SkypeOnlineConnector
        }
 catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
 }
 }
Import-Module SkypeOnlineConnector
$sfbSession = New-CsOnlineSession 
Import-PSSession $sfbSession -AllowClobber
try{
Grant-CsTeamsAppsetupPolicy -policyname "$Policyname" -Identity  $user
get-csonlineuser -Identity "$user" |ft TeamsappsetupPolicy,UserPrincipalName
}
catch{
     $_.Exception.Message | out-file -Filepath $logfile -append
     }

$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
#end of script
