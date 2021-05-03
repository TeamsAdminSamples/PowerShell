#This script will assign Teams messaging policy to user.
param(
      [Parameter(Mandatory=$true)][System.String]$User,
      [Parameter(Mandatory=$true)][System.String]$PolicyName         
      )
$start = [system.datetime]::Now
$logfile = ".\Log_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
      try{

#connecting to Skypeonline
 $credential = Get-credential
 Import-Module SkypeOnlineConnector
 $sfbSession = New-CsOnlineSession
 Import-PSSession $sfbSession

Grant-CsTeamsMessagingPolicy -Identity "$User" -PolicyName "$PolicyName"
write-host "$User is being assigned the $PolicyName"
}
catch
{
$_.Exception.Message | out-file -Filepath $logfile -append
}

$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds."

