#This script will restrict the Teams storage. 
#"Please provide SiteName ex:https://contoso.sharepoint.com/sites/HrTeam"
# "Please provide StorageQuota"
# "Please provide StorageQuotaWarningLevel"  

param(
      [Parameter(Mandatory=$true)][System.String]$SiteName,
      [Parameter(Mandatory=$true)][System.String]$StorageQuota,
      [Parameter(Mandatory=$true)][System.String]$StorageQuotaWarningLevel
      )

$logfile = ".\SharePointstoragelimitlog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now
# site name = https://contoso.sharepoint.com/sites/HrTeam
# StorageQuota = 20000 (input is in GB)
# StorageQuotaWarningLevel = 19000 (input is in GB)]
      
try{
Set-SPOSite -Identity $SiteName -StorageQuota $StorageQuota -StorageQuotaWarningLevel $StorageQuotaWarningLevel -NoWait
}
catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
