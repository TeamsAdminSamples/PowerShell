$logfile = "C:\TeamsAppInstallationlog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now

Connect-MicrosoftTeams
$AppName=Read-Host "Please enter AppName"
try{
$App=Get-TeamsApp -DisplayName "$AppName"
$APP
}
catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}

$AppId =$App.Id
$TeamId= Read-Host "Please enter TeamId"
try{
Add-TeamsAppInstallation -AppId $AppId -TeamId $TeamId
}
catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan