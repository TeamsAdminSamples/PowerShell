# This script will add App to team using Teams powershell module cmdlets
param(
[Parameter(Mandatory=$true)][System.String]$AppName,
[Parameter(Mandatory=$true)][System.String]$TeamId
)
$logfile = ".\TeamsAppInstallationlog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now

If(Get-Module -ListAvailable -Name MicrosoftTeams) 
 { 
 Write-Host "MicrosoftTeams Already Installed" 
 } 
 else { 
 try {
     Install-Module -Name MicrosoftTeams
     Write-Host "Installed MicrosoftTeams"
     }
 catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
     }
  }
 
     try{
        Connect-MicrosoftTeams
        }
     catch{
            $_.Exception.Message | out-file -Filepath $logfile -append
         }

try{
$App = Get-TeamsApp -DisplayName "$AppName"
$APP
}
catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}

$AppId =$App.Id

try{
Add-TeamsAppInstallation -AppId $AppId -TeamId $TeamId
}
catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
