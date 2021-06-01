# This script will filter microsoft teams user activity for period value: D7, D30, D90, and D180
param(
      [Parameter(Mandatory=$true)][System.String]$client_Id,
      [Parameter(Mandatory=$true)][System.String]$Client_Secret,
      [Parameter(Mandatory=$true)][System.String]$Tenantid, 
      [Parameter(Mandatory=$true)][System.String]$period
     )
$logfile = ".\Teamsuseractivitylog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now

#Grant Adminconsent 
$Grant= 'https://login.microsoftonline.com/common/adminconsent?client_id='
$admin = '&state=12345&redirect_uri=https://localhost:1234'
$Grantadmin = $Grant + $client_Id + $admin

Start-Process $Grantadmin
write-host "login with your tenant login detials to proceed further"

$proceed = Read-host " Press Y to continue "
if ($proceed -eq 'Y')
{
    write-host "Creating Access_Token"          
              $ReqTokenBody = @{
         Grant_Type    =  "client_credentials"
        client_Id     = "$client_Id"
        Client_Secret = "$Client_Secret"
        Scope         = "https://graph.microsoft.com/.default"
    } 
     
    $loginurl = "https://login.microsoftonline.com/" + "$Tenantid" + "/oauth2/v2.0/token"
    try{
    $Token = Invoke-RestMethod -Uri "$loginurl" -Method POST -Body $ReqTokenBody -ContentType "application/x-www-form-urlencoded"
    }
    catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
         } 

    $Header = @{
        Authorization = "$($token.token_type) $($token.access_token)"
    }
  function Get-Result() {
        write-host   "1-getTeamsUserActivityUserDetail
                      2-getTeamsUserActivityCounts
                      3-getTeamsUserActivityUserCounts"
                               
    $proceed = Read-host "Please provide number which Microsoft Teams user activity reports you want to proceed further"  

    if ($proceed -eq '1')
    {
         $UserActivity="https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='$period')"
         $UserActivityResults = Invoke-RestMethod -Headers $Header -Uri  $UserActivity -Method get -ContentType 'application/json'
         $resultarray = ConvertFrom-Csv -InputObject $UserActivityResults
         $resultarray | Export-Csv "getTeamsUserActivityUserDetail1.csv" -NoTypeInformation    
         }

             
elseif ($proceed -eq '2') 
{
            $UserActivity="https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityCounts(period='$period')"
            $UserActivityResults = Invoke-RestMethod -Headers $Header -Uri  $UserActivity -Method get -ContentType 'application/json'
            $resultarray = ConvertFrom-Csv -InputObject $UserActivityResults
            $resultarray | Export-Csv "getTeamsUserActivityCounts2.csv" -NoTypeInformation    

  }

elseif($proceed -eq '3')   
      {
         $UserActivity="https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserCounts(period='$period')"
         $UserActivityResults = Invoke-RestMethod -Headers $Header -Uri  $UserActivity -Method get -ContentType 'application/json'
         $resultarray = ConvertFrom-Csv -InputObject $UserActivityResults
         $resultarray | Export-Csv "getTeamsUserActivityUserCounts3.csv" -NoTypeInformation 
            
         }
   }
     
 do
{

$ProceedNext = Read-host "To Generate TeamUserActivity report, enter Y to continue"
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
#end of script
