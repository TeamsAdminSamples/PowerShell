# This script will create Team along with apps using graph api calling method.

param(
      [Parameter(Mandatory=$true)][System.String]$OwnerPrincipalName,
      [Parameter(Mandatory=$true)][System.String]$AppName,
      [Parameter(Mandatory=$true)][System.String]$Groupname,
      [Parameter(Mandatory=$true)][System.String]$Tenantid,
      [Parameter(Mandatory=$true)][System.String]$client_Id,
      [Parameter(Mandatory=$true)][System.String]$Client_Secret
      )
      
$logfile = ".\log_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now
           
#Grant Adminconsent 
$Grant= 'https://login.microsoftonline.com/common/adminconsent?client_id='
$admin = '&state=12345&redirect_uri=https://localhost:1234'
$Grantadmin = $Grant + $client_Id + $admin

start-process $Grantadmin
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
      Catch {
     $_.Exception | Out-File $logfile -Append
      }
    $Header = @{
        Authorization = "$($token.token_type) $($token.access_token)"
    }

#create Group with owner
write-host "Creating Group with owner"
$ownerurl = "https://graph.microsoft.com/v1.0/users/" + "$OwnerPrincipalName"
try{
$owner = Invoke-RestMethod -Headers $Header -Uri $ownerurl  -Method get -ContentType 'application/json'
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
$ownerid = $owner.id
$groupbody = '{
  "description": "Group with designated owner and members",
  "displayName": "Operations group",
  "groupTypes": [
    "Unified"
  ],
  "mailEnabled": true,
  "mailNickname": "operations2019",
  "securityEnabled": false,
  "owners@odata.bind": [
    "https://graph.microsoft.com/v1.0/users/'+$ownerid+'"
  ]
}'


    $groupurl = "https://graph.microsoft.com/v1.0/groups" 
    try{
    $Group = Invoke-RestMethod -Headers $Header -Uri $groupurl -body $groupbody -Method post -ContentType 'application/json'
    }
    Catch {
    $_.Exception | Out-File $logfile -Append
   }
    $id = $group.id

#convert that Group to team
write-host "Creating Team"
$teambody = '{  
  "memberSettings": {
    "allowCreateUpdateChannels": true
  },
  "messagingSettings": {
    "allowUserEditMessages": true,
    "allowUserDeleteMessages": true
  },
  "funSettings": {
    "allowGiphy": true,
    "giphyContentRating": "strict"
  }
}'
    $teamuri = "https://graph.microsoft.com/v1.0/groups/" +$id+ "/team" 
    try{
    $Team = Invoke-RestMethod -Headers $Header -Uri $teamuri -body $teambody -Method put -ContentType 'application/json'
    }
    Catch {
    $_.Exception | Out-File $logfile -Append
    }

#add app to team

$getappuri = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?filter=name%20eq%20'$AppName'"
try{
$getapp = Invoke-RestMethod -Headers $Header -Uri $getappuri  -Method get -ContentType 'application/json'
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
$Appid = $getapp.id

write-host "Adding App to Team"
$Appbody = '{
   "teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/'+$Appid+'"
}'
    $AddAppsuri = "https://graph.microsoft.com/v1.0/teams/" +$id+ "/installedApps"
    try{
    Invoke-RestMethod -Headers $Header -Uri $AddAppsuri -body $Appbody -Method post -ContentType 'application/json'
    }
    Catch {
    $_.Exception | Out-File $logfile -Append
   }
write-host "App has been added to team"
}

$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
