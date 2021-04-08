param(
      [Parameter(Mandatory=$true)][System.String]$OwnerPrincipalName,
      [Parameter(Mandatory=$true)][System.String]$AppName,
      [Parameter(Mandatory=$true)][System.String]$Tenantid,
      [Parameter(Mandatory=$true)][System.String]$client_Id,
      [Parameter(Mandatory=$true)][System.String]$Client_Secret
      )
      
Connect-MicrosoftTeams

#Grant Adminconsent 
$Grant= 'https://login.microsoftonline.com/common/adminconsent?client_id='
$admin = '&state=12345&redirect_uri=https://localhost:1234'
$Grantadmin = $Grant + $client_Id + $admin

start $Grantadmin
write-host "login with your tenant login detials to proceed further"

$proceed = Read-host "Press 1 for Add app to Team Press 2 for Add app to Tenant"
if ($proceed -eq '1')
{
        write-host "Creating Access_Token"          
                  $ReqTokenBody = @{
             Grant_Type    =  "client_credentials"
            client_Id     = "$client_Id"
            Client_Secret = "$Client_Secret"
            Scope         = "https://graph.microsoft.com/.default"
        } 

        $loginurl = "https://login.microsoftonline.com/" + "$Tenantid" + "/oauth2/v2.0/token"
        $Token = Invoke-RestMethod -Uri "$loginurl" -Method POST -Body $ReqTokenBody -ContentType "application/x-www-form-urlencoded"

        $Header = @{
            Authorization = "$($token.token_type) $($token.access_token)"
        }


#add app to team


$getappuri = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?filter=name%20eq%20'$AppName'"
$getapp = Invoke-RestMethod -Headers $Header -Uri $ownerurl  -Method get -ContentType 'application/json'
$Appid = $getapp.id

write-host "Adding App to Team"
$id = Read-host "Please provide team name"
$Appbody = '{
   "teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/'+$Appid+'"
}'
    $AddAppsuri = "https://graph.microsoft.com/v1.0/teams/" +$id+ "/installedApps"
    $Apps = Invoke-RestMethod -Headers $Header -Uri $AddAppsuri -body $Appbody -Method post -ContentType 'application/json'
    write-host "app has been installed to " $id
}

if ($proceed -eq '2')
{
        write-host "Creating Access_Token"          
                  $ReqTokenBody = @{
             Grant_Type    =  "client_credentials"
            client_Id     = "$client_Id"
            Client_Secret = "$Client_Secret"
            Scope         = "https://graph.microsoft.com/.default"
        } 

        $loginurl = "https://login.microsoftonline.com/" + "$Tenantid" + "/oauth2/v2.0/token"
        $Token = Invoke-RestMethod -Uri "$loginurl" -Method POST -Body $ReqTokenBody -ContentType "application/x-www-form-urlencoded"

        $Header = @{
            Authorization = "$($token.token_type) $($token.access_token)"
        }

        $Teams = get-team 
foreach ($team in $Teams) 
    {

        $groupid = $team.Groupid
        $displayname = $team.DisplayName
          
            #add app to team

        $getappuri = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?filter=name%20eq%20'$AppName'"
        $getapp = Invoke-RestMethod -Headers $Header -Uri $ownerurl  -Method get -ContentType 'application/json'
        $Appid = $getapp.id

write-host "Adding App to Team"
$Appbody = '{
   "teamsApp@odata.bind":"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/'+$Appid+'"
}'
    $AddAppsuri = "https://graph.microsoft.com/v1.0/teams/" +$groupid+ "/installedApps"
    $Apps = Invoke-RestMethod -Headers $Header -Uri $AddAppsuri -body $Appbody -Method post -ContentType 'application/json'
    write-host $displayname "has been installed" 

}
            
    
    }}
