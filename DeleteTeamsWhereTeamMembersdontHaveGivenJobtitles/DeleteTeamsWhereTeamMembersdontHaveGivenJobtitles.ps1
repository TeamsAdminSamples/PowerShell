#this script will Delete the all teams except given job titles match
#Provide job title phrases separate by , 
#for best practice run the script quote delete cmdlet

$logfile = "C:\log_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now

     $Tenantid=read-host "Please provide tenant id"
     $client_Id=Read-host "Please provide client id"
     $Client_Secret=read-host "Please provide client secret"
     $mailsender = read-host "Please provide mailsender"
     $KeepJobtitles = read-host "Please provide KeepJobtitles"

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
      Catch {
    $_.Exception | Out-File $logfile -Append
      }
    $Header = @{
        Authorization = "$($token.token_type) $($token.access_token)"
    }
    
   #getting All teams
   write-host "Getting All teams"

   $getTeams = "https://graph.microsoft.com/beta/groups?filter=resourceProvisioningOptions/Any(x:x eq 'Team')" 
   try{
   $Teams = Invoke-RestMethod -Headers $Header -Uri $getTeams -Method get -ContentType 'application/json'
   }
   Catch {
    $_.Exception | Out-File $logfile -Append
   }
   $values = $Teams.value
   $groupid = $values.id
   $displayname = $values | select displayName
   
   #getting members
   write-host "Getting members for each team"
   $results = foreach($team in $values)
   {
           $id = $team.id
           $TeamName = $team.displayName
           write-host "Checking job title of users in $TeamName"
            
            $memberuri = "https://graph.microsoft.com/v1.0/groups/"+ "$id" +"/members"
            try{
            $members = Invoke-RestMethod -Headers $Header -Uri $memberuri -Method get -ContentType 'application/json'
            }
            Catch {
             $_.Exception | Out-File $logfile -Append
            }
            # for each member - check the designation
            $keepTeam = $false
                        
            foreach($member in $members.value)
            {
                if( "$KeepJobtitles"  -contains $member.jobTitle)
                    {
                        $keepTeam = $true
                    }
            }
            # delete if flag is false
            if(!$keepTeam)
             {      
                    $DeletedTeam = $team | select displayName
                    $deleteURL = "https://graph.microsoft.com/v1.0/groups/" + "$id" 
                    try{
                    #$DeleteTeam = Invoke-RestMethod -Headers $Header -Uri $deleteURL -Method DELETE 
                    }
                    Catch {
                      $_.Exception | Out-File $logfile -Append
                     }
                    write-host "$Teamname has been deleted"
                    
                    $owneruri = "https://graph.microsoft.com/v1.0/groups/" + "$id" + "/owners"
                    try{
                    $Teamowners = Invoke-RestMethod -Headers $Header -Uri $owneruri -Method Get
                    }
                    Catch {
                      $_.Exception | Out-File $logfile -Append
                     }
                    $Teamownervalues = $Teamowners.value 
                    $OwneruserPrincipalName = $Teamownervalues.userPrincipalName
                    $owners = [string]::Join(", ",$OwneruserPrincipalName) 

                    $file = New-Object psobject
                    $file | add-member -MemberType NoteProperty -Name DeletedTeam $DeletedTeam.displayName
                    $file | add-member -MemberType NoteProperty -Name TeamsOwner $owners
                    $file | export-csv output.csv -NoTypeInformation -Append
                 write-host "Mail has been sent to $owners"
                 $mailuri = "https://graph.microsoft.com/v1.0/users/" + "$mailsender" + "/sendMail"

                    $body = '{
  "message": {
    "subject": "Your team ' +$Teamname+ ' has been deleted because of complaince reason",
    "body": {
      "contentType": "Text",
      "content": "Your team ' +$Teamname+ ' has been deleted because of complaince reason"
    },
    "toRecipients": [
      {
        "emailAddress": {
          "address": "' +$owners+ '"
        }
      }
    ],
    "ccRecipients": [
      {
        "emailAddress": {
          "address": "' +$mailsender+ '"
        }
      }
    ]
  },
  "saveToSentItems": "True"
}'
try{
            $smtp = Invoke-RestMethod -Headers $Header -Uri $mailuri -body $body -Method post -ContentType application/json
            }
            Catch {
    $_.Exception | Out-File $logfile -Append
   }
             }
   }
   
 } 
   
else{ write-host Re run the script and press Y}
$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan

