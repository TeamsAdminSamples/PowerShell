# This script will send notification to selected Audience using Graph api calling method.
param(    
      [Parameter(Mandatory=$true)][System.String]$client_Id,
      [Parameter(Mandatory=$true)][System.String]$Client_Secret,
      [Parameter(Mandatory=$true)][System.String]$Tenantid     
      )
$logfile = ".\SendNotificationlog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
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

    $inputmethod = Read-host "Please choose target audience number 
    1.send notification to user chat
    2.send notification to specific Team
    3.send notification to user installed app
    4.send notification to Distribution List
    5.send notification to Bulk users for installed app(csv)
    6.send notification to specific tab in application"

  if($inputmethod -eq 1){
    $chatid1 = read-host "please provide chatid"
    $userid1 = read-host "please provide userid"
    
    $uri1 = "https://graph.microsoft.com/beta/chats/" +$chatid1 + "/sendActivityNotification"
$body1 = '{
            "topic": {
                "source": "entityUrl",
                "value": "https://graph.microsoft.com/beta/chats/'+$chatid1+'"
            },
            "activityType": "taskCreated",
            "previewText": {
                "content": "New Task Created"
            },
            "recipient": {
                "@odata.type": "microsoft.graph.aadUserNotificationRecipient",
                "userId": "'+$userid1+'"
            },
            "templateParameters": [
                {
                    "name": "taskId",
                    "value": "12322"
                }
            ]
          }'

    try{
    $SendNotificationCall1 = Invoke-RestMethod -Uri $uri1 -Headers $Header -Body $body1 -Method Post -ContentType "application/json"
      }
      catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
         } 
    }
  
if($inputmethod -eq 2){

$Teamid2 = read-host "please provide teamid"
    $userid2 = read-host "please provide userid"

$uri2 = "https://graph.microsoft.com/beta/teams/" +$Teamid2 + "/sendActivityNotification"
$body2 ='
{
    "topic": {
        "source": "entityUrl",
        "value": "https://graph.microsoft.com/beta/teams/'+$Teamid2+'"
    },
    "activityType": "taskCreated",
    "previewText": {
    	"content": "New Task Created"
    },
    "recipient": {
        "@odata.type": "microsoft.graph.aadUserNotificationRecipient",
        "userId": "'+$userid2+'"
    },
    "templateParameters": [
        {
            "name": "taskId",
            "value": "12322"
        }
    ]
}'
try{
$SendNotificationCall2 = Invoke-RestMethod -Uri $uri2 -Headers $Header -Body $body2 -Method Post -ContentType "application/json"
}
catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
         } 

}

if($inputmethod -eq 3){
$userid3 = read-host "please provide userid"
$appid3 = read-host "Please provide application id"
$uri3 = "https://graph.microsoft.com/beta/users/"+"$userid3"+"/teamwork/sendActivityNotification"
$body3 ='
{
    "topic": {
        "source": "entityUrl",
        "value": "https://graph.microsoft.com/beta/users/'+"$userid3"+'/teamwork/installedApps/'+$appid3+'"
    },
    "activityType": "taskCreated",
    "previewText": {
        "content": "New Task Created"
    },
    "templateParameters": [
        {
            "name": "taskId",
            "value": "Task 12322"
        }
    ]
}'
try{
$SendNotificationCall3 = Invoke-RestMethod -Uri $uri3 -Headers $Header -Body $body3 -Method Post -ContentType "application/json"
}
catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
         } 

}

if($inputmethod -eq 4){
$userid4 
$appid4 = read-host "Please provide application id"
$DL = read-host "please provide distribution list id"
$DLuri = "https://graph.microsoft.com/v1.0/groups/"+$DL+"/members"
#$DLuri = 'https://graph.microsoft.com/v1.0/groups/?$filter=mail'+ "eq" +'email@domain.onmicrosoft.com'+'&$expand=members$select=members/displayName'
try{
$DLmembers = Invoke-RestMethod -Uri $DLuri -Headers $Header -Method Get -ContentType "application/json"
}
catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
        } 
$userdetails = $DLmembers.value
$userids = $userdetails.id
foreach($userid4 in $userids){
$uri4 = "https://graph.microsoft.com/beta/users/"+"$userid4"+"/teamwork/sendActivityNotification"
$body4 ='
{
    "topic": {
        "source": "entityUrl",
        "value": "https://graph.microsoft.com/beta/users/'+"$userid4"+'/teamwork/installedApps/'+$appid4+'"
    },
    "activityType": "taskCreated",
    "previewText": {
        "content": "New Task Created"
    },
    "templateParameters": [
        {
            "name": "taskId",
            "value": "Task 12322"
        }
    ]
}'
try{
$SendNotificationCall4 = Invoke-RestMethod -Uri $uri4 -Headers $Header -Body $body4 -Method Post -ContentType "application/json"
}
catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
      } 
}
}
if($inputmethod -eq 5){
$appid5 = read-host "Please provide application id"
 $file_List = Read-host "Please provide csv file full location(ex:d:\example.csv)"
    Foreach ($file in $file_List){
    $userid5 = $file.user
    $uri5 = "https://graph.microsoft.com/beta/users/"+"$userid3"+"/teamwork/sendActivityNotification"
    $body5 ='
{
    "topic": {
        "source": "entityUrl",
                "value": "https://graph.microsoft.com/beta/users/'+"$userid5"+'/teamwork/installedApps/'+$appid5+'"

    },
    "activityType": "taskCreated",
    "previewText": {
        "content": "New Task Created"
    },
    "templateParameters": [
        {
            "name": "taskId",
            "value": "Task 12322"
        }
    ]
}'
try{
$SendNotificationCall5 = Invoke-RestMethod -Uri $uri5 -Headers $Header -Body $body5 -Method Post -ContentType "application/json"
}
catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
         } 

}

}

 if($inputmethod -eq 6){

 $userid6 = read-host = "please provide user id"
 $appid6 = read-host = "please provide app id for application"
 $tabid6 = read-host = "please provide tab id "

 $uri6 = "https://graph.microsoft.com/beta/users/" + "$userid6" + "/teamwork/sendActivityNotification"
$body6 ='
{
    "topic": {
        "source": "text",
        "value": "Check Out Playlists!",
        "webUrl": "https://teams.microsoft.com/l/entity/'+$appid6 + "/" + $tabid6+'
    },
    "activityType": "taskCreated",
    "previewText": {
        "content": "New Task Created"
    },
    "templateParameters": [
        {
            "name": "taskId",
            "value": "Task 12322"
        }
    ]
}'


try{
 $SendNotificationCall6 = Invoke-RestMethod -Uri $uri6 -Headers $Header -Body $body6 -Method Post -ContentType "application/json"
}
catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
         } 
}

}

else{write-host "please rerun the script login with created graphapi application credentials"}

$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan

#end of script
