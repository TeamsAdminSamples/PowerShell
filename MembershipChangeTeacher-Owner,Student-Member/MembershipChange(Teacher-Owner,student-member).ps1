#This script will change the Teams membership role(owner/member) based on the user license(Teacher/student)
#If users don't have a MicrosoftTeams license it will export the user's list
param(    
      [Parameter(Mandatory=$true)][System.String]$client_Id,
      [Parameter(Mandatory=$true)][System.String]$Client_Secret,
      [Parameter(Mandatory=$true)][System.String]$Tenantid   
      )
$logfile = ".\Membershiplog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now 

#Grant Adminconsent 
$Grant= 'https://login.microsoftonline.com/common/adminconsent?client_id='
$admin = '&state=12345&redirect_uri=https://localhost:1234'
$Grantadmin = $Grant + $client_Id + $admin

try{
Connect-MicrosoftTeams
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
start-process $Grantadmin
write-host "login with your tenant login details to proceed further"

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

 
    #Get Team details
         write-host "Getting Team details..."
         $getTeams = "https://graph.microsoft.com/beta/groups?filter=resourceProvisioningOptions/Any(x:x eq 'Team')" 
         try{
         $Teams = Invoke-RestMethod -Headers $Header -Uri $getTeams -Method get -ContentType 'application/json'
         }
         Catch {
          $_.Exception | Out-File $logfile -Append
         }
         do 
        {

      foreach($Team in $Teams.value.id)
      {
       $Tmembers ="https://graph.microsoft.com/v1.0/groups/" + $Team + "/members"
       try{
        $members = Invoke-RestMethod -Headers $Header -Uri $Tmembers -Method get 
            }
            Catch {
                $_.Exception | Out-File $logfile -Append
               }
       #Get all team Owners
        $Teamowneruri ="https://graph.microsoft.com/v1.0/groups/" + $Team + "/owners"
        try{
        $ownerresult = Invoke-RestMethod -Headers $Header -Uri $Teamowneruri -Method get
        }
        Catch {
                $_.Exception | Out-File $logfile -Append
               }
        $owners = $ownerresult.value.id 

        foreach($value in  $members.value)
         {
         $member = $value.id
        $memberUPN = $value.userPrincipalName
        $memberdisplayname = $value.displayName

       $licenseuri="https://graph.microsoft.com/v1.0/users/" + $member + "/licenseDetails"
       try{
       $licenseresult=Invoke-RestMethod -Headers $Header -Uri $licenseuri -Method get
            }
            Catch {
                $_.Exception | Out-File $logfile -Append
               }
        $licensevalue = $licenseresult.value
        $license = $licensevalue.skuPartNumber
        
        
    #case1:if user having faculty license and be part of ownerlist
            if(($license -eq "M365EDU_A5_FACULTY") -and ($owners -contains $member))
                        {write-host "This user having Faculty license and already owner of the team" $memberdisplayname }

    #case2:if user having faculty license and not part of ownerlist
            elseif(($license -eq "M365EDU_A5_FACULTY") -and ($owners -notcontains $member))
                     { 
                      $facultybody='{
                            "@odata.id": "https://graph.microsoft.com/beta/users/'+$member+'"
                            }'
                            $facultyuri ="https://graph.microsoft.com/beta/groups/" + "$Team" + "/owners/`$ref"
                            try{
                            $output =Invoke-RestMethod -Headers $Header -Uri $facultyuri -Method Post -Body $facultybody -ContentType 'application/json'
                            }
                            Catch {
                                  $_.Exception | Out-File $logfile -Append
                                 }
                            write-host "Faculty Membership role has been changed to Owner for team"   $memberdisplayname 
                            
                            }
    #case3:if user having STUDENT license and not part of ownerlist
            elseif(($license -eq "M365EDU_A5_STUDENT") -and ($owners -notcontains $member))
                    {write-host "This user having STUDENT license and already member of the team" $memberdisplayname}

    #case4:if user having STUDENT license and part of ownerlist
            elseif(($license -eq "M365EDU_A5_STUDENT") -and ($owners -contains $member))
                    {
                     #add student as member
                     try{
                  Add-TeamUser -GroupId $Team -User $memberUPN -Role Member
                  }
                  Catch {
                      $_.Exception | Out-File $logfile -Append
                     }
                   #removing student as owner                                             
                $removestudenturi="https://graph.microsoft.com/v1.0/groups/" +$Team+ "/owners/" +$member+ "/`$ref"
                try{
                $output2=Invoke-RestMethod -Headers $Header -Uri $removestudenturi -Method Delete -ContentType 'application/json'
                }
                Catch {
                      $_.Exception | Out-File $logfile -Append
                     }
                write-host "student Membership role has been changed to member " $memberdisplayname
                
                                                                                  
                   }
    #case5: if user dont have license
      else
      {
            write-host "user have the different license" 
            $file = New-Object psobject
            $file | add-member -MemberType NoteProperty -Name UserName $memberUPN
            $file | add-member -MemberType NoteProperty -Name Userid $member
            $file | export-csv -path output.csv -NoTypeInformation -Append
        }
    }
    }
    
    
    
    if ($group.'@odata.nextLink' -eq $null ) 
        { 
        break 
        } 
        else 
        {
        try{
        $group = Invoke-RestMethod -Headers $Header -Uri $group.'@odata.nextLink' -Method Get 
        }
        Catch {
                $_.Exception | Out-File $logfile -Append
               }
        } 
        }while($true); 
        }
     
        else 
{
    write-host "You need to login admin consent in order to continue... " 
}
  $end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
#end of script
