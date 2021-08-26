#This script will check the user license, based on the assigned license(Teacher/Student), a policy package(Education Teacher/Education_SecondaryStudent) will be assigned

param(    
      [Parameter(Mandatory=$true)][System.String]$client_Id,
      [Parameter(Mandatory=$true)][System.String]$Client_Secret,
      [Parameter(Mandatory=$true)][System.String]$Tenantid    
      )
$logfile = ".\PolicyPackageAssignmentlog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now
#connect to teams
try{
Connect-MicrosoftTeams
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
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

    $uri = "https://graph.microsoft.com/v1.0/users"
    try{
    $group = Invoke-RestMethod -Headers $Header -Uri $uri  -Method Get
      }
      Catch {
    $_.Exception | Out-File $logfile -Append
   }
    do
    { 
        foreach($value in $group.value)
       { 
       try{
       $Token1 = Invoke-RestMethod -Uri "$loginurl" -Method POST -Body $ReqTokenBody -ContentType "application/x-www-form-urlencoded"
            }
            Catch {
                $_.Exception | Out-File $logfile -Append
               }
    $Header1 = @{
        Authorization = "$($token1.token_type) $($token1.access_token)"
    }

    $id = $value.id
    $UPN = $value.userPrincipalName

    #Check if user is assigned any license
        $licenseuri = "https://graph.microsoft.com/v1.0/users/" + "$id" + "/licenseDetails"
        try{
        $licenseresult = Invoke-RestMethod -Headers $Header1 -Uri $licenseuri  -Method Get
        }
        Catch {
                $_.Exception | Out-File $logfile -Append
               }
        $licensevalue = $licenseresult.value
        $skuids = $licensevalue.skuId
        $licenses = $licensevalue.skuPartNumber
        $serviceplan = $licensevalue.servicePlans
        $TeamslicenseStatus = $serviceplan | where {($_.servicePlanName -eq 'Teams1')}

        $provisioningStatus = $TeamslicenseStatus.provisioningStatus
        
        $useruri = "https://graph.microsoft.com/v1.0/users/" + $id
        try{
        $userresult = Invoke-RestMethod -Headers $Header1 -Uri $useruri  -Method Get
          }
          Catch {
                $_.Exception | Out-File $logfile -Append
               }
         if($licenses -contains "M365EDU_A5_FACULTY")
         {
         try{
         Grant-CsUserPolicyPackage -Identity $UPN -PackageName "Education_Teacher"
         }
         Catch {
                $_.Exception | Out-File $logfile -Append
               }
         write-host " Education_Teacher policy has been assigned to user" $UPN
         }
         elseif($licenses -contains "M365EDU_A5_STUDENT")
         {
         try{
         Grant-CsUserPolicyPackage -Identity $UPN -PackageName "Education_SecondaryStudent"
         }
         Catch {
                $_.Exception | Out-File $logfile -Append
               }
         write-host " Education_SecondaryStudent policy has been assigned to user" $UPN
         }

         else{
         Write-Host "User have the diffrent license" $UPN
            $file = New-Object psobject
            $file | add-member -MemberType NoteProperty -Name UserName $UPN
            $file | add-member -MemberType NoteProperty -Name Userid $id
            $file | export-csv -path ".\license.csv" -NoTypeInformation -Append
         }
          }
   if ($group.'@odata.nextLink' -eq $null ) 
        { 
        break 
        } 
        else 
        {
        try{
        $Token2 = Invoke-RestMethod -Uri "$loginurl" -Method POST -Body $ReqTokenBody -ContentType "application/x-www-form-urlencoded"
            }
            Catch {
                $_.Exception | Out-File $logfile -Append
               }
    $Header2 = @{
        Authorization = "$($token2.token_type) $($token2.access_token)"
    }
     try{
        $group = Invoke-RestMethod -Headers $Header2 -Uri $group.'@odata.nextLink' -Method Get 
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
