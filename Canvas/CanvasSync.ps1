<#
This script will download school data sync files from canvas
 Use at your own risk and stuff
 This project contains the main methods for the Canvas APIs as well
 as a number of test methods. See the method generator for more potential.
 Based on https://canvas.instructure.com/doc/api/index.html
#>

$logfile = ".\CanvasSynclog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now
#region Base Canvas API Methods
function Get-CanvasCredentials()
{
    if ($global:CanvasApiTokenInfo -eq $null) 
    {
    
        $ApiInfoPath = "$env:USERPROFILE\Documents\CanvasApiCreds.json"

        #TODO: Once this is a module, load it from the module path: $PSScriptRoot or whatever that is
        if (-not (test-path $ApiInfoPath)) 
           {
            $Token = Read-Host "Please enter your Canvas API API Access Token"
            $BaseUri = Read-Host "Please enter your Canvas API Base URI (for example, https://domain.beta.instructure.com)"

            $ApiInfo = [ordered]@{
                Token = $Token
                BaseUri = $BaseUri
            }

            $ApiInfo | ConvertTo-Json | Out-File -FilePath $ApiInfoPath
     }

        #load the file
        $global:CanvasApiTokenInfo = Get-Content -Path $ApiInfoPath | ConvertFrom-Json
    }

    return $global:CanvasApiTokenInfo
}

function Get-CanvasAuthHeader($Token) {
    return @{"Authorization"="Bearer "+$Token}
}

function Get-CanvasApiResult(){

    Param(
        $Uri,

        $RequestParameters,

        [ValidateSet("GET", "POST", "PUT", "DELETE")]
        $Method="GET"
    )
    
    $AuthInfo = Get-CanvasCredentials

    if ($RequestParameters -eq $null) { $RequestParameters = @{} 

    $RequestParameters["per_page"] = "10000"

    $Headers = (Get-CanvasAuthHeader $AuthInfo.Token)

    try 
    {
    $Results = Invoke-WebRequest -Uri ($AuthInfo.BaseUri + $Uri) -ContentType "multipart/form-data" `
        -Headers $headers -Method $Method -Body $RequestParameters 
    } catch 
    {
        throw $_.Exception.Message
    }

    $Content = $Results.Content | ConvertFrom-Json

    #Either PSCustomObject or Object[]
    if ($Content.GetType().Name -eq "PSCustomObject") {
        return $Content
    }

    $JsonResults = New-Object System.Collections.ArrayList

    $JsonResults.AddRange(($Results.Content | ConvertFrom-Json))

    if ($Results.Headers.link -ne $null) {
        $NextUriLine = $Results.Headers.link.Split(",") | where {$_.Contains("rel=`"next`"")}

        $PerPage = $NextUriLine.Substring($NextUriLine.IndexOf("per_page=")) -replace '(\D).*',""

        if (-not [string]::IsNullOrWhiteSpace($NextUriLine)) 
        {
            while ($Results.Headers.link.Contains("rel=`"next`"")) 
{
        
                $nextUri = $Results.Headers.link.Split(",") | `
                            where {$_.Contains("rel=`"next`"")} | `
                            % {$_ -replace ">; rel=`"next`""} |
                            % {$_ -replace "<"}
        
                #Write-Progress
                Write-Host $nextUri
                try{
                $Results = Invoke-WebRequest -Uri $nextUri -Headers $headers -Method Get -Body $RequestParameters -ContentType "multipart/form-data" `
                }
                Catch {
                    $_.Exception | Out-File $logfile -Append
                   }
                $JsonResults.AddRange(($Results.Content | ConvertFrom-Json))
            }
        }
    }

    return $JsonResults
}}

write-host "creating school.csv file"
#accounts 
try{
$results = Get-CanvasApiResult -Uri "/api/v1/accounts" -Method GET
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
$results | convertto-Csv -NoTypeInformation
$results | Export-csv ".\Accounts.csv" -Append -NoTypeInformation

try{
$results = Get-CanvasApiResult -Uri "/api/v1/accounts/1/sub_accounts" -Method GET
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
$results | convertto-Csv -NoTypeInformation
$results | Export-csv -path  ".\Accounts.csv" -Append -NoTypeInformation

 $acc = import-csv -path .\Accounts.csv
 $xyz = $acc.id
 $count = $acc.Count
foreach($x in $xyz)
{
    $uri0 = "/api/v1/accounts/"+ "$x" +"/sub_accounts"
    try{
    $results = Get-CanvasApiResult -Uri $uri0 -Method GET
    }
    Catch {
    $_.Exception | Out-File $logfile -Append
   }
    $results | convertto-Csv -NoTypeInformation 
    $results | Export-csv ".\Accounts.csv" -Append -NoTypeInformation
}

$tempCSV = Import-Csv .\Accounts.csv -Header "SIS ID","Name","workflow_state","parent_account_id","root_account_id","uuid","default_storage_quota_mb","default_user_storage_quota_mb","default_group_storage_quota_mb","default_time_zone" | select -skip 1 | sort 'SIS ID','Name' -Unique
$tempCSV | Export-CSV .\school.csv -NoTypeInformation 


#courses details
    try{
    $results = Get-CanvasApiResult -Uri "/api/v1/accounts/1/courses" -Method GET
    }
    Catch {
    $_.Exception | Out-File $logfile -Append
   }
    $results | convertto-Csv -NoTypeInformation
    $results | Export-csv "courses.csv" -NoTypeInformation
    $courses = import-csv "courses.csv"
if (Test-Path Sync.csv)
{    
$Sync = import-csv ".\Sync.csv" -Header id |select -skip 1

$id = $Sync.id
write-host "creating Section.csv file"  
#Section Details
foreach($i in $id)
{
    $uri1 = "/api/v1/courses/" + "$i" + "/sections"
    try{
    $results = Get-CanvasApiResult -Uri $uri1 -Method GET
    }
    Catch {
    $_.Exception | Out-File $logfile -Append
   }
    $results | convertto-Csv -NoTypeInformation
    $results | Export-csv ".\sectionTemp.csv" -Append -NoTypeInformation
}

try{ 
$OrdersA = Import-CSV -Path .\sectionTemp.csv
$matchcounter = 0

foreach ($order1 in $OrdersA){
    $matched = $false
    foreach ($order2 in $courses)
    {
         $obj = "" | select "SIS ID","course_id","Section Name","Term StartDate","Term EndDate","restrict_enrollments_to_section_dates","nonxlist_course_id","sis_section_id","sis_course_id","integration_id","sis_import_id","School SIS ID"
        if(($order1.'course_id' -replace "A" ) -eq $order2.'id' )
        {
            $matchCounter++
            $matched = $true
            $obj.'SIS ID' = $order1. 'id'
            $obj.'course_id' = $order1.'course_id'
            $obj.'Section Name' = $order1.'name'
            $obj.'Term StartDate' = $order1.'start_at'
            $obj.'Term EndDate' = $order1.'end_at'
            $obj.'created_at' = $order1.'created_at'
            $obj.'restrict_enrollments_to_section_dates' = $order1.'restrict_enrollments_to_section_dates'
            $obj.'nonxlist_course_id' = $order1.'nonxlist_course_id'
            $obj.'sis_section_id' = $order1.'sis_section_id'
            $obj.'sis_course_id' = $order1.'sis_course_id'
            $obj.'integration_id' = $order1.'integration_id'
            $obj.'sis_import_id' = $order1.'sis_import_id'
            $obj.'School SIS ID' = $order2.'account_id'
                       
            Write-Host "Match Found Orders " "$matchCounter"
            $obj | Export-Csv -Path .\section.csv -Append -NoTypeInformation
        }
    }
}
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
write-host "creating student.csv"
  try{
  $tempsection = import-csv sectionTemp.csv
  $secid = $tempsection.id
#user Details
foreach($s in $secid)
{
    $uri2 = "/api/v1/sections/" + "$s" + "/enrollments"
    try{
    $results = Get-CanvasApiResult -Uri $uri2 -Method GET
    }
    Catch {
    $_.Exception | Out-File $logfile -Append
   }
    $results | convertto-Csv -NoTypeInformation
    $results | Export-csv -path ".\users.csv" -Append -NoTypeInformation
}}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
#if you want user data with in section details run below code#
try{
$matchcounter = 0
$user= import-csv .\users.csv
$Section = import-csv .\section.csv

foreach ($order1 in $user){
    $matched = $false
    foreach ($order2 in $Section)
    {
         $obj = "" | select "id","SIS ID","course_id","type","Username","FirstName","LastName","Password","Section SIS ID","School SIS ID","enrollment_state","role","sis_account_id","sis_course_id","sis_section_id","sis_user_id","html_url","user"
        if(($order1.'course_id' -replace "A" ) -eq $order2.'course_id' )
        {
            $matchCounter++
            $matched = $true
            $obj.'id' = $order1.'id'
            $obj.'SIS ID' = $order1.'user_id'
            $obj.'course_id' = $order1.'course_id'
            $obj.'type' = $order1.'type'
            $obj.'Username' = ''
            $obj.'FirstName' = ''
            $obj.'LastName' = ''
            $obj.'Password' = 'P@ssword'
            $obj.'Section SIS ID' = $order1.'course_section_id'
            $obj.'School SIS ID' = $order2.'School SIS ID'
            $obj.'enrollment_state' = $order1.'enrollment_state'
            $obj.'role' = $order1.'role'
            $obj.'sis_account_id' = $order1.'sis_account_id'
            $obj.'sis_course_id' = $order1.'sis_course_id'
            $obj.'sis_section_id' = $order1.'sis_section_id'
            $obj.'sis_user_id' = $order1.'sis_user_id'
            $obj.'html_url' = $order1.'html_url'
            $obj.'user' = $order1.'user'
                       
        

            Write-Host "Match Found Orders " "$matchCounter"
            $obj | Export-Csv -Path usernew.csv -Append -NoTypeInformation
        }
    }
}
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
Import-Csv -Path .\usernew.csv | ? role -eq 'StudentEnrollment' | select 'Section SIS ID', 'SIS ID' | sort 'Section SIS ID', 'SIS ID' | Export-Csv .\studentEnrollment.csv -NoTypeInformation

Import-Csv -Path .\usernew.csv | ? role -eq 'TeacherEnrollment' | select 'Section SIS ID', 'SIS ID' | sort 'Section SIS ID', 'SIS ID' | Export-Csv .\teacherroster.csv -NoTypeInformation

try{
$results = Get-CanvasApiResult -Uri "/api/v1/accounts/1/users" -Method GET
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
$results | convertto-Csv -NoTypeInformation
$results | Export-csv -path .\Fulluser.csv -Append -NoTypeInformation

try{
$matchcounter = 0
$Fulluser = import-csv .\Fulluser.csv
$user1= import-csv .\usernew.csv

foreach ($order1 in $Fulluser){
    $matched = $false
    foreach ($order2 in $user1)
    {
         $obj = "" | select "SIS ID","course_id","Username","First Name","Last Name","Password","Section SIS ID","School SIS ID","role","login_id","Email","sis_course_id","sis_section_id","sis_user_id","html_url","User Details"
       
        if($order1.'id' -eq $order2.'SIS ID' )
        {
            $matchCounter++
            $matched = $true
            $obj.'SIS ID' = $order2.'SIS ID'
            $obj.'course_id' = $order2.'course_id'
            $obj.'Username' = $order1.'name'
            $obj.'First Name' = $order1.'sortable_name'
            $obj.'Last Name' = $order1.'short_name'
            $obj.'Password' = 'P@ssword'
            $obj.'Section SIS ID' = $order2.'Section SIS ID'
            $obj.'School SIS ID' = $order2.'School SIS ID'           
            $obj.'role' = $order2.'role'
            $obj.'login_id' = $order1.'login_id'
            $obj.'Email' = $order1.'email'
            $obj.'sis_course_id' = $order2.'sis_course_id'
            $obj.'sis_section_id' = $order2.'sis_section_id'
            $obj.'sis_user_id' = $order2.'sis_user_id'
            $obj.'html_url' = $order2.'html_url'
            $obj.'User Details' = $order2.'user'
                       
            Write-Host "Match Found Orders " "$matchCounter"
            $obj | Export-Csv -Path .\usersall.csv -Append -NoTypeInformation
        }
    }
}
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
try{
Import-Csv -Path .\usersall.csv | ? role -eq 'StudentEnrollment' | sort 'SIS ID', 'School SIS ID' -Unique | Export-Csv .\student.csv -NoTypeInformation

Import-Csv -Path .\usersall.csv | ? role -eq 'TeacherEnrollment' | sort 'SIS ID', 'School SIS ID' -Unique | Export-Csv .\teacher.csv -NoTypeInformation
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
<#
#if you want details with courses run below script
#Teacher details
foreach($i in $id)
{
    $uri3 = "/api/v1/courses/"+"$i"+"/users?enrollment_type[]=teacher&include[]=enrollments"
    try{
    $results = Get-CanvasApiResult -Uri $uri3 -Method GET
    }
    Catch {
    $_.Exception | Out-File $logfile -Append
    }
    $results | convertto-Csv -NoTypeInformation
    $results | Export-csv .\Teacher1.csv -Append -NoTypeInformation
}
#student details
foreach($i in $id)
{
write-host "$i"
$uri4 = "/api/v1/courses/"+ "$i" +"/users?enrollment_type[]=student&include[]=enrollments"
try{
$results = Get-CanvasApiResult -Uri $uri4 -Method GET
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
$results | convertto-Csv -NoTypeInformation
$results | Export-csv .\Student1.csv -Append -NoTypeInformation
}
import-csv .\Teacher1.csv | sort id -Unique | export-csv .\Teacher2.csv -NoTypeInformation
$tempCSV = Import-Csv .\Teacher2.csv -Header "SIS ID","Username","created_at","sortable_name","short_name","sis_user_id","integration_id","sis_import_id","login_id","enrollments","email","School SIS ID" | select -skip 1
$tempCSV | Export-CSV .\Teacher2.csv -NoTypeInformation
import-csv .\Student1.csv | sort id -Unique | export-csv .\Student2.csv -NoTypeInformation
$tempCSV = Import-Csv .\Student2.csv -Header "SIS ID","Username","created_at","sortable_name","short_name","sis_user_id","integration_id","sis_import_id","login_id","enrollments","email","School SIS ID" | select -skip 1
$tempCSV | Export-CSV .\Student2.csv -NoTypeInformation
$matchcounter = 0
$student = import-csv .\Student2.csv
$teacher = import-csv .\Teacher2.csv
$user1= import-csv .\usernew.csv
try{
foreach ($order1 in $student){
    $matched = $false
    foreach ($order2 in $user1)
    {
         $obj = "" | select "SIS ID","course_id","Username","First Name","Last Name","Password","Section SIS ID","School SIS ID","role","login_id","Email","sis_course_id","sis_section_id","sis_user_id","html_url","User Details"
       
        if($order1.'SIS ID' -eq $order2.'SIS ID' )
        {
            $matchCounter++
            $matched = $true
            $obj.'SIS ID' = $order1.'SIS ID'
            $obj.'course_id' = $order2.'course_id'
            $obj.'Username' = $order1.'Username'
            $obj.'First Name' = $order1.'sortable_name'
            $obj.'Last Name' = $order1.'short_name'
            $obj.'Password' = 'P@ssword'
            $obj.'Section SIS ID' = $order2.'Section SIS ID'
            $obj.'School SIS ID' = $order2.'School SIS ID'           
            $obj.'role' = $order2.'role'
            $obj.'login_id' = $order1.'login_id'
            $obj.'Email' = $order1.'email'
            $obj.'sis_course_id' = $order2.'sis_course_id'
            $obj.'sis_section_id' = $order2.'sis_section_id'
            $obj.'sis_user_id' = $order2.'sis_user_id'
            $obj.'html_url' = $order2.'html_url'
            $obj.'User Details' = $order2.'user'
                       
            Write-Host "Match Found Orders " "$matchCounter"
            $obj | Export-Csv -Path .\Student0.csv -Append -NoTypeInformation
        }
    }
}
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
try{
foreach ($order1 in $teacher){
    $matched = $false
    foreach ($order2 in $user1)
    {
         $obj = "" | select "SIS ID","course_id","Username","First Name","Last Name","Password","Section SIS ID","School SIS ID","role","login_id","Email","sis_course_id","sis_section_id","sis_user_id","html_url","User Details"
        if($order1.'SIS ID' -eq $order2.'SIS ID' )
        {
            $matchCounter++
            $matched = $true
            $obj.'SIS ID' = $order1.'SIS ID'
            $obj.'course_id' = $order2.'course_id'
            $obj.'Username' = $order1.'Username'
            $obj.'First Name' = $order1.'sortable_name'
            $obj.'Last Name' = $order1.'short_name'
            $obj.'Password' = 'P@ssword'
            $obj.'Section SIS ID' = $order2.'Section SIS ID'
            $obj.'School SIS ID' = $order2.'School SIS ID'           
            $obj.'role' = $order2.'role'
            $obj.'login_id' = $order1.'login_id'
            $obj.'Email' = $order1.'email'
            $obj.'sis_course_id' = $order2.'sis_course_id'
            $obj.'sis_section_id' = $order2.'sis_section_id'
            $obj.'sis_user_id' = $order2.'sis_user_id'
            $obj.'html_url' = $order2.'html_url'
            $obj.'User Details' = $order2.'user'
                       
            Write-Host "Match Found Orders " "$matchCounter"
            $obj | Export-Csv -Path .\Teacher0.csv -Append -NoTypeInformation
        }
    }
}
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
Import-Csv -Path '.\Teacher0.csv'| sort 'SIS ID' -Unique | Export-Csv .\Teacher10.csv -NoTypeInformation
Import-Csv -Path '.\student0.csv'| sort 'SIS ID' -Unique | Export-Csv .\Student10.csv -NoTypeInformation
Import-Csv -Path '.\Teacher0.csv' | select 'Section SIS ID', 'SIS ID' | sort 'Section SIS ID', 'SIS ID' -Unique | Export-Csv .\StudentEnrollment1.csv -NoTypeInformation
Import-Csv -Path '.\student0.csv' | select 'Section SIS ID', 'SIS ID' | sort 'Section SIS ID', 'SIS ID' -Unique | Export-Csv .\Teacherroster1.csv -NoTypeInformation
#>
try{
remove-item .\Accounts.csv
remove-item .\sectionTemp.csv
remove-item .\Fulluser.csv
remove-item .\usersall.csv
remove-item .\usernew.csv
remove-item .\users.csv
remove-item .\courses.csv
}
Catch {
    $_.Exception | Out-File $logfile -Append
   }
}
else {
write-host "You need to provide sync.csv in order to continue..." 
}

$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan

#endregion

