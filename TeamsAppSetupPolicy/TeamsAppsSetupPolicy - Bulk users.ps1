#This script will assign custom Teams app setup policy to the users using PowerShell cmdlets, declare the file path and sheet name

$logfile = ".\TeamsAppSetupPolicyBulkUserlog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now 

   $file = ".\AssignTeamsAppsetupPolicy.xlsx"
   $sheetName = "Sheet1"

#Create an instance of Excel.Application and Open Excel file
   $objExcel = New-Object -ComObject Excel.Application
   $workbook = $objExcel.Workbooks.Open($file)
   $sheet = $workbook.Worksheets.Item($sheetName)
   $objExcel.Visible = $false
#Count max row
   $rowMax = ($sheet.UsedRange.Rows).count

#Declare the starting positions
   $rowUser, $colUser = 1, 1
   $rowPolicyname, $colPolicyname = 1, 2
$credential = Get-credential
Import-Module SkypeOnlineConnector
$sfbSession = New-CsOnlineSession
Import-PSSession $sfbSession
#loop to get values and store it
   for ($i = 1; $i -le $rowMax - 1; $i++) {

   $User = $sheet.Cells.Item($rowUser + $i, $colUser).text
    $Policyname = $sheet.Cells.Item($rowPolicyname + $i, $colPolicyname).text
    try{
     Grant-CsTeamsAppsetupPolicy -policyname "$Policyname" -Identity  $user
    }
Catch {
        $_.Exception | Out-File $logfile -Append
       }
       try{
get-csonlineuser -Identity "$user" |ft DisplayName,TeamsappsetupPolicy
}
Catch {
                $_.Exception | Out-File $logfile -Append
               }
}
$objExcel.quit()
$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
#end of script
