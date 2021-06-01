#This script will restrict Office 365 group creation to the members of a particular security groups using AzureAD PowerShell cmdlets 
$logfile = ".\GroupTeamsCreationRestrictionPolicyBulklog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now
try{
   $conDetails = Connect-AzureAD
   $tenantDomain = $conDetails.TenantDomain
}
catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}

#Declare the file path and sheet name
   $file = ".\GroupTeamsCreationRestrictionPolicy.xlsx"
   $sheetName = "Sheet1"

#Create an instance of Excel.Application and Open Excel file
   $objExcel = New-Object -ComObject Excel.Application
   $workbook = $objExcel.Workbooks.Open($file)
   $sheet = $workbook.Worksheets.Item($sheetName)
   $objExcel.Visible = $false
#Count max row
   $rowMax = ($sheet.UsedRange.Rows).count

#Declare the starting positions
   $rowGroupName, $colGroupName = 1, 1
   $rowAllowGroupCreation, $colAllowGroupCreation = 1, 2

#loop to get values and store it
   for ($i = 1; $i -le $rowMax - 1; $i++) {

   $GroupName = $sheet.Cells.Item($rowGroupName + $i, $colGroupName).text
    $AllowGroupCreation = $sheet.Cells.Item($rowAllowGroupCreation + $i, $colAllowGroupCreation).text

$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
$settingsObjectID = (Get-AzureADDirectorySetting | Where-object -Property Displayname -Value "Group.Unified" -EQ).id
if(!$settingsObjectID)
    {try{
	  $template = Get-AzureADDirectorySettingTemplate | Where-object {$_.displayname -eq "group.unified"}
    $settingsCopy = $template.CreateDirectorySetting()
    New-AzureADDirectorySetting -DirectorySetting $settingsCopy
    $settingsObjectID = (Get-AzureADDirectorySetting | Where-object -Property Displayname -Value "Group.Unified" -EQ).id
     }
     catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
     }

    $settingsCopy = Get-AzureADDirectorySetting -Id $settingsObjectID
     $settingsCopy["EnableGroupCreation"] = $AllowGroupCreation

 if($GroupName)
   {try{
	$settingsCopy["GroupCreationAllowedGroupId"] = (Get-AzureADGroup -SearchString $GroupName).objectid
     }
     catch{
$_.Exception.Message | out-file -Filepath $logfile -append
}
     }

    Set-AzureADDirectorySetting -Id $settingsObjectID -DirectorySetting $settingsCopy

    (Get-AzureADDirectorySetting -Id $settingsObjectID).Values
}
$objExcel.quit()
$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
#end of script
