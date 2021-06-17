# This script will provide Total stream count and cqd report of given time using Teams powershell cqd module cmdlets
$start = [system.datetime]::Now
$logfile = ".\CQDLog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
 
$proceed = Read-host " 
Provide 1 For Total Steram count including Audio,video,Appsharing
Provide 2 for CQD Report of Given time"
if ($proceed -eq "1")
{
try{
write-host "Provide startDate and Enddate in MM/dd/yyyy H:mm(Ex:31-03-2020 4:34)"

$StartDate = read-host "Please provide start date"
$EndDate = read-host "please provide end date"

$dimensions = "AllStreams.Date","AllStreams.Media Type","AllStreams.Second UPN" 
$measures = "Measures.Total Stream Count","Measures.Audio Stream Count","Measures.Video Stream Count","Measures.AppSharing Stream Count","Measures.VBSS Stream Count"
#"Measures.Call Count","Measures.Audio Call Count","Measures.Video Call Count"

$CustomFilter = @()
$F1 = New-Object pscustomobject
$F1 | Add-Member -Type NoteProperty -Name FName -Value "AllStreams.Is Teams"
$F1 | Add-Member -Type NoteProperty -Name FValue -Value "1"
$F1 | Add-Member -Type NoteProperty -Name Op -Value 0
$CustomFilter += $F1

$F2 = New-Object pscustomobject
$F2 | Add-Member -Type NoteProperty -Name FName -Value "AllStreams.Second UserType"
$F2 | Add-Member -Type NoteProperty -Name FValue -Value "User"
$F2 | Add-Member -Type NoteProperty -Name Op -Value 0
$CustomFilter += $F2

  $CQDTableTemp= Get-CQDData -OutPutType CSV -OutPutFilePath cqdoutput.csv  -CQDVer V3 -LargeQuery -StartDate $StartDate -EndDate $EndDate -IsServerPair 'Client : Server','Client : Client'  `
    -Dimensions $dimensions -Measures $measures -customfilter $CustomFilter  -ShowQuery $true 
}
catch
{
$_.Exception.Message | out-file -Filepath $logfile -append
}
    }

if ($proceed -eq "2")
    {
    try{
    $cqd_List = Import-Csv -path ".\CQD_data.csv"


Foreach ($cqd in $cqd_List)
{

Get-CQDData -Dimensions $cqd.Dimensions -Measures $cqd.Measures -OutPutFilePath $cqd.OutPutFilePath -StartDate $cqd.StartDate -EndDate $cqd.EndDate -OutPutType $cqd.OutPutType -MediaType $cqd.MediaType -IsServerPair $cqd.IsServerPair -OverWriteOutput
}
   }
   catch
{
$_.Exception.Message | out-file -Filepath $logfile -append
}
    }

    Else   {
    write-host "Please run the script again choose option 1 or 2"
    }
$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan
#end of script

