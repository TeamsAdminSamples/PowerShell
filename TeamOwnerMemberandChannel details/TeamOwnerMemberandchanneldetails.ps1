#This script returns owners, members of a Team, and channels of a Team by providing the required input 1 or 2
$logfile = ".\TeamOwnerMemberandChanneldetailslog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now 
connect-microsoftteams
    function Get-Result(){
    Write-Host "1-To get the Team Owner and Member details
                2-To get the Channel details"
                
    $proceed = Read-host "Please provide input number 1 or 2" 

 if ($proceed -eq '1')
 {
 try{
 $Teams = get-team 
 foreach ($team in $Teams)
       {
        $groupid = $team.Groupid
        $displayname = $team.DisplayName
        $Teammember = get-teamuser -GroupId "$groupid" -Role Member
        $TeamOwner = get-teamuser -GroupId "$groupid" -Role Owner
      
        $Members = [string]::Join("; ",$Teammember.User)
        $Owner = [string]::Join("; ",$TeamOwner.User)
        #$groupid
        #$Memebrs
        #Owner

            $file = New-Object psobject
            $file | add-member -MemberType NoteProperty -Name Teamid $groupid
            $file | add-member -MemberType NoteProperty -Name TeamDisplayname $displayname
            $file | add-member -MemberType NoteProperty -Name Owner  $Owner
            $file | add-member -MemberType NoteProperty -Name Member $Members
            $file | export-csv -path ".\Teamoutput.csv" -NoTypeInformation -Append
       }     
    
    }
    catch{
    $_.Exception.Message | out-file -Filepath $logfile -append
    }
    }
    elseif($proceed -eq '2')
    {
    try{
    $Teams = get-team 
    foreach ($team in $Teams)
    {
    $channel=Get-Teamchannel -GroupId $team.Groupid
    $channels = [string]::Join("; ",$channel.DisplayName)

       $file = New-Object psobject
       $file | add-member -MemberType NoteProperty -Name Teamid $team.Groupid
       $file | add-member -MemberType NoteProperty -Name TeamDisplayname $team.displayname
       $file | add-member -MemberType NoteProperty -Name ChaneelName  $channels
       $file | export-csv -path ".\channeloutput.csv" -NoTypeInformation -Append
    }
      }
    catch{
       $_.Exception.Message | out-file -Filepath $logfile -append
         }
    }
   }
 do
 {
    $ProceedNext = Read-host "To proceed enter Y to continue"
    if ($ProceedNext -eq "Y" ) 
        { 
        Get-Result 
        } 
        else 
        { 
        break 
        } 
    }
    while($true); 

$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan   
$resultTime.TotalSeconds 
#end of script
