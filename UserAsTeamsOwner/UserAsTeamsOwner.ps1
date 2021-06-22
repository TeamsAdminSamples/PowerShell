# This script will provide the Teams details where user is owner for the Team
 $logfile = ".\CreateNew-CsTeamsMessagingPolicylog_$(get-date -format `"yyyyMMdd_hhmmsstt`").txt"
$start = [system.datetime]::Now
 
 If(Get-Module -ListAvailable -Name MicrosoftTeams) 
 { 
 Write-Host "MicrosoftTeams Already Installed" 
 } 
 else { 
 try { Install-Module -Name MicrosoftTeams
 Write-Host "Installed MicrosoftTeams"
 }
 catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
 }
 }
     
connect-microsoftteams
$user = Read-host "Please provide User Principle Name"
try{
$Data = get-team -User "$user"
}
catch{
        $_.Exception.Message | out-file -Filepath $logfile -append
 }

foreach ($teams in $Data) 
    {

        $groupid = $teams.Groupid
        $displayname = $teams.DisplayName
        $owner = get-teamuser -GroupId "$groupid" -Role Owner | Where-Object {$_.User -match "$user"}
        $owners = [string]::Join("; ",$owner.User)
        $groupid
        $owners

        if (!($owner -eq $null))
                    {
                    $file = New-Object psobject
                    $file | add-member -MemberType NoteProperty -Name Teams_Owner $owners
                    $file | add-member -MemberType NoteProperty -Name Teams_Displayname $displayname
                    $file | add-member -MemberType NoteProperty -Name Teams_Groupid $groupid
                    $file | export-csv output.csv -NoTypeInformation -Append
                    }
        }
$end = [system.datetime]::Now
$resultTime = $end - $start
Write-Host "Execution took : $($resultTime.TotalSeconds) seconds." -ForegroundColor Cyan



