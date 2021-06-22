# User As Teams Owner

# Description

 Script will provide the Teams details where user as the owner of a Team in your organization
 
# Prerequisite 

 MicrosoftTeams module. Reference-[Microsoft Teams PowerShell Release Notes](https://docs.microsoft.com/en-us/microsoftteams/teams-powershell-release-notes)
 
# Procedure

1. As an Administrator, type PowerShell in the start menu. Right-click on Windows PowerShell, then select Run as Administrator.
Click Yes at the UAC prompt

2. Type the following within PowerShell and then press enter\
    **`Install-Module MicrosoftTeams`**
  
3. Type Y at the prompt. Press enter

4. If you are prompted for an untrusted repository, then type A (Yes to All) and press Enter. The module will now install

   - Script will prompt you for input user, provide the input to proceed
   - To connect to MicrosoftTeams provide the Global Administrator credentials or Teams Administrator credentials
   - Script will get the Teams details where user as the owner for the Team in your organization
   - Exports details to an output.csv file

# Example 

User: dmx1@example.com

|Team Owner| Team Displayname|Teamid|
|-----|----|---|
|dmx1@example.com| HR| 208bfb7a-9d4c-xxxx-8677-18cc7fcxxxxx |
|dmx1@example.com|Accounts|48ddcc0e-xxxx-4131-abf8-36axxxxx86ba|

# Parameters

`-UserPrincipalName`

Specifies the user ID of the user to retrieve.

Type:	String
***
Position:	Named
***
Default value:	None
***
Accept pipeline input:	True
***
Accept wildcard characters:	False
 
# Input 

UserPrincipalName

# Output
 The script will export details of user as the owner of MicrosoftTeams in your tenant with following details to an output.csv file
|Team Owner| Team Displayname|Teamid|
|----|---|---|

A log file will be generated with exceptions, errors along with script execution time
