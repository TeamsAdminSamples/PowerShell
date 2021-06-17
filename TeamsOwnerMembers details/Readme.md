# Teams Owner Members details

# Description
The script will fetch the Teams owner and members details

# Prerequisite
MicrosoftTeams module. Reference-[Microsoft Teams PowerShell Release Notes](https://docs.microsoft.com/en-us/microsoftteams/teams-powershell-release-notes)

# Procedure

1. As an Administrator, type PowerShell in the start menu. Right-click Windows PowerShell, then select Run as Administrator.
Click Yes at the UAC prompt

2. Type the following within PowerShell command prompt and then press enter

    **`Install-Module MicrosoftTeams`** 
    
3. Type Y at the prompt, press enter

4. If you are prompted for an untrusted repository, then type A (Yes to All) and press Enter. The module will now install

- Get the script from `Teams_Owner and Members details.ps1` and paste it in Windows PowerShell command prompt
- Run the script, it will process the below steps

  1. Provide the Teams Administrator credentials to connect to MicrosoftTeams
  2. It will get the available Teams in the tenant
  3. After getting the available Teams in tenant, script will fetch the owner and members of each Team\
 Then exports the details of Teams in your tenant to a .csv file,**Output.csv** will store in the current folder
# Example 
 ```bash
 Get-Teamuser -GroupId 5e4aac3a-2547-4645-bb56-dafdb8733ccd -Role Member
 ```
```bash
 Get-Teamuser -GroupId 5e4aac3a-2547-4645-bb56-dafdb8733ccd -Role Owner 
 ```
# Output
 The details of each Team will stores in a .csv file with below details 
 
 |Team Name| Team id|Team Owner|Team member|
 |---|---|---|---|
 
 A log file will be generated with exceptions, errors along with script execution time
