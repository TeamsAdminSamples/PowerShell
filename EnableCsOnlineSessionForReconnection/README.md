# Enable-CsOnlineSessionForReconnection
Allows interactive sessions to persist for the Microsoft Teams PowerShell module. This only works with version 1.1.6 and 1.1.10-preview. It is no longer needed since version 1.1.11-preview. 

You must dot source this file and then run the function that it imports after creating an implicit remoting session for Microsoft Teams.


~~~~ 
# dot source the function
. .\Enable-CsOnlineSessionForReconnection.ps1

# Connect to Teams
Connect-MicrosoftTeams
Import-PSSession (New-CsOnlineSession)

# run the function
Enable-CsOnlineSessionForReconnection
~~~~