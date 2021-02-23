# Exchange On Premise Meeting Migration
The script can be used to migrate Skype for Business to Teams meeting for a user

# Pre-Requisites :

## Create Azure AD Application

- In a tenant where you have Global Admin permissions, sign-in to https://portal.azure.com
- Navigate to Azure Active Directory > App registrations
- Click + New registration
- Provide a Name for your app (example: callRecordsApp)
- Click Register
- After app is registered, document the following
	
	* Application (client) ID: {guid} 
	* Directory (tenant) ID: {guid}
	
- In the left rail, navigate to Certificates & secrets
- Click + New client secret
- After new secret is generated, document the following

	* Client secret: {string}
	
- In the left rail, navigate to API permissions
- Click + Add a permission

	* Click Microsoft Graph
	* Click Application permissions
	* Expand OnlineMeetings (1) and check the box for OnlineMeetings.ReadWrite.All
	* Expand User (1) and check the box for User.Read.All
	
- Click Add permissions
- Remove any other permissions automatically added via App registration process
- Finally, click Grant admin consent for {tenantName}
- Output should look like the following
![Permissions Example](https://github.com/TeamsAdminSamples/PowerShell/blob/main/ExchOnPremMeetingMigration/Screenshot/ExchOnPremMeetingMigrationPermission.png)

## Install Ews Managed Api 2.2
- https://www.microsoft.com/en-us/download/details.aspx?id=42951

# Example : 

I used this script into an Azure function app. an Admin or a user can use the application to migrate meetings to Teams.
the idea is to use Powerapps, we can provide an email address and launch the script.


We have to provide an email address and submit the request.

![PowerApps Example](https://github.com/TeamsAdminSamples/PowerShell/blob/main/ExchOnPremMeetingMigration/Screenshot/ExchOnPremMeetingMigrationPowerapps.png)

When we submit the request, a power automate will be triggered.

![Power Automate Example](https://github.com/TeamsAdminSamples/PowerShell/blob/main/ExchOnPremMeetingMigration/Screenshot/ExchOnPremMeetingMigrationFlow.png?raw=true)

