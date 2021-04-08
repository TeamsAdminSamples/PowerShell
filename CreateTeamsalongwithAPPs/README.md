# CreateTeamsalongwithAPPs

# Description

Create MicrosoftTeams along with APP for your tenant

# Prerequisites

[Create a new Azure App](https://docs.microsoft.com/en-us/graph/auth-register-app-v2)

[How to apply permissions](https://docs.microsoft.com/en-us/graph/notifications-integration-app-registration) to your newly created App

Please collect client id, client secret from created Azure app and tenant id from Azure portal
                
#### Required Permissions

| Permission type | Permissions (from least to most privileged)|
|-----------------|--------------------------------------------|
|Application|Group.Create, Group.ReadWrite.All, Directory.ReadWrite.All|

# Example

Group Name: HR, OwnerPrincipalName: AdeleV@contoso.com, AppName: OneNote

OneNote has been added to HR 

# Parameters

`-GroupId`

Specify a GroupId to convert to a Team. If specified, you cannot provide the other values that are already specified by the existing group, namely: Visibility, Alias, Description, or DisplayName. If, for example, you need to create a Team from an existing Microsoft 365 Group, use the ExternalDirectoryObjectId property value returned by [Get-UnifiedGroup](https://docs.microsoft.com/en-us/powershell/module/exchange/get-unifiedgroup?view=exchange-ps)

Type:	String
***
Position:	Named
***
Default value:	None
***
Accept pipeline input:	True
***
Accept wildcard characters:	False

`-AppId`

Teams App identifier in Microsoft Teams

Type:	String
***
Position:	Named
***
Default value:	None
***
Accept pipeline input:	True
***
Accept wildcard characters:	False

# Inputs

Groupname, OwnerPrincipalName, AppName, Client_Id, Client_Secret,Tenantid

# Procedure to run the script

   To execute `CreateTeamsalongwithAPPs` download/copy and paste the script into PowerShell
        
   Provide the input parameters Groupname, OwnerPrincipalName, AppName, Client_Id, Client_Secret, TenantId and hit enter to proceed further on the script
        
   Now the script will redirect to the web page for login
        
   ![Signin](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/Siginin.png)
        
   Provide Administrator credentials i.e user ID and password 
        
   Press enter to continue
   
   Once you are login it will show the below image for Grant permissions for the app to perform the operations

 ![GrantPermission](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/GrantPermissions.png)
 
 ![GrantPermission](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/GrantPermissions2.png)
 
 **Click Accept**

 If you have provided the correct credentials it will give success status `admin_consent = True`
 
 Now press Y to proceed further in the script
 
# Output

Script will execute and create the Team using provided Group Name, OwnerPrincipalName and APPName

App has been added to the team
