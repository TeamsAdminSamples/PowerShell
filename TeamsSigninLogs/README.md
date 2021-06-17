# TeamsSigninLogs

# Description

Retrieve the MicrosoftTeams user sign-ins for your tenant, script will check audit logs and export the file, it contains the Teams sign-in username along with device name

# Prerequisites

[Create a new Azure App](https://docs.microsoft.com/en-us/graph/auth-register-app-v2)

[How to apply permissions](https://docs.microsoft.com/en-us/graph/notifications-integration-app-registration) to your newly created App

Please collect client id, client secret from created Azure App and tenant id from Azure portal

#### Required Permissions

  | Permission type	                   |  Permissions (from least to most privileged)|
  |------------------------------------|---------------------------------------------|
  | Application	                       | AuditLog.Read.All and Directory.Read.All    |
 
# Parameters

`-Auditlogs`

 records of system activities
 
 Type: Logs

# Inputs

Client_Id, Client_Secret, Tenantid

 # Procedure to run the script
 
   To execute `TeamsSigninLogs` download/copy and paste the script into PowerShell
        
   Provide the input parameters Client_Id, Client_Secret, TenantId and hit enter to proceed further on the script
        
   Now the script will redirect to the web page for login
        
   ![Signin](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/Siginin.png)
        
   Provide Administrator credentials i.e user ID and password 
        
   Press enter to continue
   
   Once you are login it will show the below image for grant permissions for the app to perform operations

 ![GrantPermission](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/GrantPermissions.png)	
 
 ![GrantPermission](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/GrantPermissions2.png)
 
 **Click Accept**

 If you have provided the correct credentials it will give success status `admin_consent = True`
 
 Now press Y to proceed further in script

# Output

_Signinoutput. csv_ is the final output file having sign-in details 

 | UserUPN	| CreatedDateTime	| resourceDisplayName | AppDisplayName	| IsInteractive |	DeviceDetail |
 |----------|-------------------|---------------------|------------------|--------------|----------------|
 |davidchew@contoso.com|2020-03-23T15:10:59.2906713Z|Microsoft Teams Web Client	|FALSE	|@{deviceId=; displayName=; operatingSystem=Windows 10; browser=Chrome 80.0.3987; isCompliant=; isManaged=; trustType=}|
