# DeleteTeamsWhereTeamMembersdontHaveGivenJobtitles

# Description

This script checks each Team member job title, if at least one job title does not match to given job titles, the script will delete those Teams. It will generate the ouput.csv file in the current folder and sent an email to deleted Team owners

This is a Graph API script, to execute the script user needs to create an Azure App and provide the necessary permissions 

# Prerequisites

[Create a new Azure App](https://docs.microsoft.com/en-us/graph/auth-register-app-v2)

[How to apply permissions](https://docs.microsoft.com/en-us/graph/notifications-integration-app-registration) to your newly created App

Please collect client id, client secret from created Azure app and tenant id from Azure portal

### Requried Permissions

|Permission type	          |  Permissions (from least to most privileged)|
|----------|-------------------|
|Application|Directory.AccessAsUser.All, Mail.Send|

# Example

Mailsender AdeleV@contoso.com

Keepjobtitles Manager, Hr, Associate 

Script will delete HR Team if anyone of Team member job title does not match and the script will send an email to HR Team owner behalf of mailsender(AdeleV@contoso.com)

# Parameters

 `mailsender`
 
   User Principal Name(for example AdeleV@contoso.com) to send an email to Teams owner of deleted Teams 
   
   Type: string 

 `KeepJobtitles`
 
   Designation of the employe(for example Manager)
   
   Type: string 
      
# Inputs
   
   Client_Id
   
   Client_Secret
   
   TenantId
   
   Mail sender
   
   KeepJobtitles
        
 # Procedure to run the script
 
   To excute `DeleteTeamsWhereTeamMembersdontHaveGivenJobtitles` download/copy and paste the script into PowerShell
        
   Provide the input parameters Mailsender, Client_Id, Client_Secret, TenantId, KeepJobtitles and hit enter to proceed further on the script
        
   Now the script will redirect to the web page for login
        
   ![Signin](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/Siginin.png)
        
   Provide admin credentials i.e user ID and password 
        
   Press enter to continue
   
   Once you are login it will show the below image for grant permissions for the app to perform the operations

 ![GrantPermission](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/GrantPermissions.png)
 
 ![GrantPermission](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/GrantPermissions2.png)
 
 **Click Accept**

 If you have provided the correct credentials it will give success status `admin_consent = True`
 
 Now press Y to proceed further in the script
       
 # Output
 
 Script will export the output.csv file which contains a list of deleted Teams along with owners, also it will send an email to Teams owner behalf of mail sender
 
 ##### Example
 
 |DeletedTeam|TeamsOwner        |
 |-----------|------------------|
 |HR         |fannyd@contoso.com|
 |Accounts   |danas@contoso.com |
