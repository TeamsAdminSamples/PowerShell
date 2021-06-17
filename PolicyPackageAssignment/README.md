# PolicyPackageAssignment

# Description

This script will check the user license. After the script execution, teacher license users will be assigned to “Education Teacher”. Student license users will be assigned with  “Education_SecondaryStudent” policy packages

# Prerequisites

[Create a new Azure App](https://docs.microsoft.com/en-us/graph/auth-register-app-v2)

[How to apply permissions](https://docs.microsoft.com/en-us/graph/notifications-integration-app-registration) to your newly created App 

Please collect client id, client secret from created Azure App and tenant id from Azure portal

##### Required Permissions

|Permission type	|Permissions (from least to most privileged)|
|----|----|
|Application	|User.Read.All, User.ReadWrite.All, Directory.Read.All, Directory.ReadWrite.All|

# Example

If a user assigned with faculty license, script will assign Education_Teacher policy package to a user 

If the user assigned with a student license, script will assign Education_SecondaryStudent policy package to a user 

# Parameters

`-isLicensed`

Whether or not the user has any licenses assigned

Type: SkuPartNumber

# Inputs

Client_Id, Client_Secret, Tenantid

# Procedure to run the script
 
   To execute `PolicyPackageAssignment` download/copy and paste the script into PowerShell
        
   Provide the input parameters Client_Id, Client_Secret, TenantId and hit enter to proceed further on the script
   
   Please provide the global administrator credentials or Teams administrator credentials to `Connect-MicrosofTteams`
        
   Now the script will redirect to the web page for login
        
   ![Signin](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/Siginin.png)
        
   Provide admin credentials i.e user ID and password 
        
   Press enter to continue
   
   Once you are login it will show the below image for grant permissions for the app to perform operations

 ![GrantPermission](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/GrantPermissions.png)	
 
 ![GrantPermission](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/GrantPermissions2.png)
 
 **Click Accept**

 If you have provided the correct credentials it will give success status `admin_consent = True`
 
 Now press Y to proceed further in the script

# Sample Output

 Once the script executed, it generates the `license.csv` file contains no license or different license(other than Teams) users

|UserName  | Userid  |
|----------|---------|
|David Chew|81701046-cb37-439b-90ce-2afd9630af7d|
