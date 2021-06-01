# TeamsUserActivity

# Description

Use the MicrosoftTeams activity reports to get insights into the Microsoft Teams user activity in your organization. The period specifies the length of time over which the report is aggregated. The supported values for {period_value} are: D7, D30, D90, and D180

Provide the number(1,2,3) to get the MicrosoftTeams user activity reports

Reference [Microsoft Teams user activity reports](https://docs.microsoft.com/en-us/graph/api/resources/microsoft-teams-user-activity-reports?view=graph-rest-1.0)

# Prerequisites

[Create a new Azure App](https://docs.microsoft.com/en-us/graph/auth-register-app-v2)

[How to apply permissions](https://docs.microsoft.com/en-us/graph/notifications-integration-app-registration) to your newly created App

Please collect client id, client secret from created Azure App and tenant id from Azure portal

#### Required Permissions

|Permission type	|Permissions (from least to most privileged)|
|----|---|
|Application	|Reports.Read.All|

# Example

For input 1

|Report Refresh Date	|User Principal Name	|Last Activity Date	|Is Deleted|	Deleted Date	|Assigned Products	|Team Chat Message Count|	Private Chat Message Count	|Call Count	|Meeting Count|Has Other Action	|Report Period|
|---|---|---|---|---|---|---|---|---|---|---|---|
|25-02-2020|IrvinS@M365x726831.OnMicrosoft.com		|FALSE		|ENTERPRISE MOBILITY + SECURITY E5+OFFICE 365 E5	|0	|0|	0	|0|	No	|30|

# Parameters

`-Days`

Total number of days 

Type: String

# Inputs 

Client_Id, Client_Secret, Tenantid, Period

1-GetTeamsUserActivityUserDetail

2-GetTeamsUserActivityCounts

3-GetTeamsUserActivityUserCounts

 # Procedure to run the script
 
   To execute `TeamsUserActivity` download/copy and paste the script into PowerShell
        
   Provide the input parameters Client_Id, Client_Secret, TenantId, Period and hit enter to proceed further on the script
        
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

# Expected Output

Script will generate the TeamsUserActivity reports for provided input
