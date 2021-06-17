# User Policy Assignment
# Description:
UserPolicyAssignment script will work for assigning custom user policies for N no. of users\
To run the script please install [MicrosoftTeams module](https://www.powershellgallery.com/packages/MicrosoftTeams/2.0.0)
- Import the Module into Windows PowerShell 
- Get the script from the `UserPolicyAssignment.ps1` file and paste it into Windows PowerShell, then run the script
- Script has all the available policies to the user listed below, please provide the required input from 1 to 12 to apply the policy

                      1- TeamsAppSetupPolicy 
                      2- TeamsMeetingPolicy 
                      3- TeamsCallingPolicy
                      4- TeamsMessagingPolicy 
                      5- BroadcastMeetingPolicy
                      6- TeamsCallParkPolicy
                      7- CallerIdPolicy 
                      8- TeamsEmergencyCallingPolicy 
                      9- TeamsEmergencyCallRoutingPolicy
                      10-VoiceRoutingPolicy 
                      11-TeamsAppPermissionPolicy 
                      12-TeamsDailPlan

# Example
![User Policy](https://github.com/SwathiGugulot/Sample/blob/master/userpolisyAssignimage.PNG) \
In the list of policies, provided input number 2 to apply TeamsMeetingPolicy to user/users
# Input 
 Keep the UserPricipleName in Input.Csv file
 
 ![Example](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/Userpolicyassignment.PNG)
 
# Output
Custom policy assigned to the user

A log file will be generated with exceptions, errors along with script execution time
