# Group or TeamsCreationRestrictionPolicy-Bulk

# Description

You can restrict Office 365 group creation to the members of a particular security group

Office 365 global admins can create groups via any means, such as the Microsoft 365 admin center, Planner, Teams, Exchange, and SharePoint Online

The system should have the AzureADPreview module [`Install-Module azure preview`](https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0-preview#installing-the-azure-ad-module)

# Example
Restricting HR group members from creating Teams or groups

# Inputs
Import _GroupTeamsCreationRestrictionPolicy.xlsx_ file as an input which contains Groupname and AllowGroupCreation, Please refer example table

 | Groupname    | AllowGroupCreation    |
 |--------------|--------------------   |
 | Group1       | True                  |
 | HR       | False                 |

# Parameters
Groupname: The name of the created O365 security group

AllowGroupCreation: Do You want to allow this group to create Teams True/False

# Prerequisites
As an Administrator, type PowerShell in the start menu
Right-click Windows PowerShell, then select Run as Administrator. Click Yes at the UAC prompt.
1.	Type the following within PowerShell and then press Enter:

     `Install-Module AzureAd`

2.	Type Y at the prompt. Press Enter

3.	If you are prompted for an untrusted repository, then type A (Yes to All) and press Enter. The module will now install

# How to run the script
To run the script you will need to either download it or copy and paste the script into PowerShell

Provide the Global Administrator credentials or AzureAD Administrator credentials when it prompts

The script will restrict or allow the group users based on AllowGroupCreation input

# Output
The last line of the script will display the updated settings:
![output](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/Restricting%20group%20creation.png)

A log file will be generated with exceptions, errors along with script execution time
