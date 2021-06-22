# Add Teachers to Distribution list-Education domain

# Description

Script will search and filter teachers in a tenant using license parameter and adds to the *All teachers* distribution list

System should have the Azureadpreview module [`Install-Module AzureADPreview`](https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0-preview#installing-the-azure-ad-module) to execute the script

# Inputs
Global Administrator or Azure AD Administrator credentials 

# Prerequisites
As an Administrator, type PowerShell in the start menu

Right-click Windows PowerShell, then select Run as Administrator. Click Yes at the UAC prompt

1.	Type the following within PowerShell and then press Enter:

     `Install-Module AzureAd`

2.	Type Y at the prompt.Press Enter

3.	If you are prompted for an untrusted repository, then type A (Yes to All) and press Enter. The module will now install

# How to run the script
To run the script you will need to either download it or copy and paste the script into PowerShell

Provide the global administrator credentials or AzureAD admin credentials when it prompts

Hit enter to continue

# Output
Script will provide the count of teachers who has the teacher license, count of teachers in DL and log file will be generated with exceptions, errors along with script execution time
