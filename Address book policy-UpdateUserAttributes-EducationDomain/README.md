# Address book policy-UpdateUserAttributes

# Description
Script will check all groups of teachers, students and removes the part “students of” or “teachers of” and then it will update the user custom attribute with the school name, this way we can have both teachers, students of the same school can share the same attribute value

System should have the AzureADPreview module [`Install-Module AzureADPreview`](https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0-preview#installing-the-azure-ad-module) and [`Exchange Online`](https://docs.microsoft.com/en-us/powershell/module/exchange/connect-exchangeonline?view=exchange-ps) to exceute the script

# Inputs
Global Administrator or Azure AD Administrator and ExchangeOnline Administrator credentials

# Prerequisites
Exchange Online and AzureAD module

# How to run the script
As an Administrator, type PowerShell in the start menu. Right-click on Windows PowerShell, then select run as Administrator. Click Yes at the UAC prompt

Run the script 

Provide the Global Administrator credentials or AzureAD and ExchangeOnline Administrator credentials when it prompts

Hit enter to continue

# Expected output
User custom attributes are updated

A log file will be generated with exceptions, errors along with script execution time
