# CreateAndAssignAddressBookPolicy-EducationDomain

# Description

Script will create and assign the address lists, address book policies to users based on the attribute

A log file will be generated with exceptions, errors along with script execution time

# Inputs
Global Administrator or ExchangeOnline Administrator credentials 

# Prerequisites
As an Administrator, type PowerShell in the start menu

Right-click on Windows PowerShell, then select Run as Administrator. Click Yes at the UAC prompt
1.	Type the following within PowerShell and then press Enter:

     `Import-Module ExchangeOnlineManagement`
     
     `$UserCredential = Get-Credential`

2. In the Windows PowerShell credential request dialog box that appears, type global or exchange online admin account and password, and then click OK
 
# How to run the script
To run the script you will need to either download it or copy and paste the script into PowerShell

Provide the Global Administrator credentials or exchange online Administrator credentials when it prompts

Hit enter to continue
