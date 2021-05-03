# Grant New-CsTeamsMessagingPolicy

# Description
Script will assign a Teams Messaging Policy at the per-user scope, to run the script please install the [SFB online connector](https://www.microsoft.com/en-us/download/details.aspx?id=39366)
- Import the module into Windows PowerShell 
- Get the script from the 'Grant New-CsTeamsMessagingPolicy.ps1' file and paste it into windows PowerShell, then run the script

# Example
  Grant-CsTeamsMessagingPolicy -identity "Ken Myer" -PolicyName StudentMessagingPolicy

# Parameters
**`-Identity`**

Indicates the Identity of the user account the policy should be assigned to. User Identities can be specified using one of four formats: 1) the user's SIP address; 2) the user principal name (UPN); 3) the user's domain name and login name, in the form domain\logon (for example, litwareinc\kenmyer); and, 4) the user's Active Directory display name (for example, Ken Myer). User Identities can also be referenced by using the user's Active Directory distinguished name.

Type:	UserIdParameter
* * *
Position:	0
* * *
Default value:	None
* * *
Accept pipeline input:	False
* * *
Accept wildcard characters:	False

**`-PolicyName`**

The name of the custom policy that is being assigned to the user. To remove a specific assignment and fall back to the default tenant policy, you can assign to $Null.

Type:	String
* * *
Position:	1
* * *
Default value:	None
* * *
Accept pipeline input:	False
* * *
Accept wildcard characters:	False

# Prerequisite
1)	Install [SFB online connector](https://www.microsoft.com/en-us/download/details.aspx?id=39366)

# Inputs
Provide the parameters
`$policyname` Ex:StudentMessagingPolicy
`$user` Ex:davidchew@contoso.com

# How to run the script

1. As an Administrator, type PowerShell in the start menu. Right-click on Windows PowerShell, then select Run as Administrator
Click Yes at the UAC prompt

2. Run the **`Grant New-CsTeamsMessagingPolicy.ps1`**

# Output
davidchew@contoso.com is being assigned the StudentMessagingPolicy

A log file will be generated with exceptions, errors along with script execution time
