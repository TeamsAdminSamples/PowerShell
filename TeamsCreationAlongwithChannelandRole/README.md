# TeamsCreationAlongwithChannelandRole

# Description

The script is to create Teams along with channel and role

Provide the input data `TeamName` `ChannelName` `Owner` `Member` and `Visibility` in .csv format and provide the path location in script

The script will generate the `output.csv` file which holds the details of created TeamName, TeamOwner, Member, ChannelName, ChannelId

# Prerequisites

1. As an Administrator, type PowerShell in the start menu. Right-click on Windows PowerShell, then select Run as Administrator. Click Yes at the UAC prompt

2. Type the following within PowerShell and then press Enter:\
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;**`Install-Module MicrosoftTeams`**
  
3. Type Y at the prompt. Click enter

4. If you are prompted for an untrusted repository, then type A (Yes to All) and press Enter. The module will now install

# Inputs

  DisplayName, ChannelName, Owner, Member, Visibility
  
# Parameters

**`-DisplayName`**

Team display name. Characters Limit - 256
* * *
Type:	String
* * *
Position:	Named
* * *
Default value:	None
* * *
Accept pipeline input:	True
* * *
Accept wildcard characters:	False

**`-ChannelName`**

Channel display name. Names must be 50 characters or less, and can't contain the characters # % & * { } / \ : < > ? + | ' "
- - -
Type:	String
- - -
Position:	Named
- - -
Default value:	None
- - -
Accept pipeline input:	True
 - -  -
Accept wildcard characters:	False


**`-Visibility`**

Set to Public to allow all users in your organization to join the group by default. Set to Private to require that an owner approve the join request

Type:	String
* * *
Position:	Named
* * *
Default value:	Private
* * *
Accept pipeline input:	True
* * *
Accept wildcard characters:	False

**`-Owner`**

An admin who is allowed to create on behalf of another user should use this flag to specify the desired owner of the group. This user will be added as both a member and an owner of the group. If not specified, the user who creates the team will be added as both a member and an owner
* * *
Type:	String
* * *
Position:	Named
* * *
Default value:	None
* * *
Accept pipeline input:	True
* * *
Accept wildcard characters:	False

# How to run the script

To run the script you will need to either download it or copy and paste the script into PowerShell

Provide the Global Administrator credentials or Teams Administrator credentials when it prompts

After execution script will export the created Teams details of MicrosoftTeams in your tenant with following details to an `output.csv` file

# Output

| TeamName |TeamId | ChannelName |ChannelId |TeamOwner | Team Member |
|----------|-------|-------------|----------|----------|-------------|

A log file will be generated with exceptions, errors along with script execution time
