# TeamsCallingPolicyUpdate

This script will allow you to duplicate Teams Calling Policies in your tenant to enable the use of Cloud Recording for calls.

This requires the Microsoft Teams 2.3.1 module which is available in the PowerShell Gallery [here](https://www.powershellgallery.com/packages/MicrosoftTeams/2.3.1).

You can either specify the `-PolicyName` parameter to update a specific policy, or supply the `-All` switch, which will duplicate all calling policies configured in your tenant. 

The `-PolicySuffix` parameter will allow you to define what you want the new policies to be named, by default, this is set to `AllowRecording`, resulting in the new policy name of `ExistingPolicyNameAllowRecording`.

You must grant the newly created policies to your users for them to take effect:
[Grant-CsTeamsCallingPolicy](https://docs.microsoft.com/en-us/powershell/module/skype/grant-csteamscallingpolicy?view=skype-ps)