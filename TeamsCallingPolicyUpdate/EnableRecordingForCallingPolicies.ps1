#Requires -Module @{ ModuleName = 'MicrosoftTeams'; RequiredVersion = '2.3.1' }

[CmdletBinding(DefaultParameterSetName = "Individual")]
param (
    [Parameter(ParameterSetName = "Individual",
        Mandatory = $true)]
    [ValidateNotNullOrEmpty]
    [string]
    $PolicyName,

    [string]
    $PolicySuffix = "AllowRecording",

    [Parameter(ParameterSetName = "All",
        Mandatory = $true)]
    [switch]
    $All
)

# Validate MicrosoftTeams module is connected, otherwise, connect
if($null -eq [Microsoft.TeamsCmdlets.Powershell.Connect.TeamsPowerShellSession]::SessionProvider -or ![Microsoft.TeamsCmdlets.Powershell.Connect.TeamsPowerShellSession]::SessionProvider.ClientAuthenticated()) {
    Connect-MicrosoftTeams
}

$IdentityParam = @{}
if (!$All) {
    $IdentityParam['Identity'] = $PolicyName
}
$CurrentPolicies = Get-CsTeamsCallingPolicy @IdentityParam

$PolicyMembers = @(
    "Description",
    "AllowPrivateCalling",
    "AllowWebPSTNCalling",
    "AllowVoicemail",
    "AllowCallGroups",
    "AllowDelegation",
    "AllowCallForwardingToUser",
    "AllowCallForwardingToPhone",
    "PreventTollBypass",
    "BusyOnBusyEnabledType",
    "MusicOnHoldEnabledType",
    "SafeTransferEnabled",
    "AllowTranscriptionForCalling",
    "LiveCaptionsEnabledTypeForCalling",
    "AutoAnswerEnabledType",
    "SpamFilteringEnabledType"
)

foreach ($Policy in $CurrentPolicies) {
    $CurrentName = $Policy.Identity -replace '^tag:',''
    $NewName = $CurrentName + $PolicySuffix

    if ($Policy.AllowCloudRecordingForCalls) {
        Write-Warning "Policy $CurrentName already allows cloud recording, skipping..."
        continue
    }

    # validate policy with same name does not currently exist
    $NewPolicyExists = if ($All) {
        $CurrentPolicies | Where-Object { $_.Identity -replace '^tag:','' -eq $NewName }
    } else {
        Get-CsTeamsCallingPolicy -Identity $NewName -ErrorAction SilentlyContinue
    }
    if ($null -ne $NewPolicyExists) {
        Write-Warning "Policy $NewName already exists, skipping $CurrentName..."
        continue
    }
    $PolicyParams = @{
        Identity = $NewName
        AllowCloudRecordingForCalls = $true
    }
    foreach ($PM in $PolicyMembers) {
        $PolicyParams[$PM] = $Policy.$PM
    }
    New-CsTeamsCallingPolicy @PolicyParams
}