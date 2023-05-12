[CmdletBinding()]

param(
    [Parameter(Mandatory)]
    [string]
    $RGSExportPath,

    [Parameter(Mandatory)]
    [string]
    $SipDomain,

    # Default Usage Location
    [Parameter()]
    [string]
    $UsageLocation = 'US'
)

# TODO: add lines for required modules
#       need: Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Users.Actions
#       need: MicrosoftTeams

# TODO: Add switch for phone number assignment
# TODO: how to handle DR vs Service Number in addition to Hybrid vs Online accounts

$GeneratedScriptsPath = [IO.Path]::Combine($RGSExportPath, 'GeneratedScripts')
if (!(Test-Path -Path $GeneratedScriptsPath)) {
    $null = New-Item -Path $GeneratedScriptsPath -ItemType Directory
}

$Files = Get-ChildItem -Path $RGSExportPath -Filter *.zip
$i = 0
$CreatedVariables = @()
foreach ($file in $Files.Name) {
    $Source = [IO.Path]::Combine($RGSExportPath, $File)
    $Destination = [IO.Path]::Combine($RGSExportPath, [IO.Path]::GetFileNameWithoutExtension($File))
    if (!(Test-Path -Path $Destination)) {
        Expand-Archive -Path $Source -DestinationPath $Destination -Force
    }

    $RgsFilesPath = [IO.Path]::Combine($Destination, 'RgsImportExport')
    $RgsFiles = Get-ChildItem -Path $RgsFilesPath -Filter *.xml -File
    foreach ($RgsFile in $RgsFiles) {
        $FileName = $RgsFile.BaseName
        if ($CreatedVariables -notcontains $FileName) { $CreatedVariables += $FileName }
        $FileContent = Import-Clixml -Path $RgsFile.FullName
        if ($FileContent -isnot [Object[]]) {
            $FileContent = @($FileContent)
        }
        if ($i -eq 0 -and (Test-Path Variable:$FileName)) {
            Remove-Variable -Name $FileName -Force -ErrorAction SilentlyContinue
        }
        $CurrentValue = if ((Test-Path Variable:$FileName)) {
            Get-Variable -Name $FileName -ValueOnly -ErrorAction SilentlyContinue
        }
        New-Variable -Name $FileName -Value ($FileContent + $CurrentValue) -Force
    }
    $i++
}

function RoundValue {
    param (
        [int]$InputTime,
        [int]$MinimumTime = 30,
        [int]$RoundTo = 15,
        [int]$MaximumTime = 180
    )
    if ([string]::IsNullOrEmpty($InputTime) -or $InputTime -lt $MinimumTime) {
        [int]$MinimumTime
    }
    else {
        [int][Math]::Min(([Math]::Ceiling($InputTime / $RoundTo) * $RoundTo), $MaximumTime)
    }
}

function WriteDisconnectAction {
    param (
        [Text.StringBuilder]
        $CommandText,
        [string]
        $CommandHashName,
        [string]
        $ActionName,
        [switch]
        $AAMenuOption,
        [int]
        $NumTabs = 0,
        [string]
        $Prepend = ''
    )
    $Tabs = [string]::new(' ', $NumTabs * 4)
    if ($AAMenuOption) {
        $null = $CommandText.AppendLine("$Prepend$Tabs`$$CommandHashName['Action'] = 'DisconnectCall'")
    }
    else {
        $null = $CommandText.AppendLine("$Prepend$Tabs`$$CommandHashName['$ActionName'] = 'Disconnect'")
    }
}

function ConvertNonQuestionAction {
    [cmdletbinding()]
    param (
        $FlowName,
        $URI,
        $Action,
        $QueueId,
        [Text.StringBuilder]
        $CommandText,
        [string]
        $CommandHashName,
        [string]
        $ActionName,
        [string]
        $ActionTargetName,
        [Parameter(ParameterSetName = 'AA')]
        [switch]
        $AAMenuOption,
        [Parameter(ParameterSetName = 'CQ')]
        [string]
        $AudioFilePromptLocation,
        [Parameter(ParameterSetName = 'CQ')]
        [string]
        $AudioFilePromptOriginalName,
        [Parameter(ParameterSetName = 'CQ')]
        [string]
        $TextPrompt,
        [Parameter(ParameterSetName = 'CQ')]
        [string]
        $AudioPromptParamName,
        [Parameter(ParameterSetName = 'CQ')]
        [string]
        $TextPromptParamName,
        [string]
        $Prepend = ''
    )
    <#
        AudioFilePromptLocation     = $DefaultQueue.OverflowAudioStoredLocation
        AudioFilePromptOriginalName = $DefaultQueue.OverflowAudioOriginalName
        TextPrompt                  = $DefaultQueue.OverflowTextPrompt
        AudioPromptParamName        = 'OverflowSharedVoicemailAudioFilePrompt'
        TextPromptParamName         = 'OverflowSharedVoicemailTextToSpeechPrompt'
    #>
    $WarningStrings = [Text.StringBuilder]::new()
    switch ($Action) {
        'Terminate' {
            WriteDisconnectAction -CommandText $CommandText -CommandHashName $CommandHashName -ActionName $ActionName -AAMenuOption:$AAMenuOption -Prepend $Prepend
        }
        'TransferToPstn' {
            if ([string]::IsNullOrEmpty($URI)) {
                WriteDisconnectAction -CommandText $CommandText -CommandHashName $CommandHashName -ActionName $ActionName -AAMenuOption:$AAMenuOption -Prepend $Prepend
            }
            else {
                $URI = [regex]::Replace($URI, '[Xx]', '')
                $URI = [regex]::Replace($URI, ';[Ee][Tt]=', 'x')
                $URI = 'tel:+' + [regex]::Replace($URI, '[^0-9x]', '')
                $URI = [regex]::Replace($URI, 'x', ';ext=')
                if ($AAMenuOption) {
                    $null = $CommandText.AppendLine("$Prepend`$CallableEntity = New-CsAutoAttendantCallableEntity -Identity '$URI' -Type ExternalPstn")
                    $null = $CommandText.AppendLine("$Prepend`$$CommandHashName['Action'] = 'TransferCallToTarget'")
                    $null = $CommandText.AppendLine("$Prepend`$$CommandHashName['CallTarget'] = `$CallableEntity")

                }
                else {
                    $null = $CommandText.AppendLine("$Prepend`$$CommandHashName['$ActionName'] = 'Forward'")
                    $null = $CommandText.AppendLine("$Prepend`$$CommandHashName['$ActionTargetName'] = '$URI'")
                }
            }
        }
        'TransferToQueue' {
            $Queue = $ProcessedQueues.Where( { $_.Identity -eq $QueueId })[0]
            if ($null -eq $Queue) {
                $null = $WarningStrings.AppendLine("$FlowName will to transfer to the queue with ID: $QueueId. This queue was not found, the $ActionName will be set to disconnect")
                $null = $CommandText.AppendLine("${Prepend}Write-Warning 'Queue not found, set to Disconnect'")
                WriteDisconnectAction -CommandText $CommandText -CommandHashName $CommandHashName -ActionName $ActionName -AAMenuOption:$AAMenuOption -Prepend $Prepend
            }
            else {
                $QueueName = 'CQ ' + (CleanName $Queue.Name)
                $null = $WarningStrings.AppendLine("$FlowName will to transfer to the $QueueName queue. Ensure queue exists online, otherwise, the $ActionName will be set to disconnect")
                $null = $CommandText.AppendLine("$Prepend`$Queue = try {")
                $null = $CommandText.AppendLine("$Prepend    @(Get-CsCallQueue -NameFilter '$QueueName' -First 1)[0].Identity")
                $null = $CommandText.AppendLine("$Prepend} catch {")
                $null = $CommandText.AppendLine("$Prepend    `$null")
                $null = $CommandText.AppendLine("$Prepend}")

                $null = $CommandText.AppendLine("${Prepend}if ([string]::IsNullOrEmpty(`$Queue)) {")
                $null = $CommandText.AppendLine("$Prepend    Write-Warning 'TransferToQueue could not find valid object for $QueueName, $ActionName set to Disconnect'")
                WriteDisconnectAction -CommandText $CommandText -CommandHashName $CommandHashName -ActionName $ActionName -AAMenuOption:$AAMenuOption -NumTabs 1 -Prepend $Prepend
                $null = $CommandText.AppendLine("$Prepend} else {")
                if ($AAMenuOption) {
                    $null = $CommandText.AppendLine("$Prepend    `$CallableEntity = New-CsAutoAttendantCallableEntity -Identity `$Queue -Type HuntGroup")
                    $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['Action'] = 'TransferCallToTarget'")
                    $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['CallTarget'] = `$CallableEntity")

                }
                else {
                    $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['$ActionName'] = 'Forward'")
                    $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['$ActionTargetName'] = `$Queue")
                }
                $null = $CommandText.AppendLine("$Prepend}")
            }
        }
        'TransferToUri' {
            $null = $WarningStrings.AppendLine("$FlowName will to transfer to $URI. Ensure user/object exists online, otherwise, the $ActionName will be set to disconnect")
            $null = $CommandText.AppendLine("$Prepend`$Target = try {")
            $null = $CommandText.AppendLine("$Prepend    (Get-CsOnlineUser -Identity '$URI' -ErrorAction Stop).Identity")
            $null = $CommandText.AppendLine("$Prepend} catch {")
            $null = $CommandText.AppendLine("$Prepend    `$null")
            $null = $CommandText.AppendLine("$Prepend}")
            $null = $CommandText.AppendLine("${Prepend}if ([string]::IsNullOrEmpty(`$Target)) {")
            $null = $CommandText.AppendLine("$Prepend    Write-Warning 'TransferToUri could not find valid object for $URI, $ActionName set to Disconnect'")
            WriteDisconnectAction -CommandText $CommandText -CommandHashName $CommandHashName -ActionName $ActionName -AAMenuOption:$AAMenuOption -NumTabs 1 -Prepend $Prepend
            $null = $CommandText.AppendLine("$Prepend} else {")
            if ($AAMenuOption) {
                $null = $CommandText.AppendLine("$Prepend    `$CallableEntity = New-CsAutoAttendantCallableEntity -Identity `$Target -Type User")
                $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['Action'] = 'TransferCallToTarget'")
                $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['CallTarget'] = `$CallableEntity")

            }
            else {
                $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['$ActionName'] = 'Forward'")
                $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['$ActionTargetName'] = `"`$Target`"")
            }
            $null = $CommandText.AppendLine("$Prepend}")
        }
        'TransferToVoicemailUri' {
            $null = $WarningStrings.AppendLine("$FlowName will to transfer to $([regex]::replace($URI,'^[Ss][Ii][Pp]:','')). Ensure Microsoft 365 Mail Enabled group exists, otherwise, the $ActionName will be set to disconnect")
            $null = $CommandText.AppendLine("$Prepend`$Target = try {")
            $null = $CommandText.AppendLine("$Prepend    (Find-CsGroup -SearchQuery '$([regex]::replace($URI,'^[Ss][Ii][Pp]:',''))' -ExactMatchOnly `$true -MaxResults 1 -MailEnabledOnly `$true).Id.Guid")
            $null = $CommandText.AppendLine("$Prepend} catch {")
            $null = $CommandText.AppendLine("$Prepend    `$null")
            $null = $CommandText.AppendLine("$Prepend}")
            $null = $CommandText.AppendLine("${Prepend}if ([string]::IsNullOrEmpty(`$Target)) {")
            $null = $CommandText.AppendLine("$Prepend    Write-Warning 'TransferToVoicemailUri could not find valid object for $([regex]::replace($URI,'^[Ss][Ii][Pp]:','')), $ActionName set to Disconnect'")
            WriteDisconnectAction -CommandText $CommandText -CommandHashName $CommandHashName -ActionName $ActionName -AAMenuOption:$AAMenuOption -NumTabs 1 -Prepend $Prepend
            $null = $CommandText.AppendLine("$Prepend} else {")
            if ($AAMenuOption) {
                $null = $CommandText.AppendLine("$Prepend    `$CallableEntity = New-CsAutoAttendantCallableEntity -Identity `$Target -Type SharedVoicemail -EnableTranscription")
                $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['Action'] = 'TransferCallToTarget'")
                $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['CallTarget'] = `$CallableEntity")

            }
            else {
                if (![string]::IsNullOrEmpty($TextPrompt)) {
                    $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['$TextPromptParamName'] = '$TextPrompt'")
                }
                elseif (![string]::IsNullOrEmpty($AudioFilePromptLocation)) {
                    AddFileImportScript -ApplicationId HuntGroup -StoredLocation $AudioFilePromptLocation -FileName $AudioFilePromptOriginalName -CommandText $CommandText -WarningStrings $WarningStrings -Prepend $Prepend
                    $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['$AudioPromptParamName'] = `"`$(`$FileId.Id)`"")
                }
                else {
                    $null = $WarningStrings.AppendLine("$FlowName will to transfer to $([regex]::replace($URI,'^[Ss][Ii][Pp]:','')) on $ActionName. No Prompt information was found, so a sample Text-To-Speech prompt was generated.")
                    $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['$TextPromptParamName'] = 'Please leave a message.'")
                }
                $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['$ActionName'] = 'Voicemail'")
                $null = $CommandText.AppendLine("$Prepend    `$$CommandHashName['$ActionTargetName'] = `"`$Target`"")
            }
            $null = $CommandText.AppendLine("$Prepend}")
        }
        default {
            if ([string]::IsNullOrEmpty($Action)) {
                $null = $WarningStrings.AppendLine("$FlowName had no action assigned. The $ActionName will be set to disconnect")
            }
            else {
                $null = $WarningStrings.AppendLine("$FlowName attempted to $Action. This is not supported in this script, the $ActionName will be set to disconnect")
            }
            WriteDisconnectAction -CommandText $CommandText -CommandHashName $CommandHashName -ActionName $ActionName -AAMenuOption:$AAMenuOption -NumTabs 1 -Prepend $Prepend
        }
    }
    $WarningStrings.ToString()
}

function AddFileImportScript {
    param (
        [ValidateSet('HuntGroup', 'OrgAutoAttendant')]
        [string]
        $ApplicationId = 'HuntGroup',
        $StoredLocation,
        $FileName,
        [Text.StringBuilder]
        $CommandText,
        [Text.StringBuilder]
        $WarningStrings,
        [string]
        $Prepend = ''
    )
    $FilePath = Get-ChildItem -Path $StoredLocation
    $AudioFilePath = [IO.Path]::Combine($GeneratedScriptsPath, 'AudioFiles')
    $SavedFile = [IO.Path]::Combine($AudioFilePath, $FilePath.Name)
    if (!(Test-Path -Path $SavedFile)) {
        if (!(Test-Path -Path $AudioFilePath)) {
            $null = New-Item -Path $AudioFilePath -ItemType Directory
        }
        Copy-Item -Path $FilePath.FullName -Destination $AudioFilePath
    }
    $null = $WarningStrings.AppendLine("Ensure this file exists in this relative path prior to execution: .\AudioFiles\$([IO.Path]::GetFileName($SavedFile))")
    $null = $CommandText.AppendLine("$Prepend`$FilePath = [IO.Path]::Combine(`$PSScriptRoot, 'AudioFiles', '$([IO.Path]::GetFileName($SavedFile))')")
    $null = $CommandText.AppendLine("$Prepend`$FileBytes = [IO.File]::ReadAllBytes(`$FilePath)")
    $null = $CommandText.AppendLine("$Prepend`$FileId = Import-CsOnlineAudioFile -ApplicationId $ApplicationId -FileName '$FileName' -Content `$FileBytes")
}

function RoundTimeSpan ([TimeSpan] $ts) {
    $Days = $ts.Days
    $Hours = $ts.Hours
    $Minutes = $ts.Minutes

    $Minutes = 15 * [int][Math]::Round($Minutes / 15.0)
    if ($Minutes -eq 60) {
        $Hours += 1
        $Minutes = 0
    }
    if ($Hours -eq 24) {
        $Days += 1
        $Hours = 0
    }

    [TimeSpan]::new($Days, $Hours, $Minutes, 0).ToString()
}

function ConvertFrom-BusinessHoursToTimeRange {
    param (
        $Hours1,
        $Hours2
    )
    $ConfigString = '@('
    if ($null -ne $Hours1) {
        $ConfigString += "(New-CsOnlineTimeRange -Start '$(RoundTimeSpan $Hours1.OpenTime)' -End '$(RoundTimeSpan $Hours1.CloseTime)')"
    }
    if ($null -ne $Hours2) {
        if ($ConfigString.Length -gt 2) {
            $ConfigString += ",$([Environment]::NewLine)"
        }
        $ConfigString += "(New-CsOnlineTimeRange -Start '$(RoundTimeSpan $Hours2.OpenTime)' -End '$(RoundTimeSpan $Hours2.CloseTime)')"
    }
    $ConfigString += ')'
    $ConfigString
}

function CleanName ($Name, $LineURI) {
    $Name = $Name.Trim()
    $Name = [regex]::Replace($Name, 'RGS', '')
    $Name = [regex]::Replace($Name, '[^a-zA-Z0-9\-\(\)\s_]', '')
    $Name = [regex]::Replace($Name, '[\-_]', ' ')
    # delete extra spaces
    $Name = [regex]::Replace($Name, '\s+', ' ')
    $Name = [regex]::Replace($Name, 'queue$', '', 'IgnoreCase')
    # shorten longer names
    $MaxLength = 64 - 3
    if ([string]::IsNullOrWhiteSpace($LineURI)) { $MaxLength -= ($LineURI.Length + 3) }
    $Name = $Name.substring(0, [System.Math]::Min($MaxLength, $Name.Length))

    # Remove spaces at beginning and end and Add LineURI to the end to be used for the DisplayName
    if (![string]::IsNullOrWhiteSpace($LineURI)) {
        $Name.Trim() + ' (' + $LineURI + ')'
    }
    else {
        $Name.Trim()
    }
}
function HashTableToDeclareString ([hashtable]$Hashtable, $VariableName = 'HashTable', $Prepend = "") {
    # only handles string, int or bool types, or collections of those
    $MaxLength = $Hashtable.Keys.Length | Sort-Object -Descending | Select-Object -First 1
    if ($null -eq $MaxLength) { $MaxLength = 1 }
    $sb = [Text.StringBuilder]::new()
    foreach ($kv in $Hashtable.GetEnumerator()) {
        $Value = if ($null -eq $kv.Value) {
            "`$null"
        }
        else {
            switch -regex ($kv.Value.GetType().ToString().ToLower()) {
                "^system\.object\[\]$" {
                    $Nested = @('@(')
                    $Nested += foreach ($v in $kv.Value) {
                        switch -regex ($v.GetType().ToString().ToLower()) {
                            "int\d*$" {
                                $v
                                break
                            }
                            'bool' {
                                "`$$($v.ToString().ToLower())"
                            }
                            default {
                                "'$($v)'"
                                break
                            }
                        }
                    }
                    $Nested += ')'
                    [string]::Join([Environment]::NewLine, $Nested)
                    break
                }
                "int\d*$" {
                    $kv.Value
                    break
                }
                'bool' {
                    "`$$($kv.Value.ToString().ToLower())"
                }
                default {
                    # "'$($kv.Value)'"
                    $v = $kv.Value.ToString()
                    if ($v.IndexOfAny(@(' ', '"', "'")) -gt -1 -or !$v.StartsWith('$')) {
                        "'$($v -replace "'","''")'"
                    }
                    else {
                        $v
                    }
                    break
                }
            }
        }
        $null = $sb.AppendLine("$Prepend`$$VariableName['$($kv.Key)'] = $Value")
    }
    $sb.ToString()
}
function GetSoundPath($AudioFilePrompt, $InstanceId) {
    if ($null -ne $AudioFilePrompt) {
        $SoundUnique = $AudioFilePrompt.UniqueName
        $SoundFile = $AudioFilePrompt.OriginalFileName
        $actualFileName = $SoundUnique + [IO.Path]::GetExtension($SoundFile)
        $SoundPath = $RGSExportPath + '\*\RgsImportExport\RGS\Instances\' + $instanceID + '\'
        $SoundPath = (Resolve-Path $SoundPath).Path

        $FilePath = Get-ChildItem -Path $SoundPath -Recurse -Filter $actualFileName
        $SoundPath = if ($null -eq $FilePath) {
            $actualFileName = $SoundUnique + '.wav'
            $FilePath = Get-ChildItem -Path $SoundPath -Recurse -Filter $actualFileName
            if ($null -eq $FilePath) {
                Write-Warning "Instance ID: $instanceID, Sound File: $SoundFile, or Sound Unique: $SoundUnique not found! Unable to locate file"
                ''
            }
            else {
                $FilePath.FullName
            }
        }
        else {
            $FilePath.FullName
        }
        $SoundPath
    }
    else {
        ''
    }
}
function ProcessAnswer {
    param(
        $Answer,
        $Prefix
    )
    switch ($Tier2_DefaultAnswer.DtmfResponse) {
        '#' {
            $Option = ($Prefix + 'Pound')
        }
        '*' {
            $Option = ($Prefix + 'Star')
        }
        default {
            $Option = ($Prefix + $Answer.DtmfResponse)
        }
    }

    switch ($Answer.Action.Action) {
        'Terminate' {
            $InsertHash["$Option"] = $null
        }
        'TransferToQueue' {
            $TransferToQueue = $Answer.Action.QueueID.InstanceID.Guid
            $InsertHash["$Option"] = $(if ([string]::IsNullOrEmpty($Answer.Action.QueueID)) { $null } else { $TransferToQueue })
        }
        'TransferToQuestion' {
            $InsertHash["$Option"] = $(if ([string]::IsNullOrEmpty($Answer.Action.Question)) { $null } else { 'TransferToQuestion' })
        }
        'TransferToUri' {
            $InsertHash["$Option"] = $(if ([string]::IsNullOrEmpty($Answer.Action.URI)) { $null } else { $Answer.Action.URI })
        }
        'TransferToVoicemailUri' {
            $InsertHash["$Option"] = $(if ([string]::IsNullOrEmpty($Answer.Action.URI)) { $null } else { $Answer.Action.URI })
        }
        'TransferToPstn' {
            $InsertHash["$Option"] = $(if ([string]::IsNullOrEmpty($Answer.Action.URI)) { $null } else { $Answer.Action.URI })
        }
        default {
            $InsertHash["$Option"] = $null
        }
    }
}

$ProcessedWorkflows = @(foreach ($ThisFlow in $WorkFlows) {
        $InsertHash = @{}
        $InsertHash['Identity'] = $(if ([string]::IsNullOrEmpty($ThisFlow.Identity)) { $null } else { $ThisFlow.Identity.InstanceId.Guid })

        $SoundPath = GetSoundPath ($ThisFlow).NonBusinessHoursAction.Prompt.AudioFilePrompt $ThisFlow.Identity.InstanceID.Guid
        $InsertHash['NonBusinessHoursActionPromptAudioFilePromptStoredLocation'] = $(if ([string]::IsNullOrEmpty($SoundPath)) { $null } else { $SoundPath })
        $InsertHash['NonBusinessHoursActionPromptAudioFilePromptOriginalFileName'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).NonBusinessHoursAction.Prompt.AudioFilePrompt.OriginalFileName)) { $null } else { ($ThisFlow).NonBusinessHoursAction.Prompt.AudioFilePrompt.OriginalFileName })
        $InsertHash['NonBusinessHoursActionPromptTextToSpeechPrompt'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).NonBusinessHoursAction.Prompt.TextToSpeechPrompt)) { $null } else { ($ThisFlow).NonBusinessHoursAction.Prompt.TextToSpeechPrompt })
        $InsertHash['NonBusinessHoursActionQuestion'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).NonBusinessHoursAction.Question)) { $null } else { ($ThisFlow).NonBusinessHoursAction.Question })
        $InsertHash['NonBusinessHoursActionAction'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).NonBusinessHoursAction.Action)) { $null } else { ($ThisFlow).NonBusinessHoursAction.Action })
        $InsertHash['NonBusinessHoursActionQueueID'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).NonBusinessHoursAction.QueueID)) { $null } else { ($ThisFlow).NonBusinessHoursAction.QueueID.InstanceID.Guid })
        $InsertHash['NonBusinessHoursActionURI'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).NonBusinessHoursAction.URI)) { $null } else { ($ThisFlow).NonBusinessHoursAction.URI })

        $SoundPath = GetSoundPath ($ThisFlow).HolidayAction.Prompt.AudioFilePrompt $ThisFlow.Identity.InstanceID.Guid
        $InsertHash['HolidayActionPromptAudioFilePromptStoredLocation'] = $(if ([string]::IsNullOrEmpty($SoundPath)) { $null } else { $SoundPath })
        $InsertHash['HolidayActionPromptAudioFilePromptOriginalFileName'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).HolidayAction.Prompt.AudioFilePrompt.OriginalFileName)) { $null } else { ($ThisFlow).HolidayAction.Prompt.AudioFilePrompt.OriginalFileName })
        $InsertHash['HolidayActionPromptTextToSpeechPrompt'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).HolidayAction.Prompt.TextToSpeechPrompt)) { $null } else { ($ThisFlow).HolidayAction.Prompt.TextToSpeechPrompt })
        $InsertHash['HolidayActionQuestion'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).HolidayAction.Question)) { $null } else { ($ThisFlow).HolidayAction.Question })
        $InsertHash['HolidayActionAction'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).HolidayAction.Action)) { $null } else { ($ThisFlow).HolidayAction.Action })
        $InsertHash['HolidayActionQueueID'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).HolidayAction.QueueID)) { $null } else { ($ThisFlow).HolidayAction.QueueID.InstanceID.Guid })
        $InsertHash['HolidayActionURI'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).HolidayAction.URI)) { $null } else { ($ThisFlow).HolidayAction.URI })

        $SoundPath = GetSoundPath ($ThisFlow).DefaultAction.Prompt.AudioFilePrompt $ThisFlow.Identity.InstanceID.Guid
        $InsertHash['DefaultActionPromptAudioFilePromptStoredLocation'] = $(if ([string]::IsNullOrEmpty($SoundPath)) { $null } else { $SoundPath })
        $InsertHash['DefaultActionPromptAudioFilePromptOriginalFileName'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).DefaultAction.Prompt.AudioFilePrompt.OriginalFileName)) { $null } else { ($ThisFlow).DefaultAction.Prompt.AudioFilePrompt.OriginalFileName })
        $InsertHash['DefaultActionPromptTextToSpeechPrompt'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).DefaultAction.Prompt.TextToSpeechPrompt)) { $null } else { ($ThisFlow).DefaultAction.Prompt.TextToSpeechPrompt })
        $InsertHash['DefaultActionQuestion'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).DefaultAction.Question)) { $null } else { ($ThisFlow).DefaultAction.Question })
        $InsertHash['DefaultActionAction'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).DefaultAction.Action)) { $null } else { ($ThisFlow).DefaultAction.Action })
        $InsertHash['DefaultActionQueueID'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).DefaultAction.QueueID)) { $null } else { ($ThisFlow).DefaultAction.QueueID.InstanceID.Guid })
        $InsertHash['DefaultActionURI'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).DefaultAction.URI)) { $null } else { ($ThisFlow).DefaultAction.URI })

        $SoundPath = GetSoundPath ($ThisFlow).CustomMusicOnHoldFile $ThisFlow.Identity.InstanceID.Guid
        $InsertHash['CustomMusicOnHoldStoredLocation'] = $(if ([string]::IsNullOrEmpty($SoundPath)) { $null } else { $SoundPath })
        $InsertHash['CustomMusicOnHoldFileName'] = $(if ([string]::IsNullOrEmpty($ThisFlow.CustomMusicOnHoldFile.OriginalFileName)) { $null } else { $ThisFlow.CustomMusicOnHoldFile.OriginalFileName })

        $InsertHash['Name'] = $(if ([string]::IsNullOrEmpty($ThisFlow.Name)) { $null } else { $ThisFlow.Name })
        $InsertHash['Description'] = $(if ([string]::IsNullOrEmpty($ThisFlow.Description)) { $null } else { $ThisFlow.Description })
        $InsertHash['PrimaryUri'] = $(if ([string]::IsNullOrEmpty($ThisFlow.PrimaryUri)) { $null } else { $ThisFlow.PrimaryUri })
        $InsertHash['Active'] = $(if ([string]::IsNullOrEmpty($ThisFlow.Active)) { $null } else { $ThisFlow.Active })
        $InsertHash['Language'] = $(if ([string]::IsNullOrEmpty($ThisFlow.Language)) { $null } else { $ThisFlow.Language })
        $InsertHash['TimeZone'] = $(if ([string]::IsNullOrEmpty($ThisFlow.TimeZone)) { $null } else { $ThisFlow.TimeZone })
        $InsertHash['Anonymous'] = $(if ([string]::IsNullOrEmpty($ThisFlow.Anonymous)) { $null } else { $ThisFlow.Anonymous })
        $InsertHash['Managed'] = $(if ([string]::IsNullOrEmpty($ThisFlow.Managed)) { $null } else { $ThisFlow.Managed })
        $InsertHash['OwnerPool'] = $(if ([string]::IsNullOrEmpty($ThisFlow.OwnerPool)) { $null } else { $ThisFlow.OwnerPool })
        $InsertHash['DisplayNumber'] = $(if ([string]::IsNullOrEmpty($ThisFlow.DisplayNumber)) { $null } else { $ThisFlow.DisplayNumber })
        $InsertHash['EnabledForFederation'] = $(if ([string]::IsNullOrEmpty($ThisFlow.EnabledForFederation)) { $null } else { $ThisFlow.EnabledForFederation })
        $InsertHash['LineUri'] = $(if ([string]::IsNullOrEmpty($ThisFlow.LineUri)) { $null } else { $ThisFlow.LineUri })
        $InsertHash['BusinessHoursID'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).BusinessHoursID)) { $null } else { ($ThisFlow).BusinessHoursID.InstanceID.Guid })
        $InsertHash['HolidayHoursIDs'] = $(if ([string]::IsNullOrEmpty(($ThisFlow).HolidaySetIDList)) { $null } else { ($ThisFlow).HolidaySetIDList.InstanceID.Guid })
        $InsertHash['ManagersByUri'] = $(if ([string]::IsNullOrEmpty($ThisFlow.ManagersByUri)) { $null } else { $ThisFlow.ManagersByUri })

        [PSCustomObject]$InsertHash
    })

$ProcessedQueues = @(foreach ($ThisQueue in $Queues) {
        $InsertHash = @{}
        $WarningStrings = [Text.StringBuilder]::new()
        $ThisQueueID = $ThisQueue.Identity.InstanceID.Guid
        $ThisQueueIDList = if ($null -ne $ThisQueue.AgentGroupIDList[0]) {
            $ThisQueue.AgentGroupIDList[0].InstanceID.Guid
        }
        else {
            ''
        }
        if ($ThisQueue.AgentGroupIDList.Count -gt 1) {
            $null = $WarningStrings.AppendLine("$($ThisQueue.Name) has multiple agent groups, only the first group will be included")
        }

        $TimeOut = ($ThisQueue).TimeoutAction
        $Overflow = ($ThisQueue).OverflowAction

        $InsertHash['Identity'] = $(if ([string]::IsNullOrEmpty($ThisQueueID)) { $null } else { $ThisQueueID })
        $InsertHash['Name'] = $(if ([string]::IsNullOrEmpty($ThisQueue.Name)) { $null } else { $ThisQueue.Name })
        $InsertHash['Description'] = $(if ([string]::IsNullOrEmpty($ThisQueue.Description)) { $null } else { $ThisQueue.Description })
        $InsertHash['TimeoutThreshold'] = $(if ([string]::IsNullOrEmpty($ThisQueue.TimeoutThreshold)) { $null } else { $ThisQueue.TimeoutThreshold })
        $InsertHash['OverflowThreshold'] = $(if ([string]::IsNullOrEmpty($ThisQueue.OverflowThreshold)) { $null } else { $ThisQueue.OverflowThreshold })
        $InsertHash['OverflowCandidate'] = $(if ([string]::IsNullOrEmpty($ThisQueue.OverflowCandidate)) { $null } else { $ThisQueue.OverflowCandidate })
        $InsertHash['OwnerPool'] = $(if ([string]::IsNullOrEmpty($ThisQueue.OwnerPool)) { $null } else { $ThisQueue.OwnerPool })
        $InsertHash['AgentGroupIDList'] = $(if ([string]::IsNullOrEmpty($ThisQueueIDList)) { $null } else { $ThisQueueIDList })

        $SoundPath = GetSoundPath $TimeOut.Prompt.AudioFilePrompt $ThisQueue.Identity.InstanceID.Guid
        if (![string]::IsNullOrEmpty($SoundPath)) { Write-Host "CQ $($ThisFlow.Name) TimeoutPrompt $SoundPath" }
        $InsertHash['TimeoutAudioStoredLocation'] = $(if ([string]::IsNullOrEmpty($SoundPath)) { $null } else { $SoundPath })
        $InsertHash['TimeoutAudioOriginalName'] = $(if ([string]::IsNullOrEmpty($TimeOut.Prompt.AudioFilePrompt.OriginalName)) { $null } else { $TimeOut.prompt.AudioFilePrompt.OriginalName })

        $InsertHash['TimeoutTextPrompt'] = $(if ([string]::IsNullOrEmpty($TimeOut.Prompt.TextFilePrompt)) { $null } else { $TimeOut.Prompt.TextFilePrompt })
        $InsertHash['TimeoutQuestion'] = $(if ([string]::IsNullOrEmpty($TimeOut.Question)) { $null } else { $TimeOut.Question })
        $InsertHash['TimeoutAction'] = $(if ([string]::IsNullOrEmpty($TimeOut.Action)) { $null } else { $TimeOut.Action })
        $InsertHash['TimeoutQueueID'] = $(if ([string]::IsNullOrEmpty($TimeOut.QueueID)) { $null } else { $TimeOut.QueueID.InstanceID.Guid })
        $InsertHash['TimeoutUri'] = $(if ([string]::IsNullOrEmpty($TimeOut.Uri)) { $null } else { $TimeOut.Uri })


        $SoundPath = GetSoundPath $Overflow.Prompt.AudioFilePrompt $ThisQueue.Identity.InstanceID.Guid
        if (![string]::IsNullOrEmpty($SoundPath)) { Write-Host "CQ $($ThisFlow.Name) OverflowPrompt $SoundPath" }
        $InsertHash['OverflowAudioStoredLocation'] = $(if ([string]::IsNullOrEmpty($SoundPath)) { $null } else { $SoundPath })
        $InsertHash['OverflowAudioOriginalName'] = $(if ([string]::IsNullOrEmpty($Overflow.prompt.AudioFilePrompt.OriginalName)) { $null } else { $Overflow.prompt.AudioFilePrompt.OriginalName })

        $InsertHash['OverflowTextPrompt'] = $(if ([string]::IsNullOrEmpty($Overflow.Prompt.TextFilePrompt)) { $null } else { $Overflow.Prompt.TextFilePrompt })
        $InsertHash['OverflowQuestion'] = $(if ([string]::IsNullOrEmpty($Overflow.Question)) { $null } else { $Overflow.Question })
        $InsertHash['OverflowAction'] = $(if ([string]::IsNullOrEmpty($Overflow.Action)) { $null } else { $Overflow.Action })
        $InsertHash['OverflowQueueID'] = $(if ([string]::IsNullOrEmpty($Overflow.QueueID)) { $null } else { $Overflow.QueueID.InstanceID.Guid })
        $InsertHash['OverflowUri'] = $(if ([string]::IsNullOrEmpty($Overflow.URI)) { $null } else { $Overflow.URI })
        $InsertHash['Warnings'] = $WarningStrings.ToString()
        [PSCustomObject]$InsertHash
    })

$IVRs = $ProcessedWorkflows.Where( { $_.PSObject.Properties.Value.Contains('TransferToQuestion') }).ForEach('Identity')
$ProcessedIVRs = @(foreach ($RGS in $IVRs) {
        $ThisFlow = $WorkFlows.Where( { $_.Identity.InstanceID.Guid -eq $RGS })[0]

        $InsertHash = @{}
        $ThisFlowID = $RGS
        $ThisFlowName = $ThisFlow.Name
        ##### Generic write back
        $InsertHash['Identity'] = $(if ([string]::IsNullOrEmpty($ThisFlowID)) { $null } else { $ThisFlowID })
        $InsertHash['Name'] = $(if ([string]::IsNullOrEmpty($ThisFlowName)) { $null } else { $ThisFlowName })
        ##### Default write-backs

        $SoundPath = GetSoundPath $ThisFlow.DefaultAction.Question.Prompt.AudioFilePrompt $ThisFlow.Identity.InstanceID.Guid
        $InsertHash['DefaultAudioStoredLocation'] = $(if ([string]::IsNullOrEmpty($SoundPath)) { $null } else { $SoundPath })

        $ThisFlowDefaultAudioOriginalName = $ThisFlow.DefaultAction.Question.Prompt.AudioFilePrompt.OriginalFileName
        $InsertHash['DefaultAudioOriginalName'] = $(if ([string]::IsNullOrEmpty($ThisFlowDefaultAudioOriginalName)) { $null } else { $ThisFlowDefaultAudioOriginalName })
        $ThisFlowDefaultTextToSpeech = $ThisFlow.DefaultAction.Question.Prompt.TextToSpeechPrompt
        $InsertHash['DefaultTextToSpeech'] = $(if ([string]::IsNullOrEmpty($ThisFlowDefaultTextToSpeech)) { $null } else { $ThisFlowDefaultTextToSpeech })
        $ThisFlowDefaultInvalidAnswer = $ThisFlow.DefaultAction.Question.InvalidAnswerPrompt
        $InsertHash['DefaultInvalidAnswer'] = $(if ([string]::IsNullOrEmpty($ThisFlowDefaultInvalidAnswer)) { $null } else { $ThisFlowDefaultInvalidAnswer })
        $ThisFlowDefaultNoAnswer = $ThisFlow.DefaultAction.Question.NoAnswerPrompt
        $InsertHash['DefaultNoAnswer'] = $(if ([string]::IsNullOrEmpty($ThisFlowDefaultNoAnswer)) { $null } else { $ThisFlowDefaultNoAnswer })
        $ThisFlowDefaultName = $ThisFlow.DefaultAction.Question.Name
        $InsertHash['DefaultName'] = $(if ([string]::IsNullOrEmpty($ThisFlowDefaultName)) { $null } else { $ThisFlowDefaultName })

        ##### Nulling out the options for write back
        $InsertHash['DefaultOpt0'] = $null
        $InsertHash['DefaultOpt1'] = $null
        $InsertHash['DefaultOpt2'] = $null
        $InsertHash['DefaultOpt3'] = $null
        $InsertHash['DefaultOpt4'] = $null
        $InsertHash['DefaultOpt5'] = $null
        $InsertHash['DefaultOpt6'] = $null
        $InsertHash['DefaultOpt7'] = $null
        $InsertHash['DefaultOpt8'] = $null
        $InsertHash['DefaultOpt9'] = $null
        $InsertHash['DefaultOptPound'] = $null
        $InsertHash['DefaultOptStar'] = $null

        ##### Go through the answer list in case it exists and overwriting the above nulled default values
        if (![string]::IsNullOrEmpty($ThisFlow.DefaultAction.Question.AnswerList)) {
            foreach ($DefaultAnswer in $ThisFlow.DefaultAction.Question.AnswerList) {
                ProcessAnswer $DefaultAnswer 'DefaultOpt'
            }
        }

        ##### NonBus write-backs
        $SoundPath = GetSoundPath $ThisFlow.NonBusinessHoursAction.Question.Prompt.AudioFilePrompt $ThisFlow.Identity.InstanceID.Guid
        $InsertHash['NonBusAudioStoredLocation'] = $(if ([string]::IsNullOrEmpty($SoundPath)) { $null } else { $SoundPath })

        $ThisFlowNonBusAudioOriginalName = $ThisFlow.NonBusinessHoursAction.Question.Prompt.AudioFilePrompt.OriginalFileName
        $InsertHash['NonBusAudioOriginalName'] = $(if ([string]::IsNullOrEmpty($ThisFlowNonBusAudioOriginalName)) { $null } else { $ThisFlowNonBusAudioOriginalName })
        $ThisFlowNonBusTextToSpeech = $ThisFlow.NonBusinessHoursAction.Question.Prompt.TextToSpeechPrompt
        $InsertHash['NonBusTextToSpeech'] = $(if ([string]::IsNullOrEmpty($ThisFlowNonBusTextToSpeech)) { $null } else { $ThisFlowNonBusTextToSpeech })
        $ThisFlowNonBusInvalidAnswer = $ThisFlow.NonBusinessHoursAction.Question.InvalidAnswerPrompt
        $InsertHash['NonBusInvalidAnswer'] = $(if ([string]::IsNullOrEmpty($ThisFlowNonBusInvalidAnswer)) { $null } else { $ThisFlowNonBusInvalidAnswer })
        $ThisFlowNonBusNoAnswer = $ThisFlow.NonBusinessHoursAction.Question.NoAnswerPrompt
        $InsertHash['NonBusNoAnswer'] = $(if ([string]::IsNullOrEmpty($ThisFlowNonBusNoAnswer)) { $null } else { $ThisFlowNonBusNoAnswer })
        $ThisFlowNonBusName = $ThisFlow.NonBusinessHoursAction.Question.Name
        $InsertHash['NonBusName'] = $(if ([string]::IsNullOrEmpty($ThisFlowNonBusName)) { $null } else { $ThisFlowNonBusName })

        ##### Nulling out the options for write back
        $InsertHash['NonBusOpt0'] = $null
        $InsertHash['NonBusOpt1'] = $null
        $InsertHash['NonBusOpt2'] = $null
        $InsertHash['NonBusOpt3'] = $null
        $InsertHash['NonBusOpt4'] = $null
        $InsertHash['NonBusOpt5'] = $null
        $InsertHash['NonBusOpt6'] = $null
        $InsertHash['NonBusOpt7'] = $null
        $InsertHash['NonBusOpt8'] = $null
        $InsertHash['NonBusOpt9'] = $null
        $InsertHash['NonBusOptPound'] = $null
        $InsertHash['NonBusOptStar'] = $null

        ##### Go through the answer list in case it exists and overwriting the above nulled default values
        if (![string]::IsNullOrEmpty($ThisFlow.NonBusinessHoursAction.Question.AnswerList)) {
            foreach ($NonBusAnswer in $ThisFlow.NonBusinessHoursAction.Question.AnswerList) {
                ProcessAnswer $NonBusAnswer 'NonBusOpt'
            }
        }

        ##### Holiday write-backs
        $SoundPath = GetSoundPath $ThisFlow.HolidayAction.Question.Prompt.AudioFilePrompt $ThisFlow.Identity.InstanceID.Guid
        $InsertHash['HolidayAudioStoredLocation'] = $(if ([string]::IsNullOrEmpty($SoundPath)) { $null } else { $SoundPath })

        $ThisFlowHolidayAudioOriginalName = $ThisFlow.HolidayAction.Question.Prompt.AudioFilePrompt.OriginalFileName
        $InsertHash['HolidayAudioOriginalName'] = $(if ([string]::IsNullOrEmpty($ThisFlowHolidayAudioOriginalName)) { $null } else { $ThisFlowHolidayAudioOriginalName })
        $ThisFlowHolidayTextToSpeech = $ThisFlow.HolidayAction.Question.Prompt.TextToSpeechPrompt
        $InsertHash['HolidayTextToSpeech'] = $(if ([string]::IsNullOrEmpty($ThisFlowHolidayTextToSpeech)) { $null } else { $ThisFlowHolidayTextToSpeech })
        $ThisFlowHolidayInvalidAnswer = $ThisFlow.HolidayAction.Question.InvalidAnswerPrompt
        $InsertHash['HolidayInvalidAnswer'] = $(if ([string]::IsNullOrEmpty($ThisFlowHolidayInvalidAnswer)) { $null } else { $ThisFlowHolidayInvalidAnswer })
        $ThisFlowHolidayNoAnswer = $ThisFlow.HolidayAction.Question.NoAnswerPrompt
        $InsertHash['HolidayNoAnswer'] = $(if ([string]::IsNullOrEmpty($ThisFlowHolidayNoAnswer)) { $null } else { $ThisFlowHolidayNoAnswer })
        $ThisFlowHolidayName = $ThisFlow.HolidayAction.Question.Name
        $InsertHash['HolidayName'] = $(if ([string]::IsNullOrEmpty($ThisFlowHolidayName)) { $null } else { $ThisFlowHolidayName })

        ##### Nulling out the options for write back
        $InsertHash['HolidayOpt0'] = $null
        $InsertHash['HolidayOpt1'] = $null
        $InsertHash['HolidayOpt2'] = $null
        $InsertHash['HolidayOpt3'] = $null
        $InsertHash['HolidayOpt4'] = $null
        $InsertHash['HolidayOpt5'] = $null
        $InsertHash['HolidayOpt6'] = $null
        $InsertHash['HolidayOpt7'] = $null
        $InsertHash['HolidayOpt8'] = $null
        $InsertHash['HolidayOpt9'] = $null
        $InsertHash['HolidayOptPound'] = $null
        $InsertHash['HolidayOptStar'] = $null

        ##### Go through the answer list in case it exists and overwriting the above nulled default values
        if (![string]::IsNullOrEmpty($ThisFlow.HolidayAction.Question.AnswerList)) {
            foreach ($HolidayAnswer in $ThisFlow.HolidayAction.Question.AnswerList) {
                ProcessAnswer $HolidayAnswer 'HolidayOpt'
            }
        }

        [PSCustomObject]$InsertHash
    })

$Tier2IVRs = $ProcessedIVRs.Where( { $_.PSObject.Properties.Value.Contains('TransferToQuestion') }).ForEach('Identity')
$Tier2ProcessedIVRs = @(foreach ($RGS in $Tier2IVRs) {
        $InsertHash = @{}

        $ThisFlow = $WorkFlows | Where-Object { $_.Identity.InstanceID.Guid -eq $RGS }

        $InsertHash['Identity'] = $(if ([string]::IsNullOrEmpty($RGS)) { $null } else { $RGS })
        $InsertHash['Name'] = $(if ([string]::IsNullOrEmpty($ThisFlow.Name)) { $null } else { $ThisFlow.Name })
        $PrimaryAnswerLists = $ThisFlow.DefaultAction.Question.AnswerList.Where( { $_.Action.Action -eq 'TransferToQuestion' })

        if (![string]::IsNullOrEmpty($PrimaryAnswerLists)) {
            foreach ($AnswerList in $PrimaryAnswerLists) {
                ##### Nulling out the options for write back
                $InsertHash['Tier2_DefaultOpt0'] = $null
                $InsertHash['Tier2_DefaultOpt1'] = $null
                $InsertHash['Tier2_DefaultOpt2'] = $null
                $InsertHash['Tier2_DefaultOpt3'] = $null
                $InsertHash['Tier2_DefaultOpt4'] = $null
                $InsertHash['Tier2_DefaultOpt5'] = $null
                $InsertHash['Tier2_DefaultOpt6'] = $null
                $InsertHash['Tier2_DefaultOpt7'] = $null
                $InsertHash['Tier2_DefaultOpt8'] = $null
                $InsertHash['Tier2_DefaultOpt9'] = $null
                $InsertHash['Tier2_DefaultOptPound'] = $null
                $InsertHash['Tier2_DefaultOptStar'] = $null

                ##### write back which Primary DTMFResponse we need to dig into further
                $PrimaryDTMFResponse = 'DefaultAction_' + $AnswerList.DtmfResponse
                $InsertHash['OriginatorOption'] = $(if ([string]::IsNullOrEmpty($PrimaryDTMFResponse)) { $null } else { $PrimaryDTMFResponse })

                ##### getting second tier info for write back
                $SecondAnswerList = $AnswerList.Action.Question.AnswerList

                foreach ($Tier2_DefaultAnswer in $SecondAnswerList) {
                    ProcessAnswer $Tier2_DefaultAnswer 'Tier2_DefaultOpt'
                }
            }
        }
        [PSCustomObject]$InsertHash
    })

$ProcessedAgentGroups = @(foreach ($AgentGroup in $AgentGroups) {
        $InsertHash = @{}
        $WarningStrings = [Text.StringBuilder]::new()
        $InsertHash['Identity'] = $AgentGroup.Identity.InstanceID.Guid
        $InsertHash['Name'] = $AgentGroup.Name
        $InsertHash['Description'] = $AgentGroup.Description
        $InsertHash['ParticipationPolicy'] = $AgentGroup.ParticipationPolicy
        $InsertHash['AgentAlertTime'] = $AgentGroup.AgentAlertTime
        $InsertHash['RoutingMethod'] = $AgentGroup.RoutingMethod
        $InsertHash['DistributionGroupAddress'] = $AgentGroup.DistributionGroupAddress
        $InsertHash['OwnerPool'] = $AgentGroup.OwnerPool
        $InsertHash['AgentsByUri'] = $AgentGroup.AgentsByUri.AbsolutePath
        if ($null -ne $AgentGroup.DistributionGroupAddress) {
            $null = $WarningStrings.AppendLine("$($AgentGroup.Name) uses DistributionGroup $($AgentGroup.DistributionGroupAddress -join ',')")
            $null = $WarningStrings.AppendLine('Ensure the following commands return valid values:')
            foreach ($dg in $AgentGroup.DistributionGroupAddress) {
                $null = $WarningStrings.AppendLine("    Find-CsGroup -SearchQuery '$dg' -ExactMatchOnly `$true")
            }
        }
        $InsertHash['Warnings'] = $WarningStrings.ToString()
        [PSCustomObject]$InsertHash
    })

$CallQueues = $ProcessedWorkflows.Where( { $_.Identity -notin $ProcessedIVRs.Identity })
foreach ($Workflow in $CallQueues) {
    $CommandText = [Text.StringBuilder]::new()
    $DefaultQueue = $ProcessedQueues.Where( { $_.Identity -eq $Workflow.DefaultActionQueueID }, 'First')[0]
    $WarningStrings = [Text.StringBuilder]::new()
    if (![string]::IsNullOrEmpty($Workflow.Warnings)) {
        $null = $WarningStrings.AppendLine($Workflow.Warnings)
    }
    if ([string]::IsNullOrWhiteSpace($DefaultQueue.Name)) {
        $null = $WarningStrings.AppendLine("Default Queue: $($Workflow.DefaultActionQueueID) for $($Workflow.Name) has no valid name assigned")
    }
    if (![string]::IsNullOrEmpty($DefaultQueue.Warnings)) {
        $null = $WarningStrings.AppendLine($DefaultQueue.Warnings)
    }

    # Define the name based on the available data in the row if on-prem Queue Name fails use the workflow name
    $AAName = if ([string]::IsNullOrEmpty($Workflow.Name)) {
        if ([string]::IsNullOrEmpty($DefaultQueue.Name)) {
            Write-Warning 'Workflow has no valid name'
            continue
        }
        else {
            $DefaultQueue.Name
        }
    }
    else {
        $Workflow.Name
    }
    $CQName = if ([string]::IsNullOrEmpty($DefaultQueue.Name)) {
        if ([string]::IsNullOrEmpty($Workflow.Name)) {
            Write-Warning 'Workflow has no valid name'
            continue
        }
        else {
            $Workflow.Name
        }
    }
    else {
        $DefaultQueue.Name
    }
    $LineURI = [regex]::Replace($Workflow.LineURI, '[Xx]', '')
    $LineURI = [regex]::Replace($LineURI, ';[Ee][Tt]=', 'x')
    $LineURI = [regex]::Replace($LineURI, '[^0-9x]', '')

    $NameLURI = CleanName $CQName $LineURI
    $CQNameLURI = 'CQ ' + $NameLURI
    $NameLURI = CleanName $AAName $LineURI
    $AANameLURI = 'AA ' + $NameLURI
    $AADispName = $NameLURI
    $CQDispName = $CQNameLURI

    $FileName = [IO.Path]::Combine($GeneratedScriptsPath, ($AADispName + '.ps1'))

    $CallQueueParams = @{}
    # Add name to param list
    $CallQueueParams['Name'] = '$CQName'
    $CallQueueParams['ErrorAction'] = 'Stop'

    # TODO: Need to add module logic to a given script

    $null = $CommandText.AppendLine("# Original HuntGroup Name: $($Workflow.Name)")
    $null = $CommandText.AppendLine("#     Original Queue Name: $($DefaultQueue.Name)")
    $null = $CommandText.AppendLine("#             New AA Name: $AADispName")
    $null = $CommandText.AppendLine("#             New CQ Name: $CQDispName")
    $null = $CommandText.AppendLine("#                 LineUri: $LineURI")
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('[CmdletBinding()]')
    $null = $CommandText.AppendLine('param (')
    $null = $CommandText.AppendLine('    [switch]')
    $null = $CommandText.AppendLine('    $GenerateAccountsOnly,')
    $null = $CommandText.AppendLine()
    $null = $CommandText.AppendLine('    [switch]')
    $null = $CommandText.AppendLine('    $AssignLineUri')
    $null = $CommandText.AppendLine(')')
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('$CsvPath = [IO.Path]::Combine($PSScriptRoot,"$([IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)).csv")')
    $null = $CommandText.AppendLine('if (!(Test-Path -Path $CsvPath)) { Write-Warning "Cannot locate config csv at path $CsvPath. Exiting..."; exit }')
    $null = $CommandText.AppendLine('$Config = Import-Csv -Path $CsvPath')
    $null = $CommandText.AppendLine('$AAName = $Config.AADispName')
    $null = $CommandText.AppendLine('$CQName = $Config.CQDispName')
    $null = $CommandText.AppendLine('$CQAccountName = $Config.CQName')
    $null = $CommandText.AppendLine('$AAAccountName = $Config.AAName')
    $null = $CommandText.AppendLine('$DeploymentStatus = $Config.DeploymentStatus')
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('function CreateApplicationInstance {')
    $null = $CommandText.AppendLine('    param(')
    $null = $CommandText.AppendLine('        [string]$AccountName,')
    $null = $CommandText.AppendLine('        [string]$Name,')
    $null = $CommandText.AppendLine('        [switch]$CQ,')
    $null = $CommandText.AppendLine('        [switch]$AA,')
    $null = $CommandText.AppendLine('        [string]$OU')
    $null = $CommandText.AppendLine('    )')
    $null = $CommandText.AppendLine('    $Prefix = if ($CQ) { ''CQ-'' } elseif ($AA) { ''AA-'' }')
    $null = $CommandText.AppendLine('    $ApplicationId = if ($CQ) { ''11cd3e2e-fccb-42ad-ad00-878b93575e07'' } elseif ($AA) { ''ce933385-9390-45d1-9512-c8d228074e07'' }')
    $null = $CommandText.AppendLine('    # Getting a valid UPN for the Application Instance')
    $null = $CommandText.AppendLine('    $UPN = $Prefix + ($AccountName.Trim() -replace ''[^a-zA-Z0-9_\-]'', '''').ToLower()')
    $null = $CommandText.AppendLine("    `$UPN = `$UPN.Substring(0, [System.Math]::Min(20, `$UPN.Length)) + '@$SipDomain'")
    $null = $CommandText.AppendLine('    do {')
    $null = $CommandText.AppendLine('        $UPNExist = try { Get-CsOnlineApplicationInstance -Identity $UPN -ErrorAction SilentlyContinue } catch { $null }')
    $null = $CommandText.AppendLine('        if ([string]::IsNullOrEmpty($UPNExist)) {')
    $null = $CommandText.AppendLine('            $valid = $true')
    $null = $CommandText.AppendLine('        } else {')
    $null = $CommandText.AppendLine('            $valid = $false')
    $null = $CommandText.AppendLine('            $Random = -join ((97..122) | Get-Random -Count 4 | ForEach-Object { [char]$_ })')
    $null = $CommandText.AppendLine('            $UPN = $Prefix + ($AccountName.Trim() -replace ''[^a-zA-Z0-9_\-]'', '''').ToLower()')
    $null = $CommandText.AppendLine("            `$UPN = `$UPN.Substring(0, [System.Math]::Min(16, `$UPN.Length)) + `$Random + '@$SipDomain'")
    $null = $CommandText.AppendLine('        }')
    $null = $CommandText.AppendLine('    } until ($valid)')
    $null = $CommandText.AppendLine('    # Creating the Application Instance')
    $null = $CommandText.AppendLine('    try {')
    $null = $CommandText.AppendLine('        $NewInstance = New-CsOnlineApplicationInstance -UserPrincipalName $UPN -ApplicationId $ApplicationId -DisplayName $Name')
    $null = $CommandText.AppendLine('        $InstanceId = $NewInstance.ObjectID')
    $null = $CommandText.AppendLine('        $InstanceUpn = $NewInstance.UserPrincipalName -replace "^$Prefix",''''')
    $null = $CommandText.AppendLine('    } catch {')
    $null = $CommandText.AppendLine('        Write-Warning "Unable to create application instance for ${Name}... Ending processing."')
    $null = $CommandText.AppendLine('        Write-Warning $_.Exception')
    $null = $CommandText.AppendLine('        exit')
    $null = $CommandText.AppendLine('    }')
    $null = $CommandText.AppendLine('    [PSCustomObject]@{')
    $null = $CommandText.AppendLine('        InstanceId = $InstanceId')
    $null = $CommandText.AppendLine('        InstanceUpn = $InstanceUpn')
    $null = $CommandText.AppendLine('    }')
    $null = $CommandText.AppendLine('}')
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('# Creating the application instances')
    $null = $CommandText.AppendLine('if ($DeploymentStatus.ToLower() -eq ''new'') {')
    $null = $CommandText.Append("    `$CQInstance = CreateApplicationInstance -AccountName `$CQAccountName -Name `$CQName -CQ")
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('    $DeploymentStatus = ''CQAccountDeployed''')
    $null = $CommandText.AppendLine('    $Config.CQName = ($CQInstance.InstanceUpn -split ''@'')[0]')
    $null = $CommandText.AppendLine('    $Config.DeploymentStatus = $DeploymentStatus')
    $null = $CommandText.AppendLine('    $Config | Export-Csv -Path $CsvPath -NoTypeInformation')
    $null = $CommandText.AppendLine('}')

    $null = $CommandText.AppendLine('if ($DeploymentStatus.ToLower() -eq ''cqaccountdeployed'') {')
    $null = $CommandText.Append("    `$AAInstance = CreateApplicationInstance -AccountName `$AAAccountName -Name `$AAName -AA")
    $null = $CommandText.AppendLine()
    $null = $CommandText.AppendLine('    $DeploymentStatus = ''AccountsDeployed''')
    $null = $CommandText.AppendLine('    $Config.AAName = ($AAInstance.InstanceUpn -split ''@'')[0]')
    $null = $CommandText.AppendLine('    $Config.DeploymentStatus = $DeploymentStatus')
    $null = $CommandText.AppendLine('    $Config | Export-Csv -Path $CsvPath -NoTypeInformation')
    $null = $CommandText.AppendLine('}')
    $null = $CommandText.AppendLine('if ($GenerateAccountsOnly) { Write-Warning ''Accounts Created, run script again without GenerateAccountsOnly switch to continue processing.''; exit }')
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('if ($DeploymentStatus.ToLower() -eq ''accountsdeployed'') {')
    $null = $CommandText.AppendLine('    # Waiting the Call Queue and Auto Attendant Application Instances to replicate')
    $null = $CommandText.AppendLine('    do {')
    $null = $CommandText.AppendLine('        $ExistsOnline = $false')
    $null = $CommandText.AppendLine('        try {')
    $null = $CommandText.AppendLine("            `$CQInstance = Get-CsOnlineApplicationInstance -Identity ('CQ-{0}@$SipDomain' -f `$Config.CQName) -ErrorAction Stop")
    $null = $CommandText.AppendLine("            `$AAInstance = Get-CsOnlineApplicationInstance -Identity ('AA-{0}@$SipDomain' -f `$Config.AAName) -ErrorAction Stop")
    $null = $CommandText.AppendLine('            if($null -ne $CQInstance -and $null -ne $AAInstance) {')
    $null = $CommandText.AppendLine('                $ExistsOnline = $true')
    $null = $CommandText.AppendLine('            }')
    $null = $CommandText.AppendLine('        } catch {}')
    $null = $CommandText.AppendLine('    } until ($ExistsOnline)')
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('    $CQInstanceId = $CQInstance.ObjectID')
    $null = $CommandText.AppendLine('    $AAInstanceId = $AAInstance.ObjectID')
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('    # Building the Call Queue')
    $null = $CommandText.AppendLine()
    $null = $CommandText.AppendLine('    $CallQueueParams = @{}')

    $DefaultQueueAgentGroup = $ProcessedAgentGroups.Where( { $_.Identity -eq $DefaultQueue.AgentGroupIDList })[0]
    if (![string]::IsNullOrEmpty($DefaultQueueAgentGroup.Warnings)) {
        $null = $WarningStrings.AppendLine($DefaultQueueAgentGroup.Warnings)
    }

    if ($null -eq $DefaultQueueAgentGroup) {
        Write-Warning "$($DefaultQueue.Name) has no assigned Agent Groups, this will not be processed further."
        continue
    }
    $CallQueueParams['AllowOptOut'] = $DefaultQueueAgentGroup.ParticipationPolicy -eq 'Formal'     # AllowOptOut: Informal -> true, Formal -> false

    $RoundValueParams = @{
        InputTime   = $DefaultQueueAgentGroup.AgentAlertTime
        MinimumTime = 30
        RoundTo     = 15
        MaximumTime = 180
    }
    $CallQueueParams['AgentAlertTime'] = RoundValue @RoundValueParams

    if ($DefaultQueueAgentGroup.DistributionGroupAddress.Count -gt 0) {
        $null = $CommandText.AppendLine('    # Adding Distribution Groups from AgentGroup to the Queue')
        $null = $CommandText.AppendLine('    $DGroups = [Collections.Generic.List[object]]::new()')
        $null = $CommandText.AppendLine('    $DGs = @(')
        foreach ($DList in $DefaultQueueAgentGroup.DistributionGroupAddress) {
            $null = $CommandText.AppendLine("        '$DList'")
        }
        $null = $CommandText.AppendLine('    )')
        $null = $CommandText.AppendLine('    foreach ($Dlist in $DGs) {')
        $null = $CommandText.AppendLine('        $GroupId = try {')
        $null = $CommandText.AppendLine('        (Find-CsGroup -SearchQuery $Dlist -ExactMatchOnly $true -MaxResults 1).Id.Guid')
        $null = $CommandText.AppendLine('        } catch {')
        $null = $CommandText.AppendLine('            $null')
        $null = $CommandText.AppendLine('        }')
        $null = $CommandText.AppendLine('        if ($null -ne $GroupId) {')
        $null = $CommandText.AppendLine('            $DGroups.Add($GroupId) | Out-Null')
        $null = $CommandText.AppendLine('        } else {')
        $null = $CommandText.AppendLine('            Write-Warning "Could not find valid object for $DList, skipping"')
        $null = $CommandText.AppendLine('        }')
        $null = $CommandText.AppendLine('    }')

        $null = $CommandText.AppendLine('    if ($DGroups.Count -gt 0) {')
        $null = $CommandText.AppendLine('        $CallQueueParams[''DistributionLists''] = $DGroups')
        $null = $CommandText.AppendLine('    }')
        $null = $CommandText.AppendLine()
    }

    if ($DefaultQueueAgentGroup.AgentsByUri.Count -gt 0) {
        $null = $CommandText.AppendLine('    # Adding Agent Uris from AgentGroup to the Queue')
        $null = $CommandText.AppendLine('    $AgentsByUri = [Collections.Generic.List[object]]@()')
        $null = $CommandText.AppendLine('    $AgentUris = @(')
        $AgentCount = 1
        foreach ($AgentUri in $DefaultQueueAgentGroup.AgentsByUri) {
            if ($AgentCount -le 20) {
                $null = $CommandText.AppendLine("        '$AgentUri'")
            }
            else {
                if ($AgentCount -eq 21) {
                    $null = $WarningStrings.AppendLine('More than 20 agents assigned to queue, only first 20 will be added.')
                }
                $null = $CommandText.AppendLine("        `# '$AgentUri'")
            }
            $AgentCount++
        }
        $null = $CommandText.AppendLine('    )')
        $null = $CommandText.AppendLine('    foreach ($AgentUri in $AgentUris) {')
        $null = $CommandText.AppendLine('        $Agent = try {')
        $null = $CommandText.AppendLine('            (Get-CsOnlineUser -Identity $AgentUri -ErrorAction Stop).Identity')
        $null = $CommandText.AppendLine('        } catch {')
        $null = $CommandText.AppendLine('            $null')
        $null = $CommandText.AppendLine('        }')
        $null = $CommandText.AppendLine('        if ($null -ne $Agent) { ')
        $null = $CommandText.AppendLine('            $AgentsByUri.Add($Agent)')
        $null = $CommandText.AppendLine('        } else {')
        $null = $CommandText.AppendLine('            Write-Warning "Could not find valid object for $AgentUri, skipping"')
        $null = $CommandText.AppendLine('        }')
        $null = $CommandText.AppendLine('    }')
        $null = $CommandText.AppendLine('    if ($AgentsByUri.Count -gt 0) {')
        $null = $CommandText.AppendLine('        $CallQueueParams[''Users''] = $AgentsByUri')
        $null = $CommandText.AppendLine('    }')
        $null = $CommandText.AppendLine()
    }
    elseif ( $DefaultQueueAgentGroup.DistributionGroupAddress.Count -eq 0 ) {
        Write-Warning "$($DefaultQueue.Name) has no users or distribution groups in its assigned Agent Groups, this will not be processed further."
        continue
    }

    $CallQueueParams['RoutingMethod'] = if ( $DefaultQueueAgentGroup.RoutingMethod -eq 'Parallel' ) {
        'Serial'
    }
    else {
        $DefaultQueueAgentGroup.RoutingMethod
    }

    $null = $CommandText.AppendLine('    # Configuring Overflow Action')
    $ConvertActionParams = @{
        FlowName                    = $DefaultQueue.Name
        URI                         = $DefaultQueue.OverflowUri
        QueueId                     = $DefaultQueue.OverflowQueueID
        Action                      = $DefaultQueue.OverflowAction
        CommandText                 = $CommandText
        CommandHashName             = 'CallQueueParams'
        ActionName                  = 'OverflowAction'
        ActionTargetName            = 'OverflowActionTarget'
        AudioFilePromptLocation     = $DefaultQueue.OverflowAudioStoredLocation
        AudioFilePromptOriginalName = $DefaultQueue.OverflowAudioOriginalName
        TextPrompt                  = $DefaultQueue.OverflowTextPrompt
        AudioPromptParamName        = 'OverflowSharedVoicemailAudioFilePrompt'
        TextPromptParamName         = 'OverflowSharedVoicemailTextToSpeechPrompt'
        Prepend                     = '    '
    }
    $warn = ConvertNonQuestionAction @ConvertActionParams
    if (![string]::IsNullOrEmpty($warn)) {
        $null = $WarningStrings.AppendLine($warn)
    }
    if ($null -ne $DefaultQueue.OverflowThreshold) {
        $RoundValueParams = @{
            InputTime   = $DefaultQueue.OverflowThreshold
            MinimumTime = 0
            RoundTo     = 1
            MaximumTime = 200
        }
        $CallQueueParams['OverflowThreshold'] = RoundValue @RoundValueParams
    }
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('    # Configuring Timeout Action')
    $ConvertActionParams = @{
        FlowName                    = $DefaultQueue.Name
        URI                         = $DefaultQueue.TimeoutUri
        Action                      = $DefaultQueue.TimeoutAction
        QueueId                     = $DefaultQueue.TimeoutQueueID
        CommandText                 = $CommandText
        CommandHashName             = 'CallQueueParams'
        ActionName                  = 'TimeoutAction'
        ActionTargetName            = 'TimeoutActionTarget'
        AudioFilePromptLocation     = $DefaultQueue.TimeoutAudioStoredLocation
        AudioFilePromptOriginalName = $DefaultQueue.TimeoutAudioOriginalName
        TextPrompt                  = $DefaultQueue.TimeoutTextPrompt
        AudioPromptParamName        = 'TimeoutSharedVoicemailAudioFilePrompt'
        TextPromptParamName         = 'TimeoutSharedVoicemailTextToSpeechPrompt'
        Prepend                     = '    '
    }
    $warn = ConvertNonQuestionAction @ConvertActionParams
    if (![string]::IsNullOrEmpty($warn)) {
        $null = $WarningStrings.AppendLine($warn)
    }
    if ($null -ne $DefaultQueue.TimeoutThreshold) {
        $RoundValueParams = @{
            InputTime   = $DefaultQueue.TimeoutThreshold
            MinimumTime = 45
            RoundTo     = 15
            MaximumTime = 2700
        }
        $CallQueueParams['TimeoutThreshold'] = RoundValue @RoundValueParams
    }
    $null = $CommandText.AppendLine()

    if (![string]::IsNullOrEmpty($Workflow.CustomMusicOnHoldStoredLocation)) {
        $null = $CommandText.AppendLine('    # Importing Custom Hold Music audio file')
        AddFileImportScript -ApplicationId HuntGroup -StoredLocation $Workflow.CustomMusicOnHoldStoredLocation -FileName $Workflow.CustomMusicOnHoldFileName -CommandText $CommandText -WarningStrings $WarningStrings -Prepend '    '
        $null = $CommandText.AppendLine()
        $CallQueueParams['MusicOnHoldAudioFileId'] = '$FileId.Id'
    }
    else {
        $CallQueueParams['UseDefaultMusicOnHold'] = $true
    }

    # PresenceBasedRouting (off by default)
    $CallQueueParams['PresenceBasedRouting'] = $true

    # ConferenceMode (off by default)
    # $CallQueueParams['ConferenceMode'] = $true

    # LanguageId handling for sharedvoicemail? we can set this regardless
    # OverflowSharedVoicemailAudioFilePrompt

    $null = $CommandText.AppendLine('    # Adding remaining queue configuration information')
    $CommandString = HashTableToDeclareString -Hashtable $CallQueueParams -VariableName 'CallQueueParams' -Prepend '    '
    $null = $CommandText.AppendLine($CommandString)

    $null = $CommandText.AppendLine('    # Creating the Call Queue')
    $null = $CommandText.AppendLine('    try {')
    $null = $CommandText.AppendLine('        $CallQueue = New-CsCallQueue @CallQueueParams')
    $null = $CommandText.AppendLine('    } catch {')
    $null = $CommandText.AppendLine('        Write-Error "Unable to create call queue $CQAccountName!"')
    $null = $CommandText.AppendLine('        Write-Warning $_.Exception')
    $null = $CommandText.AppendLine('        exit')
    $null = $CommandText.AppendLine('    }')
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('    # Creating Call Queue Application Instance Association')
    $null = $CommandText.AppendLine('    try {')
    $null = $CommandText.AppendLine('        $CQAssoc = New-CsOnlineApplicationInstanceAssociation -ConfigurationType CallQueue -ConfigurationId $($CallQueue.Identity) -Identities $CQInstanceId')
    $null = $CommandText.AppendLine('    } catch {')
    $null = $CommandText.AppendLine('        Write-Warning "Unable to create application association for $CQName ending processing."')
    $null = $CommandText.AppendLine('        Write-Warning $_.Exception')
    $null = $CommandText.AppendLine('        exit')
    $null = $CommandText.AppendLine('    }')
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('    # Building the Auto Attendant')
    $null = $CommandText.AppendLine()
    $null = $CommandText.AppendLine('    # Building the Auto Attendant Default Action')
    $null = $CommandText.AppendLine('    $DefaultAction = New-CsAutoAttendantCallableEntity -Identity $CQInstanceId -Type ApplicationEndpoint')
    $null = $CommandText.AppendLine('    $DefaultMenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Automatic -CallTarget $DefaultAction')
    $null = $CommandText.AppendLine('    $DefaultMenu = New-CsAutoAttendantMenu -Name ''Default Menu'' -MenuOptions @($DefaultMenuOption)')

    if (![string]::IsNullOrWhiteSpace($Workflow.DefaultActionPromptAudioFilePromptStoredLocation)) {
        AddFileImportScript -ApplicationId OrgAutoAttendant -StoredLocation $Workflow.DefaultActionPromptAudioFilePromptStoredLocation -FileName $Workflow.DefaultActionPromptAudioFilePromptOriginalFileName -CommandText $CommandText -WarningStrings $WarningStrings -Prepend '    '
        $null = $CommandText.AppendLine('    $DefaultGreetingPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $FileId')
        $null = $CommandText.AppendLine('    $DefaultCallFlow = New-CsAutoAttendantCallFlow -Name ''Default Call Flow'' -Greetings @($DefaultGreetingPrompt) -Menu $DefaultMenu')
    }
    elseif (![string]::IsNullOrWhiteSpace($Workflow.DefaultActionPromptTextToSpeechPrompt)) {
        $null = $CommandText.AppendLine("    `$DefaultGreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt '$($Workflow.DefaultActionPromptTextToSpeechPrompt)'")
        $null = $CommandText.AppendLine('    $DefaultCallFlow = New-CsAutoAttendantCallFlow -Name ''Default Call Flow'' -Greetings @($DefaultGreetingPrompt) -Menu $DefaultMenu')
    }
    else {
        $null = $CommandText.AppendLine('    $DefaultCallFlow = New-CsAutoAttendantCallFlow -Name ''Default Call Flow'' -Menu $DefaultMenu')
    }

    # Add logic for business hours here
    if ($null -ne $Workflow.BusinessHoursID) {
        $null = $CommandText.AppendLine()
        $null = $CommandText.AppendLine('    # Building the Auto Attendant After Hours Action')
        $BusinessHours = @($HoursOfBusiness).Where( { $_.Identity.InstanceID.Guid -eq $Workflow.BusinessHoursID })[0]
        $HoursName = [regex]::Replace($BusinessHours.Name, '_?[A-Fa-f0-9]{8}(?:-?[A-Fa-f0-9]{4}){3}-?[A-Fa-f0-9]{12}$', '')
        $null = $CommandText.AppendLine("    `$AfterHoursSchedule = @(Get-CsOnlineSchedule).Where({ `$_.Name -eq '$HoursName' })[0]")
        $null = $CommandText.AppendLine('    if ($null -eq $AfterHoursSchedule) {')
        $null = $CommandText.AppendLine('        $OnlineScheduleParams = @{')
        $null = $CommandText.AppendLine("            Name = '$HoursName'")
        $null = $CommandText.AppendLine('            WeeklyRecurrentSchedule = $true')
        $null = $CommandText.AppendLine('            Complement = $true')
        $null = $CommandText.AppendLine('        }')
        if (![string]::IsNullOrEmpty($BusinessHours.MondayHours1) -or ![string]::IsNullOrEmpty($BusinessHours.MondayHours2)) {
            $null = $CommandText.AppendLine("        `$OnlineScheduleParams['MondayHours'] = $(ConvertFrom-BusinessHoursToTimeRange $BusinessHours.MondayHours1 $BusinessHours.MondayHours2)")
        }
        if (![string]::IsNullOrEmpty($BusinessHours.TuesdayHours1) -or ![string]::IsNullOrEmpty($BusinessHours.TuesdayHours2)) {
            $null = $CommandText.AppendLine("        `$OnlineScheduleParams['TuesdayHours'] = $(ConvertFrom-BusinessHoursToTimeRange $BusinessHours.TuesdayHours1 $BusinessHours.TuesdayHours2)")
        }
        if (![string]::IsNullOrEmpty($BusinessHours.WednesdayHours1) -or ![string]::IsNullOrEmpty($BusinessHours.WednesdayHours2)) {
            $null = $CommandText.AppendLine("        `$OnlineScheduleParams['WednesdayHours'] = $(ConvertFrom-BusinessHoursToTimeRange $BusinessHours.WednesdayHours1 $BusinessHours.WednesdayHours2)")
        }
        if (![string]::IsNullOrEmpty($BusinessHours.ThursdayHours1) -or ![string]::IsNullOrEmpty($BusinessHours.ThursdayHours2)) {
            $null = $CommandText.AppendLine("        `$OnlineScheduleParams['ThursdayHours'] = $(ConvertFrom-BusinessHoursToTimeRange $BusinessHours.ThursdayHours1 $BusinessHours.ThursdayHours2)")
        }
        if (![string]::IsNullOrEmpty($BusinessHours.FridayHours1) -or ![string]::IsNullOrEmpty($BusinessHours.FridayHours2)) {
            $null = $CommandText.AppendLine("        `$OnlineScheduleParams['FridayHours'] = $(ConvertFrom-BusinessHoursToTimeRange $BusinessHours.FridayHours1 $BusinessHours.FridayHours2)")
        }
        if (![string]::IsNullOrEmpty($BusinessHours.SaturdayHours1) -or ![string]::IsNullOrEmpty($BusinessHours.SaturdayHours2)) {
            $null = $CommandText.AppendLine("        `$OnlineScheduleParams['SaturdayHours'] = $(ConvertFrom-BusinessHoursToTimeRange $BusinessHours.SaturdayHours1 $BusinessHours.SaturdayHours2)")
        }
        if (![string]::IsNullOrEmpty($BusinessHours.SundayHours1) -or ![string]::IsNullOrEmpty($BusinessHours.SundayHours2)) {
            $null = $CommandText.AppendLine("        `$OnlineScheduleParams['SundayHours'] = $(ConvertFrom-BusinessHoursToTimeRange $BusinessHours.SundayHours1 $BusinessHours.SundayHours2)")
        }
        $null = $CommandText.AppendLine('        $AfterHoursSchedule = New-CsOnlineSchedule @OnlineScheduleParams')
        $null = $CommandText.AppendLine('    }')

        $null = $CommandText.AppendLine('    $OptionParams = @{}')
        $ConvertActionParams = @{
            FlowName         = $DefaultQueue.Name
            URI              = $Workflow.NonBusinessHoursActionURI
            QueueId          = $DefaultQueue.NonBusinessHoursActionQueueID
            Action           = $DefaultQueue.NonBusinessHoursActionAction
            CommandText      = $CommandText
            CommandHashName  = 'OptionParams'
            ActionName       = 'Action'
            ActionTargetName = 'CallTarget'
            AAMenuOption     = $true
            Prepend          = ''
        }
        $warn = ConvertNonQuestionAction @ConvertActionParams
        if (![string]::IsNullOrEmpty($warn)) {
            $null = $WarningStrings.AppendLine($warn)
        }

        $null = $CommandText.AppendLine('    $AutomaticMenuOption = New-CsAutoAttendantMenuOption -DtmfResponse Automatic @OptionParams')
        $null = $CommandText.AppendLine('    $AfterHoursMenu = New-CsAutoAttendantMenu -Name ''After Hours Menu'' -MenuOptions @($AutomaticMenuOption)')
        if (![string]::IsNullOrWhiteSpace($Workflow.NonBusinessHoursActionPromptTextToSpeechPrompt)) {
            $null = $CommandText.AppendLine("    `$AfterHoursGreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt '$($Workflow.NonBusinessHoursActionPromptTextToSpeechPrompt)'")
            $null = $CommandText.AppendLine('    $AfterHoursCallFlow = New-CsAutoAttendantCallFlow -Name ''After Hours Call Flow'' -Greetings @($AfterHoursGreetingPrompt) -Menu $AfterHoursMenu')
        }
        elseif (![string]::IsNullOrEmpty($Workflow.NonBusinessHoursActionPromptAudioFilePromptStoredLocation)) {
            AddFileImportScript -ApplicationId OrgAutoAttendant -StoredLocation $Workflow.NonBusinessHoursActionPromptAudioFilePromptStoredLocation -FileName $Workflow.NonBusinessHoursActionPromptAudioFilePromptOriginalFileName -CommandText $CommandText -WarningStrings $WarningStrings -Prepend '    '
            $null = $CommandText.AppendLine('    $AfterHoursGreetingPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $FileId')
            $null = $CommandText.AppendLine('    $AfterHoursCallFlow = New-CsAutoAttendantCallFlow -Name ''After Hours Call Flow'' -Greetings @($AfterHoursGreetingPrompt) -Menu $AfterHoursMenu')
        }
        else {
            $null = $CommandText.AppendLine('    $AfterHoursCallFlow = New-CsAutoAttendantCallFlow -Name ''After Hours Call Flow'' -Menu $AfterHoursMenu')
        }

        $null = $CommandText.AppendLine('    $AfterHoursCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type AfterHours -ScheduleId $AfterHoursSchedule.Id -CallFlowId $AfterHoursCallFlow.Id')
    }

    $c = 0
    foreach ($HolidayId in $Workflow.HolidayHoursIDs) {
        $c++
        if ($null -eq $HolidayId) {
            continue
        }
        $null = $CommandText.AppendLine()
        $null = $CommandText.AppendLine("    # Building the Auto Attendant Holiday Hours Action $c")
        $HolidaySet = @($HolidaySets).Where( { $_.Identity.InstanceID.Guid -eq $HolidayId })[0]
        $HolidayName = [regex]::Replace($HolidaySet.Name, '_?[A-Fa-f0-9]{8}(?:-?[A-Fa-f0-9]{4}){3}-?[A-Fa-f0-9]{12}$', '')
        $null = $CommandText.AppendLine("    `$HolidaySchedule = @(Get-CsOnlineSchedule).Where({ `$_.Name -eq '$HolidayName' })[0]")
        $null = $CommandText.AppendLine('    if ($null -eq $HolidaySchedule) {')
        $null = $CommandText.AppendLine('        $OnlineScheduleParams = @{')
        $null = $CommandText.AppendLine("            Name = '$HolidayName'")
        $null = $CommandText.AppendLine('            FixedSchedule = $true')
        $null = $CommandText.AppendLine('        }')

        $null = $CommandText.AppendLine('        $DateTimeRanges = @()')
        foreach ($HolidayRange in $HolidaySet.HolidayList) {
            $StartDate = $HolidayRange.StartDate.ToString('d/M/yyyy H:mm')
            $EndDate = $HolidayRange.EndDate.ToString('d/M/yyyy H:mm')
            $null = $CommandText.AppendLine("        `$dt = New-CsOnlineDateTimeRange -Start '$StartDate' -End '$EndDate'")
            $null = $CommandText.AppendLine('        $DateTimeRanges += $dt')
        }
        $null = $CommandText.AppendLine('        $OnlineScheduleParams[''DateTimeRanges''] = $DateTimeRanges')
        $null = $CommandText.AppendLine('        $HolidaySchedule = New-CsOnlineSchedule @OnlineScheduleParams')
        $null = $CommandText.AppendLine('    }')

        $null = $CommandText.AppendLine('    $OptionParams = @{}')
        $ConvertActionParams = @{
            FlowName         = $DefaultQueue.Name
            URI              = $Workflow.HolidayActionURI
            QueueId          = $DefaultQueue.HolidayActionQueueID
            Action           = $DefaultQueue.HolidayActionAction
            CommandText      = $CommandText
            CommandHashName  = 'OptionParams'
            ActionName       = 'Action'
            ActionTargetName = 'CallTarget'
            AAMenuOption     = $true
            Prepend          = '    '
        }
        $warn = ConvertNonQuestionAction @ConvertActionParams
        if (![string]::IsNullOrEmpty($warn)) {
            $null = $WarningStrings.AppendLine($warn)
        }

        $null = $CommandText.AppendLine('    $HolidayMenuOption = New-CsAutoAttendantMenuOption -DtmfResponse Automatic @OptionParams')
        $null = $CommandText.AppendLine('    $HolidayMenu = New-CsAutoAttendantMenu -Name ''Holiday Menu'' -MenuOptions @($HolidayMenuOption)')
        if (![string]::IsNullOrEmpty($Workflow.HolidayActionPromptTextToSpeechPrompt)) {
            $null = $CommandText.AppendLine("    `$HolidayGreetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt '$($Workflow.HolidayActionPromptTextToSpeechPrompt)'")
            $null = $CommandText.AppendLine("    `$HolidayCallFlow$c = New-CsAutoAttendantCallFlow -Name 'Holiday Call Flow $c' -Greetings @(`$HolidayGreetingPrompt) -Menu `$HolidayMenu")
        }
        elseif (![string]::IsNullOrEmpty($Workflow.HolidayActionPromptAudioFilePromptStoredLocation)) {
            AddFileImportScript -ApplicationId OrgAutoAttendant -StoredLocation $Workflow.HolidayActionPromptAudioFilePromptStoredLocation -FileName $Workflow.HolidayActionPromptAudioFilePromptOriginalFileName -CommandText $CommandText -WarningStrings $WarningStrings -Prepend '    '
            $null = $CommandText.AppendLine('    $HolidayGreetingPrompt = New-CsAutoAttendantPrompt -AudioFilePrompt $FileId')
            $null = $CommandText.AppendLine("    `$HolidayCallFlow$c = New-CsAutoAttendantCallFlow -Name 'Holiday Call Flow $c' -Greetings @(`$HolidayGreetingPrompt) -Menu `$HolidayMenu")
        }
        else {
            $null = $CommandText.AppendLine("    `$HolidayCallFlow$c = New-CsAutoAttendantCallFlow -Name 'Holiday Call Flow $c' -Menu `$HolidayMenu")
        }

        $null = $CommandText.AppendLine("    `$HolidayCallHandlingAssociation$c = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId `$HolidaySchedule.Id -CallFlowId `$HolidayCallFlow$c.Id")
    }

    $null = $CommandText.AppendLine()
    $null = $CommandText.AppendLine('    # Creating the Auto Attendant')
    $null = $CommandText.AppendLine('    $AAParams = @{')
    $null = $CommandText.AppendLine('        Name            = $AAName')
    $null = $CommandText.AppendLine('        DefaultCallFlow = $DefaultCallFlow')
    $null = $CommandText.AppendLine("        Language        = '$($Workflow.Language)'")
    $null = $CommandText.AppendLine("        TimeZoneId      = '$($Workflow.TimeZone)'")
    $null = $CommandText.AppendLine("        ErrorAction     = 'Stop'")
    $null = $CommandText.AppendLine('    }')

    if ($null -ne $Workflow.BusinessHoursID -or $null -ne $Workflow.HolidayHoursIDs) {
        $null = $CommandText.AppendLine('    $CallFlows = @()')
        $null = $CommandText.AppendLine('    $CallHandlingAssociations = @()')
        if ($null -ne $Workflow.BusinessHoursID) {
            $null = $CommandText.AppendLine('    $CallFlows += $AfterHoursCallFlow')
            $null = $CommandText.AppendLine('    $CallHandlingAssociations += $AfterHoursCallHandlingAssociation')
        }
        for ($i = 1; $i -le $Workflow.HolidayHoursIDs.Count; $i++) {
            if ($null -ne $Workflow.HolidayHoursIDs[($i - 1)]) {
                $null = $CommandText.AppendLine("    `$CallFlows += `$HolidayCallFlow$i")
                $null = $CommandText.AppendLine("    `$CallHandlingAssociations += `$HolidayCallHandlingAssociation$i")
            }
        }
        $null = $CommandText.AppendLine('    $AAParams[''CallFlows''] = $CallFlows')
        $null = $CommandText.AppendLine('    $AAParams[''CallHandlingAssociations''] = $CallHandlingAssociations')
    }
    $null = $CommandText.AppendLine('    try {')
    $null = $CommandText.AppendLine('        $AutoAttendant = New-CsAutoAttendant @AAParams')
    $null = $CommandText.AppendLine('    } catch {')
    $null = $CommandText.AppendLine('        Write-Warning "Unable to create new AA for $AAName ending processing."')
    $null = $CommandText.AppendLine('        Write-Warning $_.Exception')
    $null = $CommandText.AppendLine('        exit')
    $null = $CommandText.AppendLine('    }')
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('    # Creating Auto Attendant Application Instance Association')
    $null = $CommandText.AppendLine('    try {')
    $null = $CommandText.AppendLine('        $AAAssoc = New-CsOnlineApplicationInstanceAssociation -ConfigurationType AutoAttendant -ConfigurationId $($AutoAttendant.Identity) -Identities $AAInstanceId')
    $null = $CommandText.AppendLine('    } catch {')
    $null = $CommandText.AppendLine("        Write-Warning 'Unable to create application association for $($CallQueueParams['Name']) ending processing.'")
    $null = $CommandText.AppendLine('        Write-Warning $_.Exception')
    $null = $CommandText.AppendLine('        exit')
    $null = $CommandText.AppendLine('    }')
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('    # Assigning Usage Location to the created application instances')
    $null = $CommandText.AppendLine('    if ($null -eq $AAInstanceId -or $null -eq $CQInstanceId) {')
    $null = $CommandText.AppendLine('        Write-Warning ''Missing Instance ObjectIDs, cannot assign licenses, ending processing.''')
    $null = $CommandText.AppendLine('        exit')
    $null = $CommandText.AppendLine('    }')
    $null = $CommandText.AppendLine('    try {')
    $null = $CommandText.AppendLine("        Update-MgUser -UserId `$AAInstanceId -UsageLocation '$UsageLocation' -ErrorAction Stop")
    $null = $CommandText.AppendLine("        Update-MgUser -UserId `$CQInstanceId -UsageLocation '$UsageLocation' -ErrorAction Stop")
    $null = $CommandText.AppendLine('    } catch {')
    $null = $CommandText.AppendLine('        Write-Warning ''Unable to set Usage Location for objects, ending processing.''')
    $null = $CommandText.AppendLine('        Write-Warning $_.Exception')
    $null = $CommandText.AppendLine('        exit')
    $null = $CommandText.AppendLine('    }')
    $null = $CommandText.AppendLine('    do {')
    $null = $CommandText.AppendLine('        Start-Sleep -Seconds 2')
    $null = $CommandText.AppendLine('        $AALocation = (Get-CsOnlineUser -Identity $AAInstanceId).UsageLocation')
    $null = $CommandText.AppendLine('        $CQLocation = (Get-CsOnlineUser -Identity $CQInstanceId).UsageLocation')
    $null = $CommandText.AppendLine('    } while ([string]::IsNullOrEmpty($AALocation) -or [string]::IsNullOrEmpty($CQLocation))')
    $null = $CommandText.AppendLine()

    $null = $CommandText.AppendLine('    # Assigning the licenses to the Auto Attendant Instance')
    $null = $CommandText.AppendLine('    $LicenseSkuId = ''440eaaa8-b3e0-484b-a8be-62870b9ba70a'' # this guid is the phone system virtual user by default')
    $null = $CommandText.AppendLine('    $SkuFeaturesToEnable = @(''TEAMS1'',''MCOPSTN1'', ''MCOEV'', ''MCOEV_VIRTUALUSER'')')
    $null = $CommandText.AppendLine('    $StandardLicense = @(Get-MgSubscribedSku).Where({$_.SkuId -eq $LicenseSkuId})[0]')
    $null = $CommandText.AppendLine('    $SkuFeaturesToDisable = $StandardLicense.ServicePlans.Where({$_.ServicePlanName -notin $SkuFeaturesToEnable})')
    $null = $CommandText.AppendLine('    $License = @{}')
    $null = $CommandText.AppendLine('    $License[''SkuId''] = $StandardLicense.SkuId')
    $null = $CommandText.AppendLine('    $License[''DisabledPlans''] = [string[]]$SkuFeaturesToDisable.ServicePlanId')
    $null = $CommandText.AppendLine('    try {')
    $null = $CommandText.AppendLine('        Set-MgUserLicense -UserId $AAInstanceId -AddLicenses @($License) -RemoveLicenses @() -ErrorAction Stop')
    $null = $CommandText.AppendLine('    } catch {')
    $null = $CommandText.AppendLine('        Write-Warning "Unable to apply license plan to $AAInstanceId, you will need to do this manually."')
    $null = $CommandText.AppendLine('        Write-Warning $_.Exception')
    $null = $CommandText.AppendLine('        exit')
    $null = $CommandText.AppendLine('    } finally {')
    $null = $CommandText.AppendLine('        $DeploymentStatus = ''WorkflowDeployed''')
    $null = $CommandText.AppendLine('        $Config.DeploymentStatus = $DeploymentStatus')
    $null = $CommandText.AppendLine('        $Config | Export-Csv -Path $CsvPath -NoTypeInformation')
    $null = $CommandText.AppendLine('    }')
    # pretty sure we have to licence both endpoints now, but unsure...
    $null = $CommandText.AppendLine('    try {')
    $null = $CommandText.AppendLine('        Set-MgUserLicense -UserId $CQInstanceId -AddLicenses @($License) -RemoveLicenses @() -ErrorAction Stop')
    $null = $CommandText.AppendLine('    } catch {')
    $null = $CommandText.AppendLine('        Write-Warning "Unable to apply license plan to $CQInstanceId, you will need to do this manually."')
    $null = $CommandText.AppendLine('        Write-Warning $_.Exception')
    $null = $CommandText.AppendLine('        exit')
    $null = $CommandText.AppendLine('    } finally {')
    $null = $CommandText.AppendLine('        $DeploymentStatus = ''WorkflowDeployed''')
    $null = $CommandText.AppendLine('        $Config.DeploymentStatus = $DeploymentStatus')
    $null = $CommandText.AppendLine('        $Config | Export-Csv -Path $CsvPath -NoTypeInformation')
    $null = $CommandText.AppendLine('    }')
    $null = $CommandText.AppendLine('}')

    if (![string]::IsNullOrEmpty($LineURI)) {
        $LineURI = [regex]::Replace($LineURI, 'x', ';ext=')
        $null = $CommandText.AppendLine()
        $null = $CommandText.AppendLine('if ($AssignLineUri -and $DeploymentStatus.ToLower() -eq ''workflowdeployed'') {')
        $null = $CommandText.AppendLine('    # Assigning the phone number Auto Attendant Instance')
        $null = $CommandText.AppendLine('    try {')
        $null = $CommandText.AppendLine("        `$null = Set-CsPhoneNumberAssignment -Identity `$AAInstanceId -PhoneNumber '+$LineUri' -PhoneNumberType DirectRouting -ErrorAction Stop")
        $null = $CommandText.AppendLine('    } catch {')
        $null = $CommandText.AppendLine("        Write-Warning `"Unable to assign LineUri $LineUri to `$AAInstanceId, ending processing.`"")
        $null = $CommandText.AppendLine('        Write-Warning $_.Exception')
        $null = $CommandText.AppendLine('        exit')
        $null = $CommandText.AppendLine('    }')

        $null = $CommandText.AppendLine()
        $null = $CommandText.AppendLine('    $DeploymentStatus = ''LineUriAssigned''')
        $null = $CommandText.AppendLine('    $Config.DeploymentStatus = $DeploymentStatus')
        $null = $CommandText.AppendLine('    $Config | Export-Csv -Path $CsvPath -NoTypeInformation')
        $null = $CommandText.AppendLine('}')
    }

    $CsvPath = [IO.Path]::Combine($GeneratedScriptsPath, ($AADispName + '.csv'))
    $WorkflowInfo = [PSCustomObject]@{
        AADispName       = $AADispName
        CQDispName       = $CQDispName
        CQName           = $CQName
        AAName           = $AAName
        DeploymentStatus = 'New'
    }
    $WorkflowInfo | Export-Csv -Path $CsvPath -NoTypeInformation

    # Remove traling newline
    $null = $CommandText.Remove($CommandText.Length - [Environment]::NewLine.Length, [Environment]::NewLine.Length)

    $Warnings = (($WarningStrings.ToString() -split [Environment]::NewLine) | ForEach-Object { if (![string]::IsNullOrWhiteSpace($_)) { "# $_" } }) -join [Environment]::NewLine
    if ($Warnings.Length -gt 0) {
        $Warnings = '# WARNINGS' + [Environment]::NewLine + $Warnings
        $Warnings += [Environment]::NewLine + [Environment]::NewLine
    }
    $Requires = @'
#requires -Modules MicrosoftTeams
#requires -Modules Microsoft.Graph.Users
#requires -Modules Microsoft.Graph.Identity.DirectoryManagement
#requires -Modules Microsoft.Graph.Users.Actions

'@

    $Content = ($Requires + $Warnings + $CommandText.ToString()) -replace '    ', '    '
    Set-Content -Path $FileName -Value $Content -Encoding UTF8
}

$StatusScript = @'
$CsvFiles = Get-ChildItem -Path $PSScriptRoot -Filter *.csv
$Statuses = @{}
foreach ($CsvFile in $CsvFiles) {
    $status = (Import-Csv -Path $CsvFile)[0]
    if ($null -eq $Statuses[$status.DeploymentStatus]) {
        $Statuses[$status.DeploymentStatus] = [Collections.Generic.List[string]]::new()
    }
    $Statuses[$status.DeploymentStatus].Add([IO.Path]::GetFileNameWithoutExtension($CsvFile.Name))
}
[PSCustomObject]$Statuses
'@

Set-Content -Value $StatusScript -Path ([IO.Path]::Combine($GeneratedScriptsPath, 'GetDeploymentStatus.ps1'))