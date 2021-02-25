#Requires -Modules @{ ModuleName = 'MicrosoftTeams'; GUID = 'd910df43-3ca6-4c9c-a2e3-e9f45a8e2ad9'; ModuleVersion = '1.1.6' }

param (
    $PChatDataPath = ".\PersistentChatData.zip"
)

$BasePath = Split-Path -Path $PChatDataPath -Parent

if (!(Test-Path -Path "${BasePath}\PersistentChatData")) {
    Expand-Archive -Path $PChatDataPath -DestinationPath "${BasePath}\PersistentChatData"
}
$AdGroups = ([xml](Get-Content -Path "${BasePath}\PersistentChatData\AdGroups.xml" -Raw)).Pool.AdGroups.Group
$AdAffiliations = ([xml](Get-Content -Path "${BasePath}\PersistentChatData\AdAffiliations.xml" -Raw)).Pool.AdAffiliations.User
$AdUsers = ([xml](Get-Content -Path "${BasePath}\PersistentChatData\AdUsers.xml" -Raw)).Pool.AdUsers.User | Select-Object -Property principalId, UserPrincipalName

$GroupIds = $AdGroups.principalId
$GroupAffiliations = $AdAffiliations.Where( { $_.id -in $GroupIds })

$UserIds = $AdUsers.principalId
$UserAffiliations = $AdAffiliations.Where( { $_.id -in $UserIds })

function EnumerateGroup ($Group) {
    $Users = [Collections.Generic.List[object]]::new()
    $NonNestedUsers = $UserAffiliations.Where( { $_.Groups.Group.id -contains $Group.principalId })
    if ($null -ne $NonNestedUsers) {
        foreach ($NonNestedUser in $NonNestedUsers) {
            $User = $AdUsers.Where( { $_.principalId -eq $NonNestedUser.id })
            $Users.Add($User) | Out-Null
        }
    }
    $NestedGroups = $GroupAffiliations.Where( { $_.Groups.Group.id -contains $Group.principalId })
    if ($null -ne $NestedGroups) {
        foreach ($NestedGroup in $NestedGroups) {
            $NestedUsers = EnumerateGroup -Group $NestedGroup
            if ($null -ne $NestedUsers) {
                foreach ($NestedUser in $NestedUsers) {
                    $User = $AdUsers.Where( { $_.principalId -eq $NestedUser.id })
                    $Users.Add($User) | Out-Null
                }
            }
        }
    }
    $Users | Sort-Object -Property principalId -Unique
}
function EnumerateGroups ([Collections.Generic.List[object]]$GroupsToEnumerate, [Collections.Generic.List[object]]$EnumeratedGroups) {
    foreach ($Group in $GroupsToEnumerate) {
        $Members = EnumerateGroup $Group
        $UpdatedGroup = $Group | Add-Member -MemberType NoteProperty -Name Members -Value $Members -Force -PassThru
        $EnumeratedIds = $EnumeratedGroups | Select-Object -ExpandProperty principalId
        if ($UpdatedGroup.principalId -notin $EnumeratedIds) {
            $EnumeratedGroups.Add($UpdatedGroup) | Out-Null
        }
    }
    $GroupsToEnumerate.Clear() | Out-Null
}
function ProcessChatRooms ($FilePath, [Collections.Generic.List[object]]$GroupsToEnumerate, [Collections.Generic.List[object]]$EnumeratedGroups) {
    do {
        $Reprocess = $false
        $reader = [Xml.XmlReader]::Create($FilePath)
        $reader.ReadToFollowing("Pool") | Out-Null
        $reader.ReadToDescendant("ChatRooms") | Out-Null
        $ChatRooms = if ($reader.ReadToDescendant("ChatRoom")) {
            do {
                $subTreeReader = $reader.ReadSubtree()
                $subTreeReader.Skip()
    
                $isDisabled = $subTreeReader.GetAttribute("isDisabled")
                $Name = if ($subTreeReader.ReadToDescendant("Name")) {
                    $subTreeReader.ReadElementContentAsString()
                }
                $Description = if ($subTreeReader.Name -eq "Desc") {
                    $subTreeReader.ReadElementContentAsString()
                }
    
                $Roles = if ($subTreeReader.Name -eq "Roles") {
                    $RolesTemp = [PSCustomObject]@{
                        Managers = [Collections.Generic.List[object]]::new()
                        Members  = [Collections.Generic.List[object]]::new()
                        Creators = [Collections.Generic.List[object]]::new()
                    }
                    if ($subTreeReader.ReadToDescendant("Principal")) {
                        do {
                            $id = $subTreeReader.GetAttribute("id")
                            $isManager = $subTreeReader.GetAttribute("isManager") -eq '1'
                            $isMember = $subTreeReader.GetAttribute("isMember") -eq '1'
                            $isCreator = $subTreeReader.GetAttribute("isCreator") -eq '1'
    
                            $AdObject = $AdGroups | Where-Object { $id -eq $_.principalId } | Select-Object -First 1
                            if ($null -eq $AdObject) {
                                $AdObject = $AdUsers | Where-Object { $id -eq $_.principalId } | Select-Object -First 1 -Property UserPrincipalName
                            }
                            elseif ($AdObject.principalId -in $EnumeratedGroups.principalId) {
                                $AdObject = $EnumeratedGroups.Where( { $_.principalId -eq $AdObject.principalId }).Members | Select-Object -Property UserPrincipalName
                            }
                            else {
                                if ($GroupsToEnumerate.principalId -notcontains $AdObject.principalId) {
                                    $GroupsToEnumerate.Add($AdObject) | Out-Null
                                }
                                $Reprocess = $true
                            }
    
                            if ($isManager -or $isCreator) {
                                foreach ($obj in $AdObject) {
                                    Write-Host "Manager: $($obj.UserPrincipalName)"
                                    $RolesTemp.Managers.Add($obj) | Out-Null
                                }
                            }
                            elseif ($isMember) {
                                foreach ($obj in $AdObject) {
                                    Write-Host "Member: $($obj.UserPrincipalName)"
                                    $RolesTemp.Members.Add($obj) | Out-Null
                                }
                            }
                        } while ($subTreeReader.ReadToNextSibling("Principal"))
                    }
                    $RolesTemp
                }
    
                [PSCustomObject]@{
                    IsDisabled  = $isDisabled
                    Name        = $Name
                    Description = $Description
                    Roles       = $Roles
                }
                $subTreeReader.Close()
                $subTreeReader.Dispose()
            } while ($reader.ReadToNextSibling("ChatRoom"))
        }
        $reader.Close()
        $reader.Dispose()
        if ($Reprocess) {
            Write-Host "Enumerating group membership for $($GroupsToEnumerate.Count) groups"
            Write-Host "Depending on your configuration, this may take several minutes"
            EnumerateGroups $GroupsToEnumerate $EnumeratedGroups
            Write-Host "Reprocessing ChatRooms data with group memberships..."
        }
    } while ($Reprocess)
    $ChatRooms
}

$GroupsToEnumerate = [Collections.Generic.List[object]]::new()
$EnumeratedGroups = [Collections.Generic.List[object]]::new()
$ChatRooms = ProcessChatRooms "${BasePath}\PersistentChatData\ChatRooms.xml" $GroupsToEnumerate $EnumeratedGroups

foreach ($ChatRoom in $ChatRooms) {
    $NewTeamParams = @{
        Visibility                        = 'Private'
        AllowGiphy                        = $true
        GiphyContentRating                = 'Moderate'
        AllowStickersAndMemes             = $true
        AllowCustomMemes                  = $true
        AllowGuestCreateUpdateChannels    = $false
        AllowGuestDeleteChannels          = $false
        AllowCreateUpdateChannels         = $true
        AllowDeleteChannels               = $true
        AllowAddRemoveApps                = $true
        AllowCreateUpdateRemoveTabs       = $true
        AllowCreateUpdateRemoveConnectors = $true
        AllowUserEditMessages             = $true
        AllowUserDeleteMessages           = $true
        AllowOwnerDeleteMessages          = $true
        AllowTeamMentions                 = $true
        AllowChannelMentions              = $true
        ShowInTeamsSearchAndSuggestions   = $true
        RetainCreatedGroup                = $false
    }

    $NewTeamParams['DisplayName'] = $ChatRoom.Name
    if (![string]::IsNullOrWhiteSpace($ChatRoom.Description)) {
        $NewTeamParams['Description'] = $ChatRoom.Description
    }
    # check if null first
    if ($null -ne $ChatRoom.Roles -and $null -ne $ChatRoom.Roles.Managers -and $null -ne $ChatRoom.Roles.Managers[0]) {
        $NewTeamParams['Owner'] = $ChatRoom.Roles.Managers[0].UserPrincipalName
    }
    else {
        Write-Warning "No owner for $($ChatRoom.Name), current admin account will be set as owner"
    }
    try {
        $NewTeamId = Get-Team -DisplayName $NewTeamParams['DisplayName'] -ErrorAction SilentlyContinue
        if ($null -eq $NewTeamId) {
            $NewTeamId = New-Team @NewTeamParams -ErrorAction Stop
        } else {
            Write-Host "$($ChatRoom.Name) already exists, skipping creation."
        }
    }
    catch {
        Write-Warning "Failed to create $($ChatRoom.Name)`r`n$($_.Exception.Message)"
        continue
    }
    if ($null -ne $ChatRoom.Roles) {
        foreach ($owner in $ChatRoom.Roles.Managers) {
            if ($owner.UserPrincipalName -eq $NewTeamParams.Owner) { continue }
            try {
                Add-TeamUser -GroupId $NewTeamId.GroupId -User $owner.UserPrincipalName -Role Owner -ErrorAction Stop
            }
            catch {
                Write-Warning "Failed to add owner $($owner.UserPrincipalName) to $($ChatRoom.Name)`r`n$($_.Exception.Message)"
                continue
            }
        }
        foreach ($member in $ChatRoom.Roles.Members) {
            try {
                Add-TeamUser -GroupId $NewTeamId.GroupId -User $member.UserPrincipalName -Role Member -ErrorAction Stop
            }
            catch {
                Write-Warning "Failed to add member $($member.UserPrincipalName) to $($ChatRoom.Name)`r`n$($_.Exception.Message)"
                continue
            }
        }
    }
}