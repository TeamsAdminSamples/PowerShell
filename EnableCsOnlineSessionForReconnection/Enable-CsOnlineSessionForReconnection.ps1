#Requires -Modules MicrosoftTeams

function Enable-CsOnlineSessionForReconnection {
    $modules = Get-Module tmp_*
    $csModuleUrl = "/OcsPowershellOAuth"
    $isSfbPsModuleFound = $false;

    foreach ($module in $modules) {
        [string] $moduleUrl = $module.Description
        [int] $queryStringIndex = $moduleUrl.IndexOf("?")

        if ($queryStringIndex -gt 0) {
            $moduleUrl = $moduleUrl.SubString(0, $queryStringIndex)
        }

        if ($moduleUrl.IndexOf($csModuleUrl) -ge 0) {
            $isSfbPsModuleFound = $true
            $connUrl = $module.Description -replace '^.+(https)', 'https'

            $getScriptBlock = {
                param(
                    [Parameter(Mandatory = $true, Position = 0)]
                    [string] 
                    $commandName
                )
                
                function Get-TeamsTokens {
                    [cmdletbinding()]
                    param(
                        [Parameter(Mandatory = $false)]
                        [string]
                        $UserName
                    )

                    $LegacyModule = (Get-Module -Name MicrosoftTeams).Version -le [Version]::new('1.1.6')

                    $provider = if ($LegacyModule) {
                        [Microsoft.Open.Teams.CommonLibrary.TeamsPowerShellSession]::SessionProvider
                    }
                    else {
                        [Microsoft.TeamsCmdlets.Powershell.Connect.TeamsPowerShellSession]::SessionProvider
                    }

                    if ($null -eq $provider) {
                        $ConnectParams = @{}
                        if (![string]::IsNullOrEmpty($UserName)) {
                            # need to figure out how to handle auth for a specific account in new module, the behavior has changed
                            if ($LegacyModule) {
                                $ConnectParams['AccountId'] = $UserName
                            }
                        }
                        elseif (![string]::IsNullOrEmpty($script:AccountId)) {
                            if ($LegacyModule) {
                                $ConnectParams['AccountId'] = $script:AccountId
                            }
                        }
                        Connect-MicrosoftTeams @ConnectParams | Out-Null

                        $provider = if ($LegacyModule) {
                            [Microsoft.Open.Teams.CommonLibrary.TeamsPowerShellSession]::SessionProvider
                        }
                        else {
                            [Microsoft.TeamsCmdlets.Powershell.Connect.TeamsPowerShellSession]::SessionProvider
                        }
                    }

                    if ($LegacyModule) {
                        $MsAccessToken = $provider.GetAccessToken([Microsoft.Open.Teams.CommonLibrary.Endpoint]::MsGraphEndpointResourceId)
                        $AadAccessToken = $provider.GetAccessToken([Microsoft.Open.Teams.CommonLibrary.Endpoint]::AadGraphEndpointResourceId)
                        $ConfigAccessToken = $provider.GetAccessToken([Microsoft.Open.Teams.CommonLibrary.Endpoint]::ConfigApiEndpointResourceId)

                        $MsAccessToken.AuthorizeRequest([Microsoft.Open.Teams.CommonLibrary.Endpoint]::MsGraphEndpointResourceId) | Out-Null
                        $AadAccessToken.AuthorizeRequest([Microsoft.Open.Teams.CommonLibrary.Endpoint]::AadGraphEndpointResourceId) | Out-Null
                        $ConfigAccessToken.AuthorizeRequest([Microsoft.Open.Teams.CommonLibrary.Endpoint]::ConfigApiEndpointResourceId) | Out-Null
                        $script:AccountId = $AadAccessToken.UserId
                        @{
                            AccountId         = $script:AccountId
                            MsAccessToken     = $MsAccessToken.AccessToken
                            AadAccessToken    = $AadAccessToken.AccessToken
                            ConfigAccessToken = $ConfigAccessToken.AccessToken
                        }
                    }
                    else {
                        # config api 
                        $Resource = $provider.GetEndpoint([Microsoft.TeamsCmdlets.Powershell.Connect.Common.Endpoint]::ConfigApiEndpointResourceId)
                        $Scope = [Microsoft.TeamsCmdlets.Powershell.Connect.Models.ModuleConfiguration]::DefaultScopesForConfigAPI
                        $ConfigAccessToken = $provider.GetAccessToken($Resource, $Scope)
                        $ConfigAccessToken.AuthorizeRequest($Resource, $Scope) | Out-Null

                        $Resource = $provider.GetEndpoint([Microsoft.TeamsCmdlets.Powershell.Connect.Common.Endpoint]::AadGraphEndpointResourceId)
                        $Scope = [Microsoft.TeamsCmdlets.Powershell.Connect.Models.ModuleConfiguration]::DefaultScopesForGraphAPI
                        $AccessToken = $provider.GetAccessToken($Resource, $Scope)
                        $AccessToken.AuthorizeRequest($Resource, $Scope) | Out-Null
                        $script:AccountId = $AccessToken.UserId
                        @{
                            AccountId         = $script:AccountId
                            AccessToken       = $AccessToken.AccessToken
                            ConfigAccessToken = $ConfigAccessToken.AccessToken
                        }
                    }
                }

                function ParseJWTString {
                    param (
                        [string]
                        $JWT
                    )
                    $jwtString = $JWT.Split('.')[1]
                    [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String((PadBase64String $jwtString))) | ConvertFrom-Json
                }
            
                function PadBase64String([string] $stringToPad) {
                    $PaddingLength = 4 - $stringToPad.Length % 4
                    if ($PaddingLength -eq 4) {
                        $PaddingLength = 0
                    }
                    $stringToPad + [String]::new("=", $PaddingLength)
                }
            
                function IsExpired {
                    [CmdletBinding(DefaultParameterSetName = 'Token')]
                    param (
                        # Token as SecureString
                        [Parameter(Mandatory = $true, 
                            Position = 0, 
                            ParameterSetName = "Secure")]
                        [AllowNull()]
                        [SecureString]
                        $SecureToken,
            
                        # Token as plaintext
                        [Parameter(Mandatory = $true,
                            Position = 0,
                            ParameterSetName = "Raw")]
                        [AllowNull()]
                        [AllowEmptyString()]
                        [string]
                        $Token,
            
                        # Token as hashtable
                        [Parameter(Mandatory = $true,
                            Position = 0,
                            ParameterSetName = "Hash")]
                        [hashtable]
                        [AllowNull()]
                        $TokenHash,
            
                        # Minutes in the future to test for expiration
                        [Parameter(Position = 1)]
                        [int]$OffsetMinutes = 1 # 45
                    )
                    if ($PSCmdlet.ParameterSetName -eq "Hash") {
                        # return true if any are expired
                        if ($null -eq $TokenHash) { $true }
                        elseif ((IsExpired $TokenHash['AadAccessToken'] -OffsetMinutes $OffsetMinutes)) {
                            $true
                        }
                        elseif ((IsExpired $TokenHash['MsAccessToken'] -OffsetMinutes $OffsetMinutes)) {
                            $true
                        }
                        elseif ((IsExpired $TokenHash['ConfigAccessToken'] -OffsetMinutes $OffsetMinutes)) {
                            $true
                        }
                        else {
                            $false
                        }
                    }
                    else {
                        $IsExpired = if ($null -eq $Token -and $null -eq $SecureToken) { $true }
                        else {
                            if ($null -eq $SecureToken) {
                                $result = $Token
                            }
                            else {
                                $result = [System.Net.NetworkCredential]::new('', $SecureToken).Password
                            }
                            $jwt = ParseJWTString -JWT $result
                            $ExpiresTime = [DateTime]::new(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc).AddSeconds($jwt.exp)
                            [DateTime]::UtcNow.AddMinutes($OffsetMinutes) -gt $ExpiresTime
                        }
                        $IsExpired
                    }
                }

                $tmpSession = $script:PSSession
                $sessionExpired = IsExpired -SecureToken $script:PSSession.Runspace.ConnectionInfo.Credential.Password
            
                if ($sessionExpired -or `
                        $null -eq $script:PSSession -or `
                        $script:PSSession.Runspace.RunspaceStateInfo.State -ne 'Opened') {
                    Write-PSImplicitRemotingMessage "Recreating a new remote powershell session (implicit) for command: `"$commandName`" $(if($sessionExpired) { "due to pending token expiration" } )..."
                    $t = Get-TeamsTokens
                    $configToken = [System.Net.NetworkCredential]::new('', $t['ConfigAccessToken']).SecurePassword 

                    try {
                        $session = ( 
                            $( 
                                & $script:NewPSSession `
                                    -connectionUri "{0}" -ConfigurationName 'Microsoft.PowerShell' `
                                    -SessionOption (Get-PSImplicitRemotingSessionOption) `
                                    -Credential ( [PSCredential]::new('oauth', $configToken) ) `
                                    -Authentication Basic `
                                    -ErrorAction Stop
                            )
                        )
                    }
                    catch {
                        $message = $_.Exception.Message
                        $id = $_.FullyQualifiedErrorId
                        Write-PSImplicitRemotingMessage "Unable to create new session: $id => $message"
                        Write-PSImplicitRemotingMessage "Attempting to reuse existing session for this command..."
                        $session = $tmpSession
                    }

                    Set-PSImplicitRemotingSession `
                        -CreatedByModule ($tmpSession.InstanceId.Guid -ne $session.InstanceId.Guid) `
                        -PSSession $session
                }

                if ($null -ne $script:PSSession -and $tmpSession.InstanceId.Guid -ne $script:PSSession.InstanceId.Guid) {
                    # Remove old session
                    if (!$tmpSession.Name.StartsWith("SfBPowerShellSessionViaTeamsModule_")) {
                        $tmpSession | Remove-PSSession
                    }
                    $m = Get-Module -Name tmp_*
                    if ($null -eq $m) {
                        Write-PSImplicitRemotingMessage "Removing old PSSession removed module"
                    }
                }

                if ($null -eq $script:PSSession -and $null -ne $t) {
                    Set-PSImplicitRemotingSession `
                        -CreatedByModule $true `
                        -PSSession (New-CsOnlineSession)
                }
                
                # Cleanup Sessions
                Get-PSSession | Where-Object { 
                    $_.InstanceId.Guid -ne $script:PSSession.InstanceId.Guid `
                        -and ($_.Name.IndexOf([IO.Path]::GetFileNameWithoutExtension((Get-PSImplicitRemotingModuleName))) -ge 0) `
                        -and $_.State -in @("Broken", "Closed")
                } | Remove-PSSession

                if (($null -eq $script:PSSession) -or ($script:PSSession.Runspace.RunspaceStateInfo.State -ne 'Opened')) {
                    throw 'No session has been associated with this implicit remoting module.'
                }
                return [Management.Automation.Runspaces.PSSession]$script:PSSession
            }
            $sbString = $getScriptBlock.ToString()
            $sbString = $sbString -replace '([\{\}])', '$1$1' -replace '\{\{(\d)\}\}', '{$1}'
            $getScriptBlock = [ScriptBlock]::Create(($sbString -f $connUrl))

            & $module { param($SB) ${function:Get-PSImplicitRemotingSession} = $SB } $getScriptBlock
        }
    }

    if ($isSfbPsModuleFound -eq $false) {
        Write-Error "Please run this cmdlet after importing the session created by New-CsOnlineSession"
    }
}