#Requires -Modules @{ ModuleName = 'MicrosoftTeams'; GUID = 'd910df43-3ca6-4c9c-a2e3-e9f45a8e2ad9'; ModuleVersion = '1.1.6' }

<#
    .SYNOPSIS
    Grants the given TeamsMobilityPolicy
    
    .DESCRIPTION
    Grants the given TeamsMobilityPolicy for all users enabled or from a user provided list.
    It will run with multiple sessions from multiple administrator accounts as desired to increase parallelization.

    .INPUTS
    None. This script takes no pipeline input

    .OUTPUTS
    None. This script fully executes the process and outputs nothing to the pipeline. All console interaction is informational only.

    .EXAMPLE
    Grant-CsTeamsMobilityPolicyBatch -UserName admin1@contoso.com, admin2@contoso.com -PolicyName $null -IsTest $false
    This runs a policy update for all enabled users in the tenant leveraging 2 administrator accounts for a maximum of 6 concurrent sessions

    .EXAMPLE
    Grant-CsTeamsMobilityPolicyBatch -UserName admin1@contoso.com -PolicyName $null -UsersFilePath .\userstomigrate.txt
    This runs a test policy grant for the users listed in .\userstomigrate.txt using 1 administrator account
#>

param(
    #  The name of the policy instance.
    [string]$PolicyName,

    # Specifies if Grant-CsTeamsMobilityPolicyBatch should be run with WhatIf switch, default is $true, set to $false to actually perform the change
    [bool]$IsTest = $true,

    # Specifies the admin user names used to open SkypeOnline PowerShell sessions
    [Parameter(Mandatory = $true)]
    [string[]]$UserName,

    # Specifies the path to the text file containing the list of User Principal Names of the users to update.
    # If not specified, the script will run for all enabled users
    [string]$UsersFilePath,

    # Specifies the directory path where logs will be created
    [string]$LogFolderPath = ".",
    
    # Specifies whether the Completed/Remaining user files will be created with UPN
    [switch]$LogEUII,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [string]$TeamsEnvironmentName
)

function Get-CsOnlineSessionFromTokens {
    [cmdletbinding()]
    param(
        [hashtable]
        $TeamsTokens
    )
    SetupShellDependencies
    Connect-MicrosoftTeams @TeamsTokens | Out-Null
    New-CsOnlineSession
    Disconnect-MicrosoftTeams
}

function Get-TeamsTokens {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]
        $UserName,
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$TeamsEnvironmentName
    )

    # Add means to handle TeamsEnvironmentName (TeamsCloud, TeamsGCCH, TeamsDOD)
    SetupShellDependencies
    $TeamsEnvironmentParam = @{}
    if (![string]::IsNullOrEmpty($TeamsEnvironmentName)) {
        $TeamsEnvironmentParam['TeamsEnvironmentName'] = $TeamsEnvironmentName
    }
    Connect-MicrosoftTeams -AccountId $UserName @TeamsEnvironmentParam | Out-Null
    $provider = [Microsoft.Open.Teams.CommonLibrary.TeamsPowerShellSession]::SessionProvider
    $account = [Microsoft.Open.Teams.CommonLibrary.AzureAccount]::new()
    $account.Type = [Microsoft.Open.Teams.CommonLibrary.AzureAccount+AccountType]::User
    $account.Id = $UserName
    $PromptBehavior = [Microsoft.Open.Teams.CommonLibrary.ShowDialog]::Auto
    $tenantId = $provider.AccessTokens['AccessToken'].TenantId
    $endpoint = [Microsoft.Open.Teams.CommonLibrary.Endpoint]::AadGraphEndpointResourceId
    $aadToken = $provider.AuthenticationFactory.Authenticate($account, $provider.AzureEnvironment, $tenantId, $null, $PromptBehavior, $null, $provider.TokenCache, $endpoint)
    $endpoint = [Microsoft.Open.Teams.CommonLibrary.Endpoint]::MsGraphEndpointResourceId
    $msToken = $provider.AuthenticationFactory.Authenticate($account, $provider.AzureEnvironment, $tenantId, $null, $PromptBehavior, $null, $provider.TokenCache, $endpoint)
    $endpoint = [Microsoft.Open.Teams.CommonLibrary.Endpoint]::ConfigApiEndpointResourceId
    $configToken = $provider.AuthenticationFactory.Authenticate($account, $provider.AzureEnvironment, $tenantId, $null, $PromptBehavior, $null, $provider.TokenCache, $endpoint)
    Disconnect-MicrosoftTeams

    @{
        AccountId = $UserName
        AadAccessToken = $aadToken.AccessToken
        MsAccessToken = $msToken.AccessToken
        ConfigAccessToken = $configToken.AccessToken
    }
}

function GetUsedLocalFunctions {
    param (
        [ScriptBlock]
        $Script,
        [Collections.Generic.List[object]]
        $Functions = $null,
        [bool]
        $GetStrings = $true
    )
    if ($null -eq $Functions) {
        $allFunctions = Get-ChildItem -Path Function:
        $Functions = [Collections.Generic.List[object]]::new()
        foreach ($func in $allFunctions) {
            $Functions.Add($func) | Out-Null
        }
    }
    $newFunctions = [Collections.Generic.List[object]]::new()
    foreach ($func in $Functions) {
        if ($func.ScriptBlock -ne $Script) {
            $newFunctions.Add($func) | Out-Null
        }
    }
    $usedFunctions = foreach ($func in $newFunctions) { 
        if ($Script.ToString().ToLower().IndexOf($func.Name.ToLower()) -ge 0) {
            $func
            GetUsedLocalFunctions -Script $func.ScriptBlock -Functions $newFunctions -GetStrings $false
        }
    }
    $usedFunctions = $usedFunctions | Sort-Object -Property Name -Unique
    if ($GetStrings) {
        $usedFunctions | ForEach-Object { "function $($_.Name) {$([Environment]::NewLine)$($_.Definition.Trim([Environment]::NewLine))$([Environment]::NewLine)}" }
    } else {
        $usedFunctions
    }
}

function Invoke-CsOnlineBatch {
    <#
    .SYNOPSIS
    Starts the Meeting Migration service for SfB meetings to Teams meetings
    
    .DESCRIPTION
    Starts the Meeting Migration service for all users enabled in a Tenant or from a user provided list.
    It will run with multiple sessions from multiple administrator accounts as desired to increase parallelization.

    .INPUTS
    None. This script takes no pipeline input

    .OUTPUTS
    Whatever you expect output wise from the passed in JobScript
#>
    param(
        # Specifies the admin user names used to open SkypeOnline PowerShell sessions
        [Parameter(Mandatory = $true)]
        [string[]]$UserName,
    
        # Specifies the path to the text file containing the list of User Principal Names of the users to migrate.
        # If not specified, the script will run for all enabled users assigned the default TeamsUpgradePolicy
        [string]$UsersFilePath,

        [ScriptBlock]$FilterScript,

        [Parameter(Mandatory = $true)]
        [ScriptBlock]$JobScript,
    
        [object[]]$OtherArgs,

        # Specifies the directory path where logs will be created
        [string]$LogFolderPath = ".",
        
        # Specifies whether the Completed/Remaining user files will be created with UPN
        [switch]$LogEUII,

        [string]$LogFile,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$TeamsEnvironmentName
    )

    $SessionsToOpen = 3
    $ExpirationOffsetMinutes = 5

    $TeamsEnvironmentParam = @{}
    if (![string]::IsNullOrEmpty($TeamsEnvironmentName)) {
        $TeamsEnvironmentParam['TeamsEnvironmentName'] = $TeamsEnvironmentName
    }

    $LogFolderPath = Resolve-Path $LogFolderPath
    $LogDate = Get-Date -Format "yyyyMMdd_HHmmss"
    if ([string]::IsNullOrWhiteSpace($LogFile)) {
        $LogFile = "$LogFolderPath\CsOnlineBatch_$LogDate.log"
    }
    else {
        if (!$LogFile.StartsWith($LogFolderPath)) {
            $LogFile = "$LogFolderPath\$LogFile"
        }
    }
    $UsersCompleted = "$LogFolderPath\CsOnlineBatch_UsersCompleted_$LogDate.txt"
    $UsersRemaining = "$LogFolderPath\CsOnlineBatch_UsersRemaining_$LogDate.txt"

    Write-Log -Level Info -Path $LogFile -Message "##################################################"
    Write-Log -Level Info -Path $LogFile -Message "Accounts used: $($UserName.Length)"
    Write-Log -Level Info -Path $LogFile -Message "OverrideAdminDomain: $OverrideAdminDomain"
    Write-Log -Level Info -Path $LogFile -Message "UsersFilePath: $(if([string]::IsNullOrWhiteSpace($UsersFilePath)){$null}else{Resolve-Path $UsersFilePath})"
    Write-Log -Level Info -Path $LogFile -Message "FilterScript: {$($FilterScript.ToString())}"
    Write-Log -Level Info -Path $LogFile -Message "JobScript: {$($JobScript.ToString())}"

    function WriteUsers ([object[]]$Users, $Path) {
        if ($LogEUII) {
            if (!(Test-Path $Path)) {
                New-Item $Path -Force -ItemType File | Out-Null
            }
            $WrittenUsers = @(Get-Content -Path $Path).Where({ $_ -match '^[^@]+@.+$' }) | Sort-Object -Unique
            $UserString = $Users.Where({$_ -notin $WrittenUsers}) -Join [Environment]::NewLine
            $UserString | Out-File -FilePath $Path -Append
        }
    }

    function WriteCompleted ([object[]]$Users) {
        WriteUsers $Users $UsersCompleted
    }

    function WriteRemaining ([object[]]$Users) {
        WriteUsers $Users $UsersRemaining
    }

    try {
        SetupShellDependencies
    } catch {
        Write-Warning -Message $_.Exception.Message
        Write-Log -Level Error -Path $LogFile -Message $_.Exception.Message
        exit 1
    }
    # Attempt to initialize all available sessions to get valid tokens
    
    if ($null -eq $Tokens) {
        $Script:Tokens = [Collections.Generic.List[object]]::new()
    }
    foreach ($adminUser in $UserName) {
        $currentToken = $Tokens | Where-Object { $_.AccountId -eq $adminUser }
        $Tokens.Remove($currentToken) | Out-Null
        if ($null -eq $currentToken) {
            Write-Host "If prompted, please sign in using the pop up window for $adminUser."
            Write-Log -Level Info -Message "Attempting to get access token for admin account." -Path $LogFile
            $currentToken = Get-TeamsTokens -UserName $adminUser @TeamsEnvironmentParam
        }
        if (IsExpired $currentToken $ExpirationOffsetMinutes) {
            Write-Host "The access token has expired, if prompted, please sign in again using the pop up window for $adminUser."
            Write-Log -Level Warn -Message "The access token has expired, attempting to get new access token for admin account." -Path $LogFile
            $currentToken = Get-TeamsTokens -UserName $adminUser @TeamsEnvironmentParam
        }
        $Script:Tokens.Add($currentToken) | Out-Null
    }

    # Cleanup old sessions
    Get-PSSession | Where-Object { $_.Name.StartsWith("SfBPowerShellSession") } | Remove-PSSession
    $Sessions = @()
    foreach ( $Token in $Tokens ) {
        $adminUser = $token.AccountId

        if (IsExpired $Token) {
            continue
        }
        $ToOpen = $SessionsToOpen
        for ($i = 0; $i -lt $ToOpen; $i++) {
            Write-Verbose "Attempting to create session $($i+1) for $adminUser"
            Write-Log -Level Info -Message "Attempting to create session $($i+1)" -Path $LogFile
            $NewSession = $null
            try {
                $NewSession = Get-CsOnlineSessionFromTokens -TeamsTokens $Token -ErrorAction Stop
            }
            catch { 
                Write-Warning "Unable to create session: $($_.Exception.Message)" 
                Write-Log -Level Warn -Message "Unable to create session: $($_.Exception.Message)" -Path $LogFile
                
            }
            if ($null -ne $NewSession) {
                $Sessions += [Collections.Generic.KeyValuePair[string, Management.Automation.Runspaces.PSSession]]::new($adminUser, $NewSession)
            }
            else {
                $i--
                $ToOpen--
            }
        }
    }
    if ($Sessions.Count -eq 0) {
        Write-Warning "Unable to create a valid session"
        Write-Log -Level Error -Message "Unable to create a valid session" -Path $LogFile
        exit 1
    }

    if (($Sessions.Count/$SessionsToOpen) -ne $UserName.Count) {
        Write-Warning "Unable to create all possible sessions for all admin accounts!"
        Write-Log -Level Warn -Message "Unable to create all possible sessions for all admin accounts!" -Path $LogFile
        $p = Read-Host -Prompt "Press 'a' to abort, any other key to continue"
        if ($p.Trim().ToLower() -eq 'a') {
            foreach ($s in $Sessions) {
                $s.Value | Remove-PSSession -ErrorAction SilentlyContinue
            }
            Write-Log -Level Info -Message "Aborting script" -Path $LogFile
            exit 1
        }
    }

    Write-Host "Successfully created $($Sessions.Count) total sessions"
    Write-Log -Level Info -Message "Successfully created $($Sessions.Count) total sessions" -Path $LogFile
    Write-Host ""

    # Get users to migrate
    if (![string]::IsNullOrWhiteSpace($UsersFilePath) -and (Test-Path -Path $UsersFilePath)) {
        $users = Get-Content -Path $UsersFilePath | Where-Object { $_ -match '^[^@]+@.+$' } | Sort-Object -Unique
    }
    else {
        if ($null -eq $FilterScript) {
            $FilterScript = [ScriptBlock]::Create('TeamsUpgradePolicy -eq $null -and Enabled -eq $true')
        }
        $MainSession = $Sessions[0].Value
        $users = Invoke-Command -Session $MainSession -ScriptBlock { Get-CsOnlineUser -Filter $using:FilterScript } | Select-Object -ExpandProperty UserPrincipalName | Sort-Object -Unique
    }

    Write-Host "Found $($Users.Count) users"
    Write-Log -Level Info -Message "Found $($Users.Count) users" -Path $LogFile
    Write-Host ""

    # Split into groups for running per session
    $SessionSize = [Math]::Floor($Users.Count / $Sessions.Count)
    $SessionBatchesSizes = @()
    foreach ($s in $Sessions) {
        $SessionBatchesSizes += $SessionSize
    }
    if (($SessionSize * $Sessions.Count) -ne $Users.Count) {
        # if total users is not evenly divisible, add the remainder of users to the first batch
        $SessionBatchesSizes[0] += $Users.Count - ($SessionSize * $Sessions.Count)
    }
    $Batches = @(foreach ($s in $SessionBatchesSizes) {
            , @($users | Select-Object -First $s)        # return a forced array of user objects (that is the leading comma)
            $users = $users | Select-Object -Skip $s    # remove batched users from main user group for further processing
        })

    $BaseScript = [ScriptBlock]::Create(@'
param(
    $FunctionStrings,
    $users,
    $ExpirationOffsetMinutes,
    $OAuthToken,
    $OverrideAdminDomain,
    $RemoteScript,
    [Object[]] $OtherArgs
)
Import-Module -Name MicrosoftTeams -RequiredVersion 1.1.6 -ErrorAction Stop -InformationVariable $null

$RemoteScript = [ScriptBlock]::Create($RemoteScript.ToString())

$FunctionStrings | Invoke-Expression

if (IsExpired $OAuthToken $ExpirationOffsetMinutes) {
    [PsCustomObject]@{
        Retry = $true
        TokenExpired = $true
        UsersRemaining = $users
        UsersCompleted = @()
        Output = $null
        Error = $null
    }
    exit
}

$Session = Get-CsOnlineSessionFromTokens -TeamsTokens $OAuthToken -ErrorAction Stop
$filter = [Text.StringBuilder]::new()
$count = 0
$Results = [Collections.Generic.List[object]]::new()
$UsersRemaining = $users
$batchUsers = [Collections.Generic.List[object]]::new()
foreach ($user in $users) {
    $batchUsers.Add($user) | Out-Null
    # Escape single quotes in UPN
    $user = $user -replace "'", "''"
    $filter.Append("(UserPrincipalName -eq '$user')") | Out-Null
    $count++
    # maximum filter string length is 32k
    # UPN is a maximum of 1024 characters, other characters total 31, so leaving enough space to ensure no conflict
    if ($filter.Length -gt 30950 -or $count -eq $users.Count) {
        if ((IsExpired $OAuthToken $ExpirationOffsetMinutes) -or $Session.Runspace.RunspaceStateInfo.State -ne 'Opened') {
            [PsCustomObject] @{
                Retry = $true
                UsersRemaining = $UsersRemaining
                UsersCompleted = $UsersCompleted
                Output = $Results
                Error = $null
            }
            # Tear down PSSession
            $Session | Remove-PSSession
            exit
        }
        Write-Host "Sending batch request of $($batchUsers.Count) users..."
        $filterSB = [ScriptBlock]::Create($filter.ToString())
        try {
            $currentResult = @(Invoke-Command -Session $Session -ScriptBlock $RemoteScript -ArgumentList @(@($filterSB) + $OtherArgs) -ErrorAction Stop)
            $Results.AddRange($currentResult) | Out-Null
        } catch {
            if ($_.Exception.Message.Contains("AADSTS50013") -or $_.Exception.Message.Contains("is not equal to Open")) {     # InvalidAssertion or closed session
                # token issue, retry get new token
                [PsCustomObject] @{
                    Retry = $true
                    UsersRemaining = $UsersRemaining
                    UsersCompleted = $UsersCompleted
                    Output = $Results
                    Error = $_.Exception
                }
            } else {
                # return non-finished users
                [PsCustomObject] @{
                    Retry = $false
                    UsersRemaining = $UsersRemaining
                    UsersCompleted = $UsersCompleted
                    Output = $Results
                    Error = $_.Exception
                }
            }
            $Session | Remove-PSSession
            exit
        } finally {
            Write-Host "Batch request sent."
            Write-Host ""
        }
        # Reset
        $UsersRemaining = $UsersRemaining | Where-Object { $_ -notin $batchUsers }
        $UsersCompleted = @($users | Where-Object { $_ -notin $UsersRemaining })
        $batchUsers.Clear() | Out-Null
        $filter.Clear() | Out-Null
    } else {
        $filter.Append(" -or ") | Out-Null
    }
}
Write-Host "Finished sending requests for all $($UsersCompleted.Count) users."
# Tear down PSSession
$Session | Remove-PSSession
[PsCustomObject] @{
    Retry = $false
    UsersRemaining = $UsersRemaining
    UsersCompleted = $UsersCompleted
    Output = $Results
    Error = $null
}
'@)
    # Get local functions for import to script block
    $FunctionStrings = GetUsedLocalFunctions -Script $BaseScript

    # Set Up Jobs
    $RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, $Sessions.Count)
    $RunspacePool.Open()

    $Jobs = [Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $Sessions.Count; $i++) {
        $AdminUser = $Sessions[$i].Key
        $authToken = $Tokens | Where-Object { $_.AccountId -eq $AdminUser }
        $Sessions[$i].Value | Remove-PSSession
        if ($null -eq $Batches[$i] -or $Batches[$i].Count -le 0) { continue }
        Write-Host "Starting job for $($Batches[$i].Count) Users"
        Write-Log -Level Info -Message "Starting job for $($Batches[$i].Count) Users" -Path $LogFile

        $ArgHash = @{
            FunctionStrings         = $FunctionStrings
            users                   = $Batches[$i]
            ExpirationOffsetMinutes = $ExpirationOffsetMinutes
            OAuthToken              = $authToken
            OverrideAdminDomain     = $OverrideAdminDomain
            RemoteScript            = $JobScript.ToString()
            OtherArgs               = $OtherArgs
        }

        $newJob = [PSCustomObject]@{
            AdminUser          = $AdminUser
            PowerShellInstance = [PowerShell]::Create()
            Job                = $null
        }
        $newJob.PowerShellInstance.RunspacePool = $RunspacePool
        $newJob.PowerShellInstance.AddScript($BaseScript).AddParameters($ArgHash) | Out-Null
        $newJob.Job = $newJob.PowerShellInstance.BeginInvoke()
        $Jobs.Add($newJob) | Out-Null
    }

    Write-Host ""
    Write-Host "All jobs started, awaiting job results"
    Write-Log -Level Info -Message "All jobs started, awaiting job results" -Path $LogFile

    do {
        $RunningJobs = @($Jobs | Where-Object { !$_.Job.IsCompleted })
        $FinishedJobs = @($Jobs | Where-Object { $_.Job.IsCompleted })
        foreach ($finishedJob in $FinishedJobs) {
            $adminUser = $finishedJob.AdminUser
            try {
                $currentResults = $finishedJob.PowerShellInstance.EndInvoke($finishedJob.Job)
            }
            catch {
                Write-Warning $_.Exception.Message
                Write-Log -Level Error -Message "Job exit failed.`r`nException:$($_.Exception.Message)" -Path $LogFile
                $inner = $_.Exception.InnerException
                while ($null -ne $inner) {
                    Write-Error "Inner Exception:$($inner.Message)"
                    Write-Log -Level Error -Message "Inner Exception:$($inner.Message)" -Path $LogFile
                    $inner = $inner.InnerException
                }
            }

            $Info = if ($null -ne $finishedJob.PowerShellInstance.Streams.Information) {
                $finishedJob.PowerShellInstance.Streams.Information.ReadAll()
            }
            foreach ($i in $Info) {
                if (![string]::IsNullOrWhiteSpace($i.MessageData)) {
                    if ($i.Source -match 'Microsoft\.Teams\.Config\.psm1$') { continue }
                    Write-Host "Job $($i.ManagedThreadId) at $($i.TimeGenerated): $($i.MessageData)"
                    Write-Log -Level Info -Message "Job $($i.ManagedThreadId) at $($i.TimeGenerated): $($i.MessageData)" -Path $LogFile
                }
            }

            $poshErr = if ($null -ne $finishedJob.PowerShellInstance.Streams.Error) {
                $finishedJob.PowerShellInstance.Streams.Error.ReadAll()
            }
            foreach ($i in $poshErr) {
                Write-Warning "Unhandled Exception: $($i.Exception.Message)"
                Write-Log -Level Error -Message "Job finished with unhandled exception.`r`nException: $($i.Exception.Message)" -Path $LogFile
                if ($null -ne $i.InvocationInfo){
                    Write-Log -Level Error -Message "Invocation Line: $($i.InvocationInfo.Line)" -Path $LogFile
                    Write-Log -Level Error -Message "Script Line Number: $($i.InvocationInfo.ScriptLineNumber)" -Path $LogFile
                }
                $inner = $i.Exception.InnerException
                while ($null -ne $inner) {
                    Write-Error "Inner Exception: $($inner.Message)"
                    Write-Log -Level Error -Message "Inner Exception: $($inner.Message)" -Path $LogFile
                    $inner = $inner.InnerException
                }
            }

            if ($currentResults.Retry) {
                $currentToken = $Tokens | Where-Object { $_.AccountId -eq $adminUser }
                if (IsExpired $currentToken $ExpirationOffsetMinutes) {
                    # Get new token and update token array for given adminUser (to avoid multiple prompts)
                    Write-Host "The access token has expired, if prompted, please sign in again using the pop up window for $adminUser."
                    Write-Log -Level Warn -Message "The access token has expired, attempting to get new access token for admin account." -Path $LogFile
                    
                    try {
                        $Tokens.Remove($currentToken) | Out-Null
                        $currentToken = Get-TeamsTokens -UserName $adminUser @TeamsEnvironmentParam
                        $Tokens.Add($currentToken) | Out-Null
                    }
                    catch {
                        Write-Error "Token Retrieval failed with $(@($currentResults.UsersRemaining).Count) users left in run.`r`nException: $($_.Exception.Message)"
                        Write-Log -Level Error -Message "Token Retrieval failed with $(@($currentResults.UsersRemaining).Count) users left in run.`r`nException: $($_.Exception.Message)" -Path $LogFile
                        $inner = $_.Exception.InnerException
                        while ($null -ne $inner) {
                            Write-Error "Inner Exception: $($inner.Message)"
                            Write-Log -Level Error -Message "Inner Exception: $($inner.Message)" -Path $LogFile
                            $inner = $inner.InnerException
                        }
                        $currentResult.Error = $true
                    }
                }
                if (!$currentResult.Error){
                    # restart job with new smaller user set and new token
                    Write-Host "Retrying job for $($currentResults.UsersRemaining.Count) Users"
                    Write-Log -Level Info -Message "Retrying job for $($currentResults.UsersRemaining.Count) Users" -Path $LogFile

                    $ArgHash = @{
                        FunctionStrings         = $FunctionStrings
                        users                   = $currentResults.UsersRemaining
                        ExpirationOffsetMinutes = $ExpirationOffsetMinutes
                        OAuthToken              = $currentToken
                        OverrideAdminDomain     = $OverrideAdminDomain
                        RemoteScript            = $JobScript.ToString()
                        OtherArgs               = $OtherArgs
                    }
            
                    $newJob = [PSCustomObject]@{
                        AdminUser          = $AdminUser
                        PowerShellInstance = [PowerShell]::Create()
                        Job                = $null
                    }
                    $newJob.PowerShellInstance.RunspacePool = $RunspacePool
                    $newJob.PowerShellInstance.AddScript($BaseScript).AddParameters($ArgHash) | Out-Null
                    $newJob.Job = $newJob.PowerShellInstance.BeginInvoke()
                    $Jobs.Add($newJob) | Out-Null
                    # to ensure at least one more cycle of do/while runs
                    $RunningJobs += $newJob
                }
            }
            elseif ($null -ne $currentResults.Error) {
                if ($null -ne $currentResult.Error.Message) {
                    Write-Error "Job failed with $(@($currentResults.UsersRemaining).Count) users left in run.`r`nException: $($currentResults.Error.Message)"
                    Write-Log -Level Error -Message "Job failed with $(@($currentResults.UsersRemaining).Count) users left in run.`r`nException: $($currentResults.Error.Message)" -Path $LogFile
                    $inner = $currentResults.Error.InnerException
                    while ($null -ne $inner) {
                        Write-Error "Inner Exception: $($inner.Message)"
                        Write-Log -Level Error -Message "Inner Exception: $($inner.Message)" -Path $LogFile
                        $inner = $inner.InnerException
                    }
                }
                WriteRemaining $currentResults.UsersRemaining
            }

            WriteCompleted $currentResults.UsersCompleted

            # output results from job
            $currentResults.Output | Select-Object -Property * -ExcludeProperty PSComputerName, RunspaceId
            # remove finished job from jobs collection
            $finishedJob.PowerShellInstance.Dispose() | Out-Null
            $Jobs.Remove($finishedJob) | Out-Null
        }

        if ($RunningJobs.Count -gt 0) {
            Write-Verbose "$($RunningJobs.Count) jobs still running"
            Start-Sleep -Milliseconds 100
        }
    } while ($RunningJobs.Count -gt 0)
    $RunspacePool.Close()
    Write-Host "All Jobs completed, Log saved to $((Resolve-Path $LogFile).Path)"
    Write-Log -Level Info -Message "All Jobs completed" -Path $LogFile
}

function IsExpired {
    param (
        # Token as SecureString
        [Parameter(Mandatory = $true, 
            Position = 0, 
            ParameterSetName = "Secure")]
        [SecureString]
        $SecureToken,

        # Token as plaintext
        [Parameter(Mandatory = $true,
            Position = 0,
            ParameterSetName = "Raw")]
        [string]
        $Token,

        # Token as hashtable
        [Parameter(Mandatory = $true,
            Position = 0,
            ParameterSetName = "Hash")]
        [hashtable]
        $TokenHash,

        # Minutes in the future to test for expiration
        [Parameter(Position = 1)]
        [int]$OffsetMinutes = 0
    )
    if ($PSCmdlet.ParameterSetName -eq "Hash") {
        # return true if any are expired
        if ((IsExpired $TokenHash['AadAccessToken'] -OffsetMinutes $OffsetMinutes)) {
            $true
        } elseif ((IsExpired $TokenHash['MsAccessToken'] -OffsetMinutes $OffsetMinutes)) {
            $true
        } elseif ((IsExpired $TokenHash['ConfigAccessToken'] -OffsetMinutes $OffsetMinutes)) {
            $true
        } else {
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
            $jwtString = $result.Split('.')[1]
            $jwt = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String((PadBase64String $jwtString))) | ConvertFrom-Json
            [DateTime]::Now.AddMinutes($OffsetMinutes) -gt [DateTime]::new(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc).AddSeconds($jwt.exp).ToLocalTime()
        }
        $IsExpired
    }
}

function PadBase64String([string] $stringToPad) {
    $PaddingLength = 4 - $stringToPad.Length % 4
    if ($PaddingLength -eq 4) {
        $PaddingLength = 0
    }
    $stringToPad + [String]::new("=",$PaddingLength)
}

function SetupShellDependencies {
    $sp = $null
    $err = $null
    try { $sp = [Microsoft.Open.Teams.CommonLibrary.TeamsPowerShellSession]::SessionProvider } catch { $err = $_.FullyQualifiedErrorId }
    try { $sp = [Microsoft.Open.Teams.CommonLibrary.AzureAccount]::new() } catch { $err = $_.FullyQualifiedErrorId }
    if ($null -ne $err) {
        $err = $null
        try { $sp = [Microsoft.TeamsCmdlets.Powershell.Connect.TeamsPowerShellSession]::SessionProvider } catch { $err = $_.FullyQualifiedErrorId }
        if ($null -ne $err) {
            throw "MicrosoftTeams module is either too old or not loaded, Ensure version 1.1.6 is the only loaded version before running!"
        } else {
            throw "MicrosoftTeams module is updated beyond the supported version for this script. Ensure version 1.1.6 is the only loaded version before running!"
        }
    }
}

function Write-Log {
    [CmdletBinding()]
    Param
    (
        [Parameter(
            Mandatory = $true, 
            ValueFromPipelineByPropertyName = $true)] 
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [string]$Path = '.\PowerShell.log',

        [ValidateSet("Error", "Warn", "Info")]
        [string]$Level = "Info"
    )
 
    process {
        if (!(Test-Path $Path)) {
            New-Item $Path -Force -ItemType File | Out-Null
        }

        $LogTextBuilder = [Text.StringBuilder]::new()
        $LogTextBuilder.Append((Get-Date -Format "yyyy-MM-dd HH:mm:ss")) | Out-Null
        $LogTextBuilder.Append(' ') | Out-Null
        switch ($Level) {
            'Error' {
                $LogTextBuilder.Append('ERROR:') | Out-Null
            }
            'Warn' {
                $LogTextBuilder.Append('WARNING:') | Out-Null
            }
            'Info' {
                $LogTextBuilder.Append('INFO:') | Out-Null
            }
        }
        $LogTextBuilder.Append(' ') | Out-Null
        $LogTextBuilder.Append($Message) | Out-Null

        $LogTextBuilder.ToString() | Out-File -FilePath $Path -Append
    }
}



$LogFolderPath = Resolve-Path $LogFolderPath
$LogDate = Get-Date -Format "yyyyMMdd_HHmmss"
$LogFile = "$LogFolderPath\CsTeamsMobilityPolicyBatch_$LogDate.log"

Write-Log -Level Info -Path $LogFile -Message "Grant-CsTeamsMobilityPolicyBatch"

if ($IsTest) {
    $IsTestWarning = "This is running as a test. Please add -IsTest `$false to your command in order to actually perform this action."
    Write-Warning -Message $IsTestWarning
    Write-Log -Level Warn -Path $LogFile -Message $IsTestWarning
}

$FilterScript = [ScriptBlock]::Create('Enabled -eq $true')
$JobScript = [ScriptBlock]::Create(@'
param($filterSB, $ArgHash)
Get-CsOnlineUser -Filter $filterSB | Grant-CsTeamsMobilityPolicy @ArgHash
'@)

$ArgHash = @{
    PolicyName = $PolicyName
    WhatIf = $IsTest
}
$BatchParams = @{
    FilterScript  = $FilterScript
    JobScript     = $JobScript
    UserName      = $UserName
    OtherArgs     = @($ArgHash)
    LogFolderPath = $LogFolderPath
    LogEUII       = $PSBoundParameters.ContainsKey("LogEUII") -and $LogEUII
    LogFile       = $LogFile
}
if ($PSBoundParameters.ContainsKey('UsersFilePath')) {
    $BatchParams['UsersFilePath'] = $UsersFilePath
}
if ($PSBoundParameters.ContainsKey('TeamsEnvironmentName')) {
    $BatchParams['TeamsEnvironmentName'] = $TeamsEnvironmentName
}


Invoke-CsOnlineBatch @BatchParams
