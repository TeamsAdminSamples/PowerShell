<#
  .SYNOPSIS
  Migrate Skype meeting to Teams using EWS.

  .DESCRIPTION
  The ExchOnPremMeetingMigration.ps1 script updates Skype for Business Meeting to Teams Meeting

  .INPUTS
  ExchOnPremMeetingMigration.ps1 need a user email address

  .OUTPUTS
  ExchOnPremMeetingMigration.ps1 generate a number of meetings migrated

  .EXAMPLE
  PS> .\ExchOnPremMeetingMigration.ps1

#>



## Load Managed API dll  
$dllpath = "D:\home\site\wwwroot\EwsApi\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath) 

### Variables
#As an example, an azure function received content from power automate ($Request)
$requestBody = Get-Content $Request -Raw | ConvertFrom-Json 
## Get the Mailbox to Access from the 1st commandline argument
$MailboxName = $requestBody.EmailAddress

# add your Tenant Id
$TenantId = ""
# Exchange Admin Account who impersonate user mailbox, i used key vault in this script
$username = $env:AdminContosolab
$password = $env:AdminPasswordContosolab

## Create Exchange Service Object  
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)   
$creds = New-Object System.Net.NetworkCredential($username,$password)  
$service.Credentials = $creds      

  
# CAS URL Hardcoded  
# Configure your ews url 
$uri=[system.URI] "https://webmail.contosolab.fr/ews/exchange.asmx"  
$service.Url = $uri
# Configure your Skype Meeting Url
$SkypeMeetingUrl = "meet.contosolab.fr"

# GetUserId retrieve the UserId. we need the userid to create Teams Meeting
# Configure ClientId, ClientSecret and TenantId
function GetUserId{
    param
    (
           [Parameter(Mandatory=$true)]
           $upn
    )

    #This is the key of the registered AzureAD app
    $ClientID=""
    $ClientSecret= ""

    #This is your Office 365 Tenant Domain Name or Tenant Id
    $TenantId = ""



    #-------------------- Get Access Token for UserId  ----------## 
    $AccessToken = $null
     try {
        if($null -eq $AccessToken){
            $Body = @{client_id=$ClientID;client_secret=$ClientSecret;grant_type="client_credentials";scope="https://graph.microsoft.com/.default";}
            $OAuthReq = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $Body
            $AccessToken = $OAuthReq.access_token
        }
    }
    catch {
        if($null -ne $AccessToken){
            return $AccessToken
        }
    }


    $headers= @{'Authorization' = "Bearer $AccessToken"
                'Content-Type'='application/json'
                'Accept'='application/json'   }                               
    


    $apiurlUserId ="https://graph.microsoft.com/v1.0/users/$upn ?$select=id"
    $GraphUser = Invoke-RestMethod -Headers $headers -Uri $apiurlUserId -Method get
    return $GraphUser
}

# CreateTeamsMeeting will create Teams Meeting with OnlineMeetings API with application right
# Configure ClientId, ClientSecret and TenantId
# Don't forget Grant application access Policy https://docs.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy
function CreateTeamsMeeting{
    param(
        [Parameter(Mandatory=$true)]$Upn,
        [Parameter(Mandatory=$true)]$Subject,
        [Parameter(Mandatory=$true)]$StartTime,
        [Parameter(Mandatory=$true)]$EndTime
    )
    #This is the key of the registered AzureAD app
    $ClientID=""
    $ClientSecret= ""
    #This is your Office 365 Tenant Domain Name or Tenant Id
    $TenantId = ""



    #-------------------- Get Access Token for OnlineMeetings  ----------## 
    $AccessToken = $null
     try {
        if($null -eq $AccessToken){
            $Body = @{client_id=$ClientID;client_secret=$ClientSecret;grant_type="client_credentials";scope="https://graph.microsoft.com/.default";}
            $OAuthReq = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $Body
            $AccessToken = $OAuthReq.access_token
        }
    }
    catch {
        if($null -ne $AccessToken){
            return $AccessToken
        }
    }


    $headers= @{'Authorization' = "Bearer $AccessToken"
                'Content-Type'='application/json'
                'Accept'='application/json'                                  
    }

    $Body = @{
      startDateTime=$StartTime
      endDateTime=$EndTime
      subject=$Subject
    }
    $BodyJson = $Body | ConvertTo-Json
    $userid = (GetUserId -upn $Upn).id
    $apiurlOnlineMeetings ="https://graph.microsoft.com/beta/users/$userid/onlineMeetings"
    $GraphOnlineMeetings = Invoke-RestMethod -Headers $headers -Uri $apiurlOnlineMeetings -Method post -Body $BodyJson
    return $GraphOnlineMeetings
    return $upn
}
Function TrustAllCerts(){
    # Implement call-back to override certificate handling (and accept all)
    $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
    $Compiler=$Provider.CreateCompiler()
    $Params=New-Object System.CodeDom.Compiler.CompilerParameters
    $Params.GenerateExecutable=$False
    $Params.GenerateInMemory=$True
    $Params.IncludeDebugInformation=$False
    $Params.ReferencedAssemblies.Add("System.dll") | Out-Null

    $TASource=@'
        namespace Local.ToolkitExtensions.Net.CertificatePolicy {
        public class TrustAll : System.Net.ICertificatePolicy {
            public TrustAll()
            {
            }
            public bool CheckValidationResult(System.Net.ServicePoint sp,
                                                System.Security.Cryptography.X509Certificates.X509Certificate cert, 
                                                System.Net.WebRequest req, int problem)
            {
                return true;
            }
        }
        }
'@ 
    $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
    $TAAssembly=$TAResults.CompiledAssembly

    ## We now create an instance of the TrustAll and attach it to the ServicePointManager
    $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
    [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
}

  
## Impersonate 
$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 

# Bind to the Calendar Folder
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)   
TrustAllCerts
$Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$Recurring = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Appointment, 0x8223,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean); 
$psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
$psPropset.Add($Recurring)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Duration)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Organizer)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::RequiredAttendees)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsAllDayEvent)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsRecurring)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Body)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::DisplayTo)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::DisplayCc)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Location)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::FirstOccurrence)
$psPropset.add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::LastOccurrence)

$OnlineMeetingInternalLink = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "OnlineMeetingInternalLink", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($OnlineMeetingInternalLink)
$OnlineMeetingExternalLink = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "OnlineMeetingExternalLink", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($OnlineMeetingExternalLink)
$OnlineMeetingConfLink = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "OnlineMeetingConfLink", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($OnlineMeetingConfLink)
$OnlineMeetingConferenceId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "OnlineMeetingConferenceId", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($OnlineMeetingConferenceId)
$ConferenceTelURI = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "ConferenceTelURI", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($ConferenceTelURI)
$OnlineMeetingTollNumber = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "OnlineMeetingTollNumber", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($OnlineMeetingTollNumber)

$UCCapabilities = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "UCCapabilities", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($UCCapabilities)
$UCCapabilitiesLength = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "UCCapabilitiesLength", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($UCCapabilitiesLength)
$UCInband = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "UCInband", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($UCInband)
$UCInbandLength = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "UCInbandLength", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($UCInbandLength)
$UCMeetingSetting = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "UCMeetingSetting", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($UCMeetingSetting)
$UCMeetingSettingSent = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "UCMeetingSettingSent", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($UCMeetingSettingSent)
$UCOpenedConferenceID = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "UCOpenedConferenceID", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($UCOpenedConferenceID)

$SchedulingServiceMeetingOptionsUrl = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "SchedulingServiceMeetingOptionsUrl", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($SchedulingServiceMeetingOptionsUrl)
$SchedulingServiceUpdateUrl = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "SchedulingServiceUpdateUrl", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($SchedulingServiceUpdateUrl)
$SkypeTeamsMeetingUrl = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "SkypeTeamsMeetingUrl", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($SkypeTeamsMeetingUrl)
$SkypeTeamsProperties = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "SkypeTeamsProperties", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($SkypeTeamsProperties)

$ptagCreatorSimpleDispName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "ptagCreatorSimpleDispName", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($ptagCreatorSimpleDispName)
$ptagLastModifierSimpleDispName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "ptagLastModifierSimpleDispName", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($ptagLastModifierSimpleDispName)
$ptagSenderSimpleDispName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "ptagSenderSimpleDispName", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($ptagSenderSimpleDispName)
$ptagSentRepresentingSimpleDispName = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings, "ptagSentRepresentingSimpleDispName", [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropset.add($ptagSentRepresentingSimpleDispName)

$psPropset.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::html;

$AppointmentState = @{0 = "None" ; 1 = "Meeting" ; 2 = "Received" ;4 = "Canceled" ; }

#Define Date to Query 
$StartDate = (Get-Date)
$EndDate = (Get-Date).AddMonths(2)

$RptCollection = @()
  
#Define the calendar view  
$CalendarView = New-Object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,1000)    

$fiItems = $service.FindAppointments($Calendar.Id,$CalendarView)
if($fiItems.Items.Count -gt 0){
	$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
	$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.Item" -as "Type")
	$ItemColl = [Activator]::CreateInstance($type)
	foreach($Item in $fiItems.Items){
		$ItemColl.Add($Item)
	}	
	[Void]$service.LoadPropertiesForItems($ItemColl,$psPropset)  
}
foreach($Item in $ItemColl){      
    if(($item.IsOnlineMeeting -eq $true) -and ($item.body -match $SkypeMeetingUrl ) -and ($item.Organizer.Address -match $MailboxName ) ){
        $RptCollection += $Item
        $Startime = get-date $Item.Start -Format " yyyy-MM-ddTHH:mm:ss.0K"
        $Endtime = get-date $Item.End -Format " yyyy-MM-ddTHH:mm:ss.0K"
        $OnlineMeeting = CreateTeamsMeeting -Upn $MailboxName -StartTime $Startime -EndTime $Endtime -Subject $Item.Subject
        $UserId = $OnlineMeeting.participants.organizer.identity.user.id
        $ThreadId = $OnlineMeeting.chatInfo.threadId
        $ConfId = $OnlineMeeting.chatInfo.threadId -replace ":", "_"
        $UserIdUnformat = $UserId -replace "-", ""
        $TenantIdUnformat = $TenantId -replace "-", ""

   	    #delete Skype Meeting information
		$item.RemoveExtendedProperty($OnlineMeetingInternalLink)
        $item.RemoveExtendedProperty($OnlineMeetingExternalLink)
        $item.RemoveExtendedProperty($OnlineMeetingConfLink)
        $item.RemoveExtendedProperty($OnlineMeetingConferenceId)
        $item.RemoveExtendedProperty($ConferenceTelURI)
        $item.RemoveExtendedProperty($OnlineMeetingTollNumber)
        $item.RemoveExtendedProperty($UCCapabilities)
        $item.RemoveExtendedProperty($UCCapabilitiesLength)
        $item.RemoveExtendedProperty($UCCapabilities)
        $item.RemoveExtendedProperty($UCInband)
        $item.RemoveExtendedProperty($UCInbandLength)
        $item.RemoveExtendedProperty($UCMeetingSetting)
        $item.RemoveExtendedProperty($UCMeetingSettingSent)
        $item.RemoveExtendedProperty($UCOpenedConferenceID)
        $item.RemoveExtendedProperty($ptagCreatorSimpleDispName)
        $item.RemoveExtendedProperty($ptagLastModifierSimpleDispName)
        $item.RemoveExtendedProperty($ptagSenderSimpleDispName)
        $item.RemoveExtendedProperty($ptagSentRepresentingSimpleDispName)
       
        #Teams meeting Informations
        $Content = "<HTML><head></head><Body>"
        $Content += [System.Web.HttpUtility]::UrlDecode($OnlineMeeting.joininformation.content) -replace "data:text/html,`r`n",''
        $Content += "</body></html>"
        $item.body = $Content
        $item.Location = "Microsoft Teams Meeting"
        $NewOnlineMeetingConfLink = "conf:$MailboxName;gruu;opaque=app:conf:focus:id:teams:2:0!" + $ThreadId + "!$UserIdUnformat!$TenantIdUnformat"
		$item.SetExtendedProperty($OnlineMeetingConfLink,$NewOnlineMeetingConfLink )
        $item.SetExtendedProperty($SchedulingServiceMeetingOptionsUrl, "https://teams.microsoft.com/meetingOptions/?organizerId=$UserId&tenantId=$TenantId&threadId=$ConfId&messageId=0&language=en-US")
        $item.SetExtendedProperty($SchedulingServiceUpdateUrl,"https://scheduler.teams.microsoft.com/teams/$TenantId/$UserId/$ConfId/0")
        $item.SetExtendedProperty($SkypeTeamsMeetingUrl,$OnlineMeeting.joinWebUrl)
        $item.SetExtendedProperty($SkypeTeamsProperties,'{"cid":"'+ $ThreadId +'","private":true,"type":0,"mid":0,"rid":0,"uid":null}')
       
        $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite, [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::mode)
    }
	
}   
$JsonOutput = $RptCollection.Count |ConvertTo-Json

# Associate values to output bindings by calling 'Push-OutputBinding'.
# Send response to Power Automate
Out-File -Encoding Ascii -FilePath $response -inputObject $JsonOutput

#Get-PSSession | Remove-PSSession


