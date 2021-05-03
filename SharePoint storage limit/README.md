# SharePoint storage limit

# Description

The script is to modify the SharePoint storage limit and also to set the StorageQuotaWarningLevel, please provide the storage quota and StorageQuotaWarningLevel inputs in MB

You must be a SharePoint Online Administrator or Global Administrator and be a site collection administrator to run the script

### What is Storage

Each team in Microsoft Teams has a team site in SharePoint Online, and each channel in a team gets a folder within the default team site document library. Files shared within a conversation are automatically added to the document library, and permissions and file security options set in SharePoint are automatically reflected within Teams

# Example

     Set-SPOSite -Identity https://contoso.sharepoint.com/sites/HrTeam -StorageQuota 1500 -StorageQuotaWarningLevel 1400

Example updates the HR site storage limit to 1500 MB and storage quota warning level to 1400 MB 

# Inputs

 $SiteName Ex:https://contoso.sharepoint.com/sites/HrTeam
 
 $StorageQuota Ex:1500 (input is in MB)
 
 $StorageQuotaWarningLevel Ex:1400 (input is in MB)
 
# Parameters

**`-StorageQuota`**

Specifies the storage quota in megabytes of the site collection
- - -
Type:	Int64
- - -
Position:	Named
- - -
Default value:	None
- - -
Accept pipeline input:	False
- - -
Accept wildcard characters:	False
- - -
Applies to:	SharePoint Online


**`-StorageQuotaWarningLevel`**

Specifies the warning level in megabytes of the site collection to warn the site collection administrator that the site is approaching the storage quota
- - -
Type:	Int64
- - -
Position:	Named
- - -
Default value:	None
- - -
Accept pipeline input:	False
- - -
Accept wildcard characters:	False
- - -
Applies to:	SharePoint Online
 
# Prerequisites:

 Install [sharepoint-online](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online?view=sharepoint-ps)

# How to run the script

1. To run the script you will need to either download it or copy and paste the **` SharePoint storage limit.ps1`** into PowerShell
2. Provide the SharePoint Online Administrator or Global Administrator or site collection administrator credentials when it prompts
3. Provide the parameters `$SiteName`,`$StorageQuota` and `$StorageQuotaWarningLevel`

# Output

 After executing the script storage quota is updated to 1500 megabytes and the storage quota warning level is updated to 1400 megabytes
 
 A log file will be generated with exceptions, errors along with script execution time
