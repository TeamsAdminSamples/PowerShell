# DomainValidation
# Description
Script fetches the SIP enabled domains from Tenant, skips the DNS query for the domains contain .onmicrosoft.com, for domains which do not contain .onmicrosoft.com does a DNS name query resolution for **Lync discover** records and validate if they are pointing to webdir.online.lync.com 

If all the domains are pointing to Webdir.online.lync.com script displays the **Overall status is Ok**, even if one of the domain is pointing to outside webdir.online.lync.com then script displays the **Overall status is not Ok** message

# Prerequisite
[SFB online connector](https://www.microsoft.com/en-us/download/details.aspx?id=39366)
# Input
Global Administrator or Skype for Business Online administrator account user principal name and password, and then select OK
# Examples
##### Example 1
If domain resolving to webdir.online.lync.com 
###### Output
|Name  |   Status|  ErrorMessage |
|---|----|-----|
|xyz.com     |  OK    | Resolving to webdir.online.lync.com |
##### Example 2
If domain resloving to other than webdir.online.lync.com 
###### Output
|Name      |Status  |ErrorMessage |
|----|---|---|
|xyz.com       |NOT_OK  |Resolving to 52.xxx.12.xx4(IP Address)|
##### Example 3
If the domain is not resolving 
###### Output
|Name     | Status  |ErrorMessage |
|---|---|---|
|xyz.com   |NOT_OK |lyncdiscover.xyz.com : DNS name does not exist|

#### Parameters

`-Type`

Specifies the DNS query type that is to be issued. By default the type is A_AAAA, the A and AAAA types will both be queried.

Type:	RecordType
***
Accepted values:	UNKNOWN, A_AAAA, A, NS, MD, MF, CNAME, SOA, MB, MG, MR, NULL, WKS, PTR, HINFO, MINFO, MX, TXT, RP, AFSDB, X25, ISDN, RT, AAAA, SRV, DNAME, OPT, DS, RRSIG,                        NSEC, DNSKEY, DHCID, NSEC3, NSEC3PARAM, ANY, ALL, WINS|
***
Position:	1
***
Default value:	None
***
Accept pipeline input:	True
***
Accept wildcard characters:	False

# Output
A log file will be generated with exceptions, errors along with script execution time

Output contains

Skip domains list 

Overall Tenant status - Ok/Not Ok

|Domain Name |Status |ErrorMessage|

#### Example
![Sample Output](https://github.com/Geetha63/MS-Teams-Scripts/blob/master/Images/DomainValidation.jpg)
