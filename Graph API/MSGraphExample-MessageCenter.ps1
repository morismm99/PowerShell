## Examples on how to authenticat to MS Graph using Azure AD App Registration and Secret Value
## You can find instructions on how to encrypt the secret value here: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/convertto-securestring?view=powershell-7.3#example-1-convert-a-secure-string-to-an-encrypted-string
## The app being used in this example has the following Application type permissions: ServiceHealth.Read.All, ServiceMessage.Read.All, Sites.Read.All

## Get Messages from the Message Center
$ClientID = "<AppID>"
$EncryptedSecret = "<encriptedsecret>"
$Pass0 = $EncryptedSecret | ConvertTo-SecureString
$BSTR0 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Pass0)
$clientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR0)
#$clientSecret = "<plaintextsecret>"
$loginURL = "https://login.microsoftonline.com/"
$tenantdomain = "<contoso.onmicrosoft.com>"
$TenantGUID = "<tenantguid>"
$resource = "https://graph.microsoft.com"
### Azure auth
$body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
$oauth = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body
$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
### URL and invoke to receive MS service health dashboard data
$Uri = "https://graph.microsoft.com/beta/admin/serviceAnnouncement/Messages"
$InvokeResponse = Invoke-RestMethod -Method Get -Uri $Uri -Headers $headerParams
	#If(!($InvokeResponse)){Write-Log "Error: Couldn't get Service Health dashboard issues from Microsoft Graph. Exiting";Exit}

$Items = $InvokeResponse.value


#Getting Sharepoint Messages list information using a slightly different method to authenticate to MS Graph 
$appID = "<appid>"
$AppEncryptedSecret = "<encriptedsecret>"
$Pass = $AppEncryptedSecret | ConvertTo-SecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Pass)
$appSecret = [System.Runtime.InteropServices.Marshal]::PTrToStringAuto($BSTR)
#$appSecret = "<plaintextsecret>"
#Token Endpoint URI - make sure to replace the Tenant ID with your own
$tokenAuthURI = "https://login.windows.net/<tenantid>/oauth2/token"
$requestBody = "grant_type=client_credentials" + 
    "&client_id=$appID" +
    "&client_secret=$appSecret" +
    "&resource=https://graph.microsoft.com/"
	
##Then we use the Token Endpoint URI and pass it the values in the body of the request
$tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenAuthURI -body $requestBody -ContentType "application/x-www-form-urlencoded"

##This response provides our Bearer Token
$accessToken = $tokenResponse.access_token

#Replace the URIMSGSP with your actual SPO site URI, you can use Graph Explorer to find this information/URI - https://graph.microsoft.com/v1.0/sites?search=keyword
$MSGSPTable = @()
#$UriMSGSP = "https://graph.microsoft.com/v1.0/sites/<tenantname>.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/lists/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/items?expand=fields(select=Title,MessageCenterID,Link,Description,SolutionsAffected,ActByDate,Category,Status,Severity,CreatedDate,ModifiedDate,EndDate,ModifiedDateUTC,RoadMapID,ID,IsMajorChange,Tags)"
$UriMSGSP = "https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/lists/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/items?expand=fields(select=Title,MessageCenterID,Link,Description,SolutionsAffected,ActByDate,Category,Status,Severity,CreatedDate,ModifiedDate,EndDate,ID)"
$MSGgraphResponse = Invoke-RestMethod -Method Get -Uri $UriMSGSP -Headers @{"Authorization"="Bearer $accessToken"} -ErrorAction Stop
$MSGSPTable = $MSGgraphResponse.value


#Get files in a User's OneDrive
$UriOneDrive = "https://graph.microsoft.com/v1.0/users/<usersUPN>/drive/root/children"
$MSGgraphResponse = Invoke-RestMethod -Method Get -Uri $UriOneDrive -Headers @{"Authorization"="Bearer $accessToken"} -ErrorAction Stop
