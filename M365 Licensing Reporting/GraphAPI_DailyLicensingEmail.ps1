#Office 365 Licensing Reporting

### Set TLS to 1.2
[System.Net.ServicePointManager]::SecurityProtocol = 'TLS12'

#Html Header for email to be sent
$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

$global:stopwatch = $null
$global:accessToken = $null

#function to get a token for an app registration that has organization.read.all, directory.read.all and sites.selected permissions
Function Connect-GraphAPI 
{
		### - Getting Azure Access Token and starting timer
		$appID = "<AppRegistrationID>"
        $encryptedsecret = "<encryptedsecret>"
        $Pass = $encryptedsecret | ConvertTo-SecureString
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Pass)
        $appSecret = [System.Runtime.InteropServices.Marshal]::PTrToStringAuto($BSTR)
        $tokenAuthURI = "https://login.windows.net/<AADTenantID>/oauth2/token"
        $requestBody = "grant_type=client_credentials" + 
        "&client_id=$appID" +
        "&client_secret=$appSecret" +
        "&resource=https://graph.microsoft.com/"

		##Then we use the Token Endpoint URI and pass it the values in the body of the request
		$tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenAuthURI -body $requestBody -ContentType "application/x-www-form-urlencoded" 

		##This response provides our Bearer Token
		$tokenResponse.access_token

		$Global:stopwatch =  [system.diagnostics.stopwatch]::StartNew()	
}

#function that does a get from graph using app registration/token from connect-graphapi function above
Function get-GraphAPI ($HTTP) 
{
	if($global:stopwatch -eq $null -OR $global:stopwatch.elapsed.minutes -gt '55'){
		write-host -foregroundcolor yellow "Refreshing Access Token"
		$global:accessToken = Connect-GraphAPI
	}
	Invoke-RestMethod -Method Get -Uri $Http -Headers @{"Authorization"="Bearer $global:accessToken"}
}

$date = (get-date).ToSHortDateString()
$graphResponse = $null
$graphSubscribedSkusTable = @()

#This section will query the subscribed licenses for the tenant using /subscribedSkus Graph endpoint
Do{
	Write-Host -foregroundcolor yellow "Gathering Office 365 subscription License Details"

	$graphSubscribedSkusTable = @()

	if($graphResponse -eq $null){$HTTP = 'https://graph.microsoft.com/v1.0/subscribedSkus'}
	$graphResponse = get-GraphAPI -HTTP $HTTP
	
	if($graphResponse -ne $null){
		$subscribedSkusTable += $graphResponse.value

			foreach($sku in $subscribedSkusTable)
			{
				$remainingUnits = $sku.prepaidUnits.enabled - $sku.consumedUnits
				
				$objaverage = New-object System.Object
				$objaverage | add-member -type NoteProperty -Name Date -value $Date
				$objaverage | add-member -type NoteProperty -Name skuPartNumber -value $sku.skuPartNumber
				$objaverage | add-member -type NoteProperty -Name skuID -value $sku.skuId
				$objaverage | add-member -type NoteProperty -Name consumedUnits -value $sku.consumedUnits
				$objaverage | add-member -type NoteProperty -Name prepaidUnits -value $sku.prepaidUnits.enabled
				$objaverage | add-member -type NoteProperty -Name remainingUnits -value $remainingUnits

				$graphSubscribedSkusTable += $objaverage
				
			if($graphResponse.'@odata.nextLink' -ne $null){$HTTP = $graphResponse.'@odata.nextLink'}
			}
		}
}
While($graphResponse.'@odata.nextLink' -ne $null)

#This section can be used if you need to pull information about group based licensing and what these groups are consuming - it also assumes a hybrid scenario with groups being created on premises and synced to AAD
<#
New-Variable -Name UserTable -Option AllScope -Value @()
New-Variable -Name ContractorTable -Option AllScope -Value @()
New-Variable -Name ServiceAccountTable -Option AllScope -Value @()
New-Variable -Name ProjectTable -Option AllScope -Value @()

Function Get-GroupInfo {

	param($ADGroup, $GroupType, $SkuName, $SkuID)
	
	$Count = ((Get-ADGroup $ADGroup -Server <servername> -Properties member).member).count
	
	If($GroupType -eq 'EndUser'){
	$objaverage = New-object System.Object
	$objaverage | add-member -type NoteProperty -Name 'GroupType' -value $GroupType
	$objaverage | add-member -type NoteProperty -Name 'Group' -value $ADGroup
	$objaverage | add-member -type NoteProperty -Name 'MemberCount' -value $Count
	$objaverage | add-member -type NoteProperty -Name 'SkuName' -value $SkuName
	$objaverage | add-member -type NoteProperty -Name 'SkuID' -value $SkuID
			
	$UserTable += $objaverage
	}

	If($GroupType -eq 'Contractor'){
	$objaverage = New-object System.Object
	$objaverage | add-member -type NoteProperty -Name 'GroupType' -value $GroupType
	$objaverage | add-member -type NoteProperty -Name 'Group' -value $ADGroup
	$objaverage | add-member -type NoteProperty -Name 'MemberCount' -value $Count
	$objaverage | add-member -type NoteProperty -Name 'SkuName' -value $SkuName
	$objaverage | add-member -type NoteProperty -Name 'SkuID' -value $SkuID
			
	$ContractorTable += $objaverage
	}

	If($GroupType -eq 'SVC'){
	$objaverage = New-object System.Object
	$objaverage | add-member -type NoteProperty -Name 'GroupType' -value $GroupType
	$objaverage | add-member -type NoteProperty -Name 'Group' -value $ADGroup
	$objaverage | add-member -type NoteProperty -Name 'MemberCount' -value $Count
	$objaverage | add-member -type NoteProperty -Name 'SkuName' -value $SkuName
	$objaverage | add-member -type NoteProperty -Name 'SkuID' -value $SkuID
			
	$ServiceAccountTable += $objaverage
	}
	
	If($GroupType -eq 'Project'){
	$objaverage = New-object System.Object
	$objaverage | add-member -type NoteProperty -Name 'GroupType' -value $GroupType
	$objaverage | add-member -type NoteProperty -Name 'Group' -value $ADGroup
	$objaverage | add-member -type NoteProperty -Name 'MemberCount' -value $Count
	$objaverage | add-member -type NoteProperty -Name 'SkuName' -value $SkuName
	$objaverage | add-member -type NoteProperty -Name 'SkuID' -value $SkuID
			
	$ProjectTable += $objaverage
	}
}

#>

<#Get group info if using group based licensing - add additional groups as needed
Get-GroupInfo -ADGroup <GroupName> -GroupType EndUser -SkuName "Microsoft 365 E5" -SkuID "06ebc4ee-1bb5-47dd-8120-11324bc54e06"
Get-GroupInfo -ADGroup <GroupName> -GroupType EndUser -SkuName "Microsoft Teams Advanced Comms" -SkuID "e4654015-5daf-4a48-9b37-4f309dddd88b"
Get-GroupInfo -ADGroup <GroupName> -GroupType Contractor -SkuName "Microsoft 365 E5" -SkuID "06ebc4ee-1bb5-47dd-8120-11324bc54e06"
Get-GroupInfo -ADGroup <GroupName> -GroupType Contractor -SkuName "Exchange Online" -SkuID "19ec0d23-8335-4cbd-94ac-6050e30712fa"
Get-GroupInfo -ADGroup <GroupName> -GroupType SVC -SkuName "Microsoft 365 E5" -SkuID "06ebc4ee-1bb5-47dd-8120-11324bc54e06"
Get-GroupInfo -ADGroup <GroupName> -GroupType SVC -SkuName "Teams Meeting Room" -SkuID "6070a4c8-34c6-4937-8dfb-39bbc6397a60"
Get-GroupInfo -ADGroup <GroupName> -GroupType SVC -SkuName "Power Automate Plan 2(Flow_P2)" -SkuID "4755df59-3f73-41ab-a249-596ad72b5504"
Get-GroupInfo -ADGroup <GroupName> -GroupType SVC -SkuName "Power Apps Per User" -SkuID "b30411f5-fea1-4a59-9ad9-3db7c7ead579"
Get-GroupInfo -ADGroup <GroupName> -GroupType SVC -SkuName "Exchange Online" -SkuID "19ec0d23-8335-4cbd-94ac-6050e30712fa"
Get-GroupInfo -ADGroup <GroupName> -GroupType Project -SkuName "Microsoft 365 E5" -SkuID "06ebc4ee-1bb5-47dd-8120-11324bc54e06"
#>

#Can filter to just specific SKUs
#$graphSubscribedSkusTable = $graphSubscribedSkusTable | where {$_.skupartnumber -eq "SPE_E5" -or $_.skupartnumber -eq "EXCHANGEENTERPRISE" -or $_.skupartnumber -eq "ATP_ENTERPRISE" -or $_.skupartnumber -eq "EMS" -or $_.skupartnumber -eq "WIN_DEF_ATP" -or $_.skupartnumber -eq "AAD_PREMIUM" -or $_.skupartnumber -eq "MEETING_ROOM" -or $_.skupartnumber -eq "Flow_P2" -or $_.skupartnumber -eq "VISIOCLIENT" -or $_.skupartnumber -eq "POWERAPPS_PER_USER" -or $_.skupartnumber -eq "ADV_COMMS"} | sort consumedunits -Descending

#Add contents from subscribedskustable to SPO list using Graph API
Foreach ($Sku in $graphSubscribedSkusTable){   

	$skuPartNumber = $Sku.skupartnumber
    $SkuID = $sku.SkuID
    $consumedUnits = $Sku.consumedUnits
    $prepaidUnits = $sku.prepaidUnits
    $remainingUnits = $sku.remainingUnits
    $Date = $sku.Date

	#Use graph explorer to get the site id, in the example below the content after sites/ needs to be replaced with your actual SPO site ID     
	$IncURI = "https://graph.microsoft.com/v1.0/sites/m365x38197403.sharepoint.com,1bfc8d23-a367-4bc9-a677-82392a32be01,ae4b3405-b330-4a16-bab6-473e50a87d60/lists/8c4c40db-10e8-42cb-a934-ce12f44795a0/items"
	$body_patch = "{`"fields`": {`"Title`": `"$skuPartNumber`",`"SkuID`": `"$SkuID`",`"ConsumedUnits`": `"$consumedUnits`",`"PrepaidUnits`": `"$prepaidUnits`",`"RemainingUnits`": `"$RemainingUnits`",`"Date`": `"$Date`"}}"
    $Response_Post = Invoke-RestMethod -Method POST -Uri $IncURI -Headers @{"Authorization"="Bearer $global:accessToken"} -ContentType 'application/json' -body $body_patch
	
	Clear-Variable skuPartNumber,SkuID,consumedUnits,prepaidUnits,remainingUnits,Date

}            
### Email HTML Setup and Send
$EmailBodyHTML = @()

$EmailBodyHTML += $graphSubscribedSkusTable | convertto-html -head $Header
#$EmailBodyHTML += "<br>"
#$EmailBodyHTML += $UserTable | Sort MemberCount -Descending | convertto-html -head $Header
#$EmailBodyHTML += "<br>"
#$EmailBodyHTML += $ContractorTable | Sort MemberCount -Descending | convertto-html -head $Header
#$EmailBodyHTML += "<br>"
#$EmailBodyHTML += $ServiceAccountTable | Sort MemberCount -Descending | convertto-html -head $Header
#$EmailBodyHTML += "<br>"
#$EmailBodyHTML += $ProjectTable | Sort MemberCount -Descending | convertto-html -head $Header

Send-MailMessage -To "<emailaddress>" -From <emailaddress> -BodyAsHtml "$EmailBodyHTML" -Subject "Daily M365 License Stats" -SmtpServer <smtpserver>