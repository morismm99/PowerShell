<###
Microsoft Message Center Items
This script queries the M365 Message Center using the Graph API and stores/maintains items up to date on an SPO list
Required permissions for App registration to query SHD via Graph: ServiceHealth.Read.All (Delegated and/or Application type permsissions)
Required permissions for App registartion to query SPO list: can grant permissions to read or write to all sites and just a single SPO site -more info here: https://ashiqf.com/2021/03/15/how-to-use-microsoft-graph-sharepoint-sites-selected-application-permission-in-a-azure-ad-application-for-more-granular-control/
###>

[string]$LogFileDate = (get-date -f d).replace("/","-")
$LogFile = "C:Temp\Microsoft_MessageCenter_$LogFileDate.log"
Function Write-Log {
	Param ([string]$string)
	# Get the current date
	[string]$date = Get-Date -Format G
	
	# Write everything to our log file
	( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
}

### Set TLS to 1.2
[System.Net.ServicePointManager]::SecurityProtocol = 'TLS12'

Write-Log "****************************************Script has started...****************************************"

### Setup connection variables to Message Center using App registration and GRAPH Api - replace tenant and app registration info accordingly
Write-Log "Getting Messages from the Message Center"
$ClientID = "graphappid"
$EncryptedSecret = "encryptedgraphappsecret"
$Pass0 = $EncryptedSecret | ConvertTo-SecureString
$BSTR0 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Pass0)
$clientSecret = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR0)
$loginURL = "https://login.microsoftonline.com/"
$tenantdomain = "contoso.com"
$TenantGUID = "guid-getfromazure"
$resource = "https://graph.microsoft.com"
### Azure auth
$body = @{grant_type="client_credentials";resource=$resource;client_id=$ClientID;client_secret=$ClientSecret}
$oauth = Invoke-RestMethod -Method Post -Uri $loginURL/$tenantdomain/oauth2/token?api-version=1.0 -Body $body
$headerParams = @{'Authorization'="$($oauth.token_type) $($oauth.access_token)"}
### URL and invoke to receive MS service health dashboard data
$Uri = "https://graph.microsoft.com/beta/admin/serviceAnnouncement/Messages"
$InvokeResponse = Invoke-RestMethod -Method Get -Uri $Uri -Headers $headerParams
	If(!($InvokeResponse)){Write-Log "Error: Couldn't get Service Health dashboard issues from Microsoft Graph. Exiting";Exit}

$Items = $InvokeResponse.value

$Messages = $Items
$Messagescount = $Messages.count
Write-Log "There are $Messagescount messages currently in the Message Center"
			
#region Getting Sharepoint Messages list information
Write-Log "Getting Sharepoint info..."
$appID = "graphappid"
$AppEncryptedSecret = "encryptedappsecret"
$Pass = $AppEncryptedSecret | ConvertTo-SecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Pass)
$appSecret = [System.Runtime.InteropServices.Marshal]::PTrToStringAuto($BSTR)
$tokenAuthURI = "https://login.windows.net/<guid-getfromazure>/oauth2/token"
$requestBody = "grant_type=client_credentials" + 
    "&client_id=$appID" +
    "&client_secret=$appSecret" +
    "&resource=https://graph.microsoft.com/"
	
##Then we use the Token Endpoint URI and pass it the values in the body of the request
$tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenAuthURI -body $requestBody -ContentType "application/x-www-form-urlencoded"

##This response provides our Bearer Token
$accessToken = $tokenResponse.access_token

#Replace the URIMSGSP with your actual SPO site URI, you can use Graph Explorer to find this information/URI
$MSGSPTable = @()
$UriMSGSP = "https://graph.microsoft.com/v1.0/sites/<tenantname>.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/lists/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/items?expand=fields(select=Title,MessageCenterID,Link,Description,SolutionsAffected,ActByDate,Category,Status,Severity,CreatedDate,ModifiedDate,EndDate,ModifiedDateUTC,RoadMapID,ID,IsMajorChange,Tags)"
$MSGgraphResponse = Invoke-RestMethod -Method Get -Uri $UriMSGSP -Headers @{"Authorization"="Bearer $accessToken"} -ErrorAction Stop

#This section will get rows from Microsoft Message Center SPO list
$MSGSPLrows = $MSGgraphResponse.value | Select-Object -ExpandProperty Fields | select Title,Link,MessageCenterID,Description,SolutionsAffected,ActByDate,Category,ActionType,MSStatus,Status,Severity,CreatedDate,ModifiedDate,EndDate,ModifiedDateUTC,RoadMapID,ID,IsMajorChange,Tags,'@odata.etag'
$MSGSPTable+= $MSGSPLrows

Do{

    $MSGgraphResponse = Invoke-RestMethod -Method Get -Uri $MSGgraphResponse.'@odata.nextLink' -Headers @{"Authorization"="Bearer $accessToken"} -ErrorAction Stop
    $MSGSPLrows = $MSGgraphResponse.value | Select-Object -ExpandProperty Fields | select Title,Link,MessageCenterID,Description,SolutionsAffected,ActByDate,Category,ActionType,MSStatus,Status,Severity,CreatedDate,ModifiedDate,EndDate,ModifiedDateUTC,RoadMapID,ID,IsMajorChange,Tags,'@odata.etag'
    $MSGSPTable+= $MSGSPLrows

}
While($MSGgraphResponse.'@odata.nextLink' -ne $null)
If(!($MSGSPTable)){Write-Log "Error: Couldn't get items from SharePoint Site. Exiting";Exit}
#filter to unique items as the do/while may add duplicates to our SPO table
$MSGSPTable = $MSGSPTable | sort MessageCenterID -unique
$MSGSPTableCount = $MSGSPTable.count
Write-log "$MSGSPTableCount items pulled from the SP list"

#Check to see if the Message exist in SPO, if not, add it, if yes, update it if the modified date has changed.
Foreach ($Message in $Messages){   

        $Msgid = $Message.id
        $SmartDoubleQuotes = '[\u201C\u201D]'
        $Msgtitleraw = $Message.title.Replace('"',"^")
        $Msgtitle = $Msgtitleraw -Replace $SmartDoubleQuotes,'^'
        $Msgdescriptionraw = $Message.body.content.Replace('"',"^")        
        $Msgdescription = $Msgdescriptionraw -Replace $SmartDoubleQuotes,'^'
        $Msgseverity = $Message.severity
        $MsgCategory = $Message.category
        #$MsgActionType = $Message.Actiontype
        #$MsgMSStatus = $Message.Status
        $Msglink = "https://portal.office.com/AdminPortal/home?switchtomodern=true#/MessageCenter?id=" + $Msgid
        $MsgWorkloads = $Message.services.split("`n") -join ", "
        $MsgStartTime = $Message.startDateTime
        $MsgLastUpdatedTime = $Message.lastModifiedDateTime
        $MsgLastUpdatedUTC = $Message.lastModifiedDateTime
        $MsgEndTime = $Message.endDateTime
        #$MsgMileStoneDate = $Message.MileStoneDate
		    $RMidList = @()
		    Select-String "((?i)(?<=featureid=)(\d{5})(?=\`"))|((?i)(?<=searchterms=)(\d{5})(?=\^))" -input $Msgdescriptionraw -AllMatches | Foreach {$RMidList += ($_.matches.value)}
		$MsgRoadMapID = ($RMidList | Select-Object -Unique) -join ","
		#if($MsgMileStoneDate -eq $null){$MsgMileStoneDate = $MsgStartTime}
        $MsgActionRequiredByDate = $Message.actionRequiredByDateTime
        if($MsgActionRequiredByDate -eq $null){$MsgActionRequiredByDate = $MsgStartTime}
        $MsgIsMajorChange = $Message.isMajorChange
        $MsgTags = $Message.tags.split("`n") -join ", "

        If($MSGSPTable.MessageCenterID -notcontains $Msgid){

            Write-log "Adding new item to SP list: $Msgid"
		    $MSGURI = "https://graph.microsoft.com/v1.0/sites/<tenantname>.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/lists/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/items"
		    $body_patch = "{`"fields`": {`"Title`": `"$Msgtitle`",`"MessageCenterID`": `"$Msgid`",`"Link`": `"$Msglink`",`"Description`": `"$Msgdescription`",`"SolutionsAffected`": `"$MsgWorkloads`",`"ActByDate`": `"$MsgActionRequiredByDate`",`"Category`": `"$MsgCategory`",`"Severity`": `"$MsgSeverity`",`"CreatedDate`": `"$MsgStartTime`",`"ModifiedDate`": `"$MsgLastUpdatedTime`",`"EndDate`": `"$MsgEndTime`",`"ModifiedDateUTC`": `"$MsgLastUpdatedUTC`",`"RoadMapID`": `"$MsgRoadMapID`",`"IsMajorChange`": `"$MsgIsMajorChange`",`"Tags`": `"$MsgTags`"}}"
            Try{$Response_Post = Invoke-RestMethod -Method POST -Uri $MSGURI -Headers @{"Authorization"="Bearer $accesstoken"} -ContentType 'application/json' -body $body_patch}
			    Catch{Write-Log "Error: [$($_.exception.message)]";continue }

         }
         else{

            $MSGSPLrowsfiltered = $null
            $MSGSPLrowsfiltered = $MSGSPTable | Where {$_.MessageCenterID -contains $Msgid}

            Foreach($MSGSPLrowfiltered in $MSGSPLrowsfiltered){

                If($MSGSPLrowfiltered.ModifiedDateUTC -ne $MsgLastUpdatedUTC){

                    $msgCustomStatus = "Updated"

                    Write-log "Getting Sharepoint online list ID for the row $Msgid"
                    $itemID = $null
                    $itemID = $MSGSPLrowfiltered.ID   

         	        Write-log "Updating the sharepoint row from Microsoft MessageCenter item $Msgid"
		            $MSGURI = "https://graph.microsoft.com/v1.0/sites/<tenantname>.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/lists/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/items/"+$itemID+"/fields"
		            $body_patch = "{`"Title`": `"$Msgtitle`",`"MessageCenterID`": `"$Msgid`",`"Link`": `"$Msglink`",`"Description`": `"$Msgdescription`",`"SolutionsAffected`": `"$MsgWorkloads`",`"ActByDate`": `"$MsgActionRequiredByDate`",`"Category`": `"$MsgCategory`",`"Status`": `"$MsgCustomStatus`",`"Severity`": `"$MsgSeverity`",`"CreatedDate`": `"$MsgStartTime`",`"ModifiedDate`": `"$MsgLastUpdatedTime`",`"EndDate`": `"$MsgEndTime`",`"ModifiedDateUTC`": `"$MsgLastUpdatedUTC`",`"RoadMapID`": `"$MsgRoadMapID`",`"IsMajorChange`": `"$MsgIsMajorChange`",`"Tags`": `"$MsgTags`"}"
		            Try{$Response_Post = Invoke-RestMethod -Method PATCH -Uri $MSGURI -Headers @{"Authorization"="Bearer $accesstoken"} -ContentType 'application/json' -body $body_patch}
			            Catch{Write-Log "Error: [$($_.exception.message)]";continue }
                }

            }
         
         }
         
         Clear-Variable MSgid,Msgtitleraw,Msgtitle,Msgdescriptionraw,Msgdescription,MsgCategory,Msgseverity,Msglink,Msgseverity,MsgWorkloads,MsgStartTime,MsgLastUpdatedTime,MsgLastUpdatedUTC,MSGEndTime,MSGActionREquiredByDate,MSGRoadMapID,MSGIsMajorChange,MSGTags -ErrorAction SilentlyContinue

}

Write-Log "****************************************Script has ended...****************************************"