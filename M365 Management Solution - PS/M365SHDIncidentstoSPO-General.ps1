<###
Microsoft Service Health Dashboard Items
This script queries the M365 Service Health Dashboard using the Graph API and stores/maintains items up to date on an SPO list
Required permissions for App registration to query SHD via Graph: ServiceMessage.Read.All (Delegated and/or Application type permsissions)
Required permissions for App registartion to query SPO list: can grant permissions to read or write to all sites and just a single SPO site -more info here: https://ashiqf.com/2021/03/15/how-to-use-microsoft-graph-sharepoint-sites-selected-application-permission-in-a-azure-ad-application-for-more-granular-control/
###>

[string]$LogFileDate = (get-date -f d).replace("/","-")
$LogFile = "C:Temp\Microsoft_SHD_Incidents_$LogFileDate.log"
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

### Setup connection variables to Service Health Dashboard using App registration and GRAPH Api - replace tenant and app registration info accordingly
Write-Log "Getting Incidents from the Service Health Dashboad"
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
$Uri = "https://graph.microsoft.com/beta/admin/serviceAnnouncement/issues"
$InvokeResponse = Invoke-RestMethod -Method Get -Uri $Uri -Headers $headerParams
	If(!($InvokeResponse)){Write-Log "Error: Couldn't get Service Health dashboard issues from Microsoft Graph. Exiting";Exit}

$Items = $InvokeResponse.value

$Incidents = $Items
$IncidentsCount = $Incidents.count 
Write-Log "There are $IncidentsCount incidents currently in the SHD"
			
#region Getting Sharepoint Incidents list information
Write-Log "Getting Sharepoint info..."
$appID = "graphappid"
$AppEncryptedSecret = "graphappencryptedsecret"
$Pass = $AppEncryptedSecret | ConvertTo-SecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Pass)
$appSecret = [System.Runtime.InteropServices.Marshal]::PTrToStringAuto($BSTR)
$tokenAuthURI = "https://login.windows.net/<guid-getfromauzre>/oauth2/token"
$requestBody = "grant_type=client_credentials" + 
    "&client_id=$appID" +
    "&client_secret=$appSecret" +
    "&resource=https://graph.microsoft.com/"
	
##Then we use the Token Endpoint URI and pass it the values in the body of the request
$tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenAuthURI -body $requestBody -ContentType "application/x-www-form-urlencoded"

##This response provides our Bearer Token
$accessToken = $tokenResponse.access_token

#Replace the UriINCSP with your actual SPO site URI, you can use Graph Explorer to find this information/URI
$INCSPTable = @()
$UriINCSP = "https://graph.microsoft.com/v1.0/sites/<tenantname>.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/lists/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/items?expand=fields(select=Title,Service,Status,UserImpact,Alert,Class,IncID,UpdatedUTC,Updated,StartTimeUTC,StartTime,EndTimeUTC,EndTime,Twitter,LastWrite,ID,IsResolved,HighImpact,Origin,Feature,FeatureGroup)"
$INCgraphResponse = Invoke-RestMethod -Method Get -Uri $UriINCSP -Headers @{"Authorization"="Bearer $accessToken"} -ErrorAction Stop

#This section will get rows from Microsoft Incidents SP list
$INCSPLrows = $INCgraphResponse.value | Select-Object -ExpandProperty Fields | select-Object Title,Service,Status,UserImpact,Alert,Class,IncID,UpdatedUTC,Updated,StartTimeUTC,StartTime,EndTimeUTC,EndTime,Twitter,LastWrite,ID,IsResolved,HighImpact,Origin,Feature,FeatureGroup,'@odata.etag'
$INCSPTable+= $INCSPLrows

Do{

    $INCgraphResponse = Invoke-RestMethod -Method Get -Uri $INCgraphResponse.'@odata.nextLink' -Headers @{"Authorization"="Bearer $accessToken"} -ErrorAction Stop
    $INCSPLrows = $INCgraphResponse.value | Select-Object -ExpandProperty Fields | select-Object Title,Service,Status,UserImpact,Alert,Class,IncID,UpdatedUTC,Updated,StartTimeUTC,StartTime,EndTimeUTC,EndTime,Twitter,LastWrite,ID,IsResolved,HighImpact,Origin,Feature,FeatureGroup,'@odata.etag'
    $INCSPTable+= $INCSPLrows

}
While($INCgraphResponse.'@odata.nextLink' -ne $null)
If(!($INCSPTable)){Write-Log "Error: Couldn't get items from SharePoint Site. Exiting";Exit}
#filter to unique items as the do/while may add duplicates to our SP table
$INCSPTable = $INCSPTable | sort IncId -unique
$INCSPTableCount = $INCSPTable.count
Write-log "$INCSPTableCount items pulled from the SP list"

#Check to see if the Incidents exist in SPO, if not, add it, if yes, update it if the modified date has changed.
Foreach ($Incident in $Incidents){   
                             
        $SmartDoubleQuotes = '[\u201C\u201D]'
        $Inctitleraw = $Incident.title.Replace('"',"^")
        $Inctitle = $Inctitleraw -Replace $SmartDoubleQuotes,'^'
        $Incservice = $Incident.service
        $Incstatus = $Incident.status
        $Incimpactraw = $Incident.impactDescription.Replace('"',"^")
        $Incimpact = $Incimpactraw -Replace $SmartDoubleQuotes,'^'
        $Incalert = $Incident.classification
        $Incclass = $Incident.classification
        $Incid = $Incident.id
        #$Incseverity = $Incident.Severity
        $IncstarttimeUTC = $Incident.StartDateTime
        $IncStartTime = $Incident.StartDateTime
        $IncLastUpdatedUTC = $Incident.LastModifiedDateTime
        $IncLastUpdated = $Incident.LastModifiedDateTime
        $IncendtimeUTC = $Incident.EndDateTime
        $IncEndTime = $Incident.EndDateTime
		$FinalStatusRaw = If([string]($incident.posts.description.content) -match '(?<=(?i)(Final\sStatus:\s))(.[^\n]+.)'){$matches[0].Replace('"',"^")}
		$FinalStatus = $FinalStatusRaw -Replace $SmartDoubleQuotes,'^'
		$matches = $null
		$RootCauseRaw = If([string]($incident.posts.description.content) -match '(?<=(?i)(Preliminary\sroot\scause:\s)|(Root\scause:\s))(.[^\n]+\.)'){$matches[0].Replace('"',"^")}
		$RootCause = $RootCauseRaw -Replace $SmartDoubleQuotes,'^'
        $IncHighImpact = $Incident.highImpact
        $IncOrigin = $Incident.origin
        $IncFeature = $Incident.feature
        $IncFeatureGroup = $Incident.FeatureGroup
        $IncIsResolved = $Incident.isResolved

        If($INCSPTable.IncID -notcontains $Incid){

            If($IncendtimeUTC -eq $null){
            
            Write-log "Adding new open item to SP list: $Incid"
            $Inctwitter = "False"
		    $IncURI = "https://graph.microsoft.com/v1.0/sites/<tenantname>.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/lists/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/items"
		    $body_patch = "{`"fields`": {`"Title`": `"$Inctitle`",`"Service`": `"$Incservice`",`"Status`": `"$Incstatus`",`"UserImpact`": `"$Incimpact`",`"Alert`": `"$Incalert`",`"HighImpact`": `"$IncHighImpact`",`"Class`": `"$Incclass`",`"IncID`": `"$Incid`",`"UpdatedUTC`": `"$IncLastUpdatedUTC`",`"Updated`": `"$IncLastUpdated`",`"StartTimeUTC`": `"$IncStartTimeUTC`",`"StartTime`": `"$IncStartTime`",`"Twitter`": `"$IncTwitter`",`"IsResolved`": `"$IncIsResolved`",`"Origin`": `"$IncOrigin`",`"Feature`": `"$IncFeature`",`"FeatureGroup`": `"$IncFeatureGroup`"}}"
            Try{$Response_Post = Invoke-RestMethod -Method POST -Uri $IncURI -Headers @{"Authorization"="Bearer $accesstoken"} -ContentType 'application/json' -body $body_patch}
			    Catch{Write-Log "Error: [$($_.exception.message)]";continue }
            }
            else{

                Write-log "Adding new closed item to SP list: $Incid"
                $Inctwitter = "False"
		        $IncURI = "https://graph.microsoft.com/v1.0/sites/<tenantname>.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/lists/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/items"
		        $body_patch = "{`"fields`": {`"Title`": `"$Inctitle`",`"Service`": `"$Incservice`",`"Status`": `"$Incstatus`",`"UserImpact`": `"$Incimpact`",`"Alert`": `"$Incalert`",`"HighImpact`": `"$IncHighImpact`",`"Class`": `"$Incclass`",`"IncID`": `"$Incid`",`"UpdatedUTC`": `"$IncLastUpdatedUTC`",`"Updated`": `"$IncLastUpdated`",`"StartTimeUTC`": `"$IncStartTimeUTC`",`"StartTime`": `"$IncStartTime`",`"EndTimeUTC`": `"$IncEndTimeUTC`",`"EndTime`": `"$IncEndTime`",`"FinalStatus`": `"$FinalStatus`",`"PreliminaryRootCause`": `"$RootCause`",`"Twitter`": `"$IncTwitter`",`"IsResolved`": `"$IncIsResolved`",`"Origin`": `"$IncOrigin`",`"Feature`": `"$IncFeature`",`"FeatureGroup`": `"$IncFeatureGroup`"}}"
                Try{$Response_Post = Invoke-RestMethod -Method POST -Uri $IncURI -Headers @{"Authorization"="Bearer $accesstoken"} -ContentType 'application/json' -body $body_patch}
			        Catch{Write-Log "Error: [$($_.exception.message)]";continue }

                }

         }
         else{

            $INCSPLrowsfiltered = $null
            $INCSPLrowsfiltered = $INCSPTable | Where {$_.IncID -contains $Incid}

            Foreach($IncSPLrowfiltered in $INCSPLrowsfiltered){

                If($IncSPLrowfiltered.UpdatedUTC -ne $IncLastUpdatedUTC){


                    Write-log "Getting Sharepoint online list ID for the row $Incid"
                    $itemID = $null
                    $itemID = $IncSPLrowfiltered.ID
                    $IncLastWrite = "MSUpdated"
         
                    If($IncendtimeUTC -eq $null){

         	        Write-log "Updating the sharepoint row from Microsoft open Incident item $Incid"
		            $IncURI = "https://graph.microsoft.com/v1.0/sites/<tenantname>.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/lists/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/items/"+$itemID+"/fields"
		            $body_patch = "{`"Title`": `"$Inctitle`",`"Service`": `"$Incservice`",`"Status`": `"$Incstatus`",`"UserImpact`": `"$Incimpact`",`"Alert`": `"$Incalert`",`"HighImpact`": `"$IncHighImpact`",`"Class`": `"$Incclass`",`"IncID`": `"$Incid`",`"UpdatedUTC`": `"$IncLastUpdatedUTC`",`"Updated`": `"$IncLastUpdated`",`"StartTimeUTC`": `"$IncStartTimeUTC`",`"StartTime`": `"$IncStartTime`",`"LastWrite`": `"$Inclastwrite`",`"IsResolved`": `"$IncIsResolved`",`"Origin`": `"$IncOrigin`",`"Feature`": `"$IncFeature`",`"FeatureGroup`": `"$IncFeatureGroup`"}"
		            Try{$Response_Post = Invoke-RestMethod -Method PATCH -Uri $IncURI -Headers @{"Authorization"="Bearer $accesstoken"} -ContentType 'application/json' -body $body_patch}
			            Catch{Write-Log "Error: [$($_.exception.message)]";continue }
                    }
                    else{

                        Write-log "Updating the sharepoint row from Microsoft closed Incident item $Incid"
		                $IncURI = "https://graph.microsoft.com/v1.0/sites/<tenantname>.sharepoint.com,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx,xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/lists/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/items/"+$itemID+"/fields"
		                $body_patch = "{`"Title`": `"$Inctitle`",`"Service`": `"$Incservice`",`"Status`": `"$Incstatus`",`"UserImpact`": `"$Incimpact`",`"Alert`": `"$Incalert`",`"HighImpact`": `"$IncHighImpact`",`"Class`": `"$Incclass`",`"IncID`": `"$Incid`",`"UpdatedUTC`": `"$IncLastUpdatedUTC`",`"Updated`": `"$IncLastUpdated`",`"StartTimeUTC`": `"$IncStartTimeUTC`",`"StartTime`": `"$IncStartTime`",`"EndTimeUTC`": `"$IncEndTimeUTC`",`"EndTime`": `"$IncEndTime`",`"FinalStatus`": `"$FinalStatus`",`"PreliminaryRootCause`": `"$RootCause`",`"LastWrite`": `"$Inclastwrite`",`"IsResolved`": `"$IncIsResolved`",`"Origin`": `"$IncOrigin`",`"Feature`": `"$IncFeature`",`"FeatureGroup`": `"$IncFeatureGroup`"}"
		                Try{$Response_Post = Invoke-RestMethod -Method PATCH -Uri $IncURI -Headers @{"Authorization"="Bearer $accesstoken"} -ContentType 'application/json' -body $body_patch}
			                Catch{Write-Log "Error: [$($_.exception.message)]";continue }

                        }
                }

            }
         
         }

         Clear-Variable Inctitleraw,Inctitle,Incservice,Incstatus,Incimpactraw,Incimpact,Incalert,Incclass,Incid,IncLastUpdatedUTC,IncLastUpdated,IncstarttimeUTC,Incstarttime,IncendtimeUTC,Incendtime,FinalStatusRaw,FinalStatus,RootCauseRaw,RootCause,matches,IncIsResolved,IncHighImpact,IncOrigin,IncFeature,IncFatureGroup -ErrorAction SilentlyContinue

}

Write-Log "****************************************Script has ended...****************************************"