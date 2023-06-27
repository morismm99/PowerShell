#This PowerShell script retrieves calendar events for a list of rooms using Microsoft Graph API. 
#It iterates through each room, retrieves the events for a specific time period, and then iterates through each event to extract relevant information 
#such as the room name, subject, start and end time, organizer, and attendees. 
#The extracted information is then stored in a table and added to a list of tables. 

#The script requires an App regisration with Application type permissions set for Calendars.Read

#App Registration information for authentication
$appID = "AppID"
$encryptedsecret = "generateandencryptanappregistrationsecret"
$Pass = $encryptedsecret | ConvertTo-SecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Pass)
$appSecret = [System.Runtime.InteropServices.Marshal]::PTrToStringAuto($BSTR)
$tokenAuthURI = "https://login.windows.net/<tenantID>/oauth2/token"
$requestBody = "grant_type=client_credentials" + 
    "&client_id=$appID" +
    "&client_secret=$appSecret" +
    "&resource=https://graph.microsoft.com/"

##Then we use the Token Endpoint URI and pass it the values in the body of the request
$tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenAuthURI -body $requestBody -ContentType "application/x-www-form-urlencoded"

##This response provides our Bearer Token
$accessToken = $tokenResponse.access_token

##For this report, we are importing a subset of rooms from a csv file
$RoomsCSV = "C:\Temp\rooms.csv"
$Rooms = Import-Csv -Path $RoomsCSV
$Table = @()
$i = 0
$Count = $Rooms.count

#Getting usage data for each room
foreach ($Room in $Rooms) 
{
    $i++
    $UPN = $Room.UserPrincipalName
    Write-Progress -Activity "Getting Room Data" -Status "Working on $UPN, ($i / $count)" -PercentComplete (($i / $count) * 100)
    ##Note: this will pull the last 50 meetings for each room, single meetings and series masters, we can increase that by changing the value for $top= the max is 1000
    ###$Uri = "https://graph.microsoft.com/v1.0/users/$UPN/calendar/events?`$top=50"
    ##Different Uri looking at the calendar view, which also grabs occurrences of series for a given tim frame, set to top 300
    #$Uri = "https://graph.microsoft.com/v1.0/users/$UPN/calendarView?`$top=300&startDateTime=2019-11-01T00:00:00Z&endDateTime=2019-12-31T00:00:00Z"
    #With this URI, we grab everything based on a specific start and end time using a do until loop
    $Uri = "https://graph.microsoft.com/v1.0/users/$UPN/calendarView?startDateTime=2020-06-29T00:00:00Z&endDateTime=2020-12-31T00:00:00Z"

    $Events = @()
    $graphResponse = $null
    $graphResponse = Invoke-RestMethod -Method Get -Uri $Uri -Headers @{"Authorization"="Bearer $accessToken";"Prefer"='outlook.timezone="Central Standard Time"'}
    $Graphdata = $graphResponse.value
    $Events+= $Graphdata

    #Getting all occurrences/meetings for the time period provided
    Do{
    $graphResponse = Invoke-RestMethod -Method Get -Uri $graphResponse.'@odata.nextLink' -Headers @{"Authorization"="Bearer $accessToken"} -ErrorAction Stop
    $Graphdata = $graphResponse.value
    $Events+= $Graphdata
    }
    While($graphResponse.'@odata.nextLink' -ne $null)
    $Events = $Events | Sort-Object id -unique

        Foreach ($Event in $Events){

            $organizer = $event.organizer.emailAddress | Select-Object -ExpandProperty address
            $attendees = $event.attendees
            $attendeecnt = $event.attendees.count
            $start = $event.start.dateTime
            $end = $event.end.dateTime
            $reqattendeeslist = ""
            $optattendeeslist = ""


 ##Foreach loops to split start into a date and time field, note, you could combine this into one foreach, but I like have these separate
            Foreach ($star in $start){

                $startarr = $start -split 'T'
                $startDate = $startarr[0]
                $startTime = $startarr[1]
            }

            Foreach ($en in $end){

                $endarr = $end -split 'T'
                $endDate = $endarr[0]
                $endTime = $endarr[1]      
            }  

 ##if there are more than one attendee, this will list them all        

         if ($attendeecnt -gt 1){
                 foreach ($attendee in $attendees){
                     if ($attendee.type -eq "required"){
                        
                          $reqattendee = $attendee.emailaddress | Select-Object -ExpandProperty address
                          $reqattendeeslist += $reqattendee + ";"

                      }
                     elseif ($attendee.type -eq "optional"){
                         
                          $optattendee = $attendee.emailaddress | Select-Object -ExpandProperty address
                          $optattendeeslist += $optattendee + ";"
                     }

                 }
          }
          else {
                if ($attendees.type -eq "required"){
                     $reqattendee = $attendees.emailaddress | Select-Object -ExpandProperty address
                     $reqattendeeslist += $reqattendee + ";"
                }

                elseif ($attendees.type -eq "optional"){
                    $optattendee = $attendees.emailaddress | Select-Object -ExpandProperty address
                    $optattendeeslist += $optattendee + ";"

                }
        }
        
           
        #Table with data is built
        $Result = "" | Select-Object RoomName,Subject,CreatedDateTime,LastModifiedDateTime,Type,StartDate,StartTime,EndDate,EndTime,Organizer,RequiredAttendees,OptionalAttendees,TotalAttendees
        $Result.RoomName = $Room.DisplayName
        $Result.Subject = $event.subject
        $Result.CreatedDateTime = $event.createdDateTime
        $Result.LastModifiedDateTime = $event.lastModifiedDateTime
        $Result.Type = $event.type
        $Result.StartDate = $startDate
        $Result.StartTime = $startTime
        $Result.EndDate = $endDate
        $Result.EndTime = $endTime
        $Result.Organizer = $organizer
        $Result.RequiredAttendees = $reqattendeeslist
        $Result.OptionalAttendees = $optattendeeslist
        $Result.TotalAttendees = $attendeecnt
        $Table += $Result
      }  
}
$FileName = "Rooms-" + (Get-Date -f yyyyMMMdd).ToString()
$Table | Export-Csv C:\Temp\$FileName.csv -NoTypeInformation
Write-Host "Your report is saved in C:\Temp\$FileName.csv"