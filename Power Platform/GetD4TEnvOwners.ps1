#Quick script to get all Dataverse for Teams environments and their ownership which can send an email alert
#to Power Platform Admins so action can be taken before the inactivtity policy soft deletes the DV4T environment

#authenticate to Power Apps/Power Automate
Add-PowerAppsAccount

#Authenticate to EXchange Online
Connect-ExchageOnline

#Get Dataverse for Teams environments
$DV4TS = Get-AdminPowerAppEnvironment | where {$_.EnvironmentType -eq "NotSpecified"} | select-object -property internal

#check to see if the MS Teams behind the Dataverse for Teams environments have owners.
foreach ($DV4T in $DV4TS){

    $MSTeamID = $null
    $MSTeamID = $DV4T.Internal.properties.connectedGroups.ID

    $Owners = Get-UnifiedGroup -Identity $MSTeamID | Get-UnifiedGroupLinks -LinkType Owner

    If($owners -eq $null){

        Write-Output "$MSTeamID Team has no owners"
        #You can use the Send-MailMessage commandlet to send an alert for each Dataverse for Teams which has no owners to a desire admin/DL
        #Send-MailMessage -To "admin@example.onmicrosoft.com" -From noreply@example.onmicrosoft.com -Body "$MSTeamID Team has no owners" -Subject "$MSTeamID Team has no owners" -SmtpServer smtplb.example.onmicrosoft.com

    }
    else
    {
        Write-Output "$MSTeamID Team has owners"
    }

}
