#V1.0
#Written by Moris Montejo
#Microsoft Stream Activities daily audit, which appends a csv file example script

[string]$LogDate = (get-date -f s).replace(":","-")

#Writes output to a log file with a time date stamp
Function Write-Log {
	Param ([string]$string)
	# Get the current date
	[string]$date = Get-Date -Format G
	
    #Log file name

    $LogFile = "C:\Temp\StreamAudit_$Script:LogDate.log"

	# Write everything to our log file
	( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
}

#Import new Exchange Online V2 Module
Import-Module ExchangeOnlineManagement

Function Connect-EXOPSSession {

	Write-Log "Importing new EXO Powershell Session"
	$O365ExchangeAdminPW = "encryptedpassword"
	$O365ExchangeAdminSecure = $O365ExchangeAdminPW | ConvertTo-SecureString
	$creds = New-Object System.Management.Automation.PSCredential ("O365ExchangeAdmin@contoso.net", $O365ExchangeAdminSecure)
	Connect-ExchangeOnline -Credential $creds -ShowProgress $true

}

#Connect to Exchange Online V2
Write-Log "Connecting to Exchange Online"
Connect-EXOPSSession

$Yesterday = (get-date).AddDays(-1) | get-date -Format d
$Today = get-date -Format d

#Path to Stream audit report file
$StreamAuditCSVPath = "c:\Temp\StreamAudit.csv"

#Set variable value to 1 to trigger loop below
$AuditData = 1

Write-Log "Searching the audit log for Stream Activities"
while ($AuditData){

    #$AuditData = Search-UnifiedAuditLog -startdate $Yesterday -EndDate $Today -RecordType MicrosoftStream -ResultSize 5000 | select -ExpandProperty AuditData | ConvertFrom-Json
    $AuditData = Search-UnifiedAuditLog -startdate $Yesterday -EndDate $Today -RecordType MicrosoftStream -SessionID "1" -SessionCommand ReturnLargeSet | Select-Object -ExpandProperty AuditData | ConvertFrom-Json

    $AuditDataFiltered = $AuditData | Select-Object ResourceTitle,ResourceURL,UserID,CreationTime,Operation | Where-Object {$_.Operation -eq "StreamCreateVideo" -or $_.Operation -eq "StreamInvokeVideoView" -or $_.Operation -eq "StreamInvokeVideoUpload" -or $_.Operation -eq "StreamDeleteVideo"}

    $AuditDataFiltered | export-csv -path $StreamAuditCSVPath -NoTypeInformation -Append

    }

Write-Log "The raw appended report is saved in C:\Temp\StreamAudit.csv"

Write-Log "Checking for and deleting duplicate rows"
$StreamAuditCSV = Import-Csv $StreamAuditCSVPath | Sort-Object * -Unique

$StreamAuditCSV | Export-Csv $StreamAuditCSVPath -NoTypeInformation
Write-log "Report has been checked for duplicate rows and cleansed, it is saved in C:\temp\StreamAudit.csv"

#Connect to SharePoint to copy reports using PnP or you could use other methods to write file to SharePoint - 
#This exampled used an older versoin of PnP which could use credentials stored in credential manager to authenticate to SharePoint
Write-Log "Connecting to <SPO site> using PnP to copy csv files from local to SPO"

Connect-PnPOnline -Url <https://contoso.sharepoint.com/sites/Reports>

Add-PnPFile -Path $StreamAuditCSVPath -Folder "Shared Documents\Platform\StreamAuditReport"
Write-Log "StreamAudit has been copied"

Write-Log "Script has ended"

Exit