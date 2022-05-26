#Script to find and delete abandoned flows in any given Power Platform Environment
#Updated 5/26/2022

#$LogFileName = "C:\Scripts\ScheduledTasks\AbandonedFlowsApps\Logs\testflows.log"

#Import modules
Import-Module Microsoft.PowerApps.Administration.PowerShell

#write log function
[string]$LogFileDate = (get-date -f d).replace("/","-")
$LogFile = "C:\Temp\Delete_Abandoned_Flows_$LogFileDate.log"
Function Write-Log {
	Param ([string]$string)
	# Get the current date
	[string]$date = Get-Date -Format G
	
	# Write everything to our log file
	( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
}

#You can encrypt the passwords using these instructions https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/convertto-securestring?view=powershell-7.2#example-1--convert-a-secure-string-to-an-encrypted-string
#Function to connect to PowerApps - need to give identity Power Platform Admin Role
Function Connect-PowerApps {

    $PowerPlatformAdminPW = Get-Content C:\Scripts\EncryptedSecrets\PowerPlatformAdmin_Pass.txt
    $PowerPlatformAdminSecure = $PowerPlatformAdminPW | convertto-secureSTring
    $creds = New-Object System.Management.Automation.PSCredential ("Username@contoso.net", $PowerPlatformAdminSecure)
	Add-PowerAppsAccount -username "username@contoso.net" -Password $creds.Password
}

#Function to connecto to AAD 
Function Connect-AAD {
	$PowerPlatformAdminPW = Get-Content C:\Scripts\EncryptedSecrets\PowerPlatformAdmin_Pass.txt
	$PowerPlatformAdminSecure = $PowerPlatformAdminPW | convertto-secureSTring
	$creds = New-Object System.Management.Automation.PSCredential ("Username@contoso.com", $PowerPlatformAdminSecure)
	Connect-AzureRmAccount -TenantId 'gettenantidfromazure' -Credential $creds
	Get-AzureRmADUser -UserPrincipalName Username@contoso.com
	$ctx = Get-AzureRmContext
	$cache = $ctx.TokenCache
	$cacheItems = $cache.ReadItems()
	$token = ($cacheItems | where { $_.Resource -eq "https://graph.windows.net/" })
	Connect-AzureAD -AadAccessToken $token.AccessToken -AccountId $ctx.Account.Id -TenantId $ctx.Tenant.Id
}

Write-Log "**************************************** Script has started ****************************************"

Connect-PowerApps
Connect-AAD

#set days variable to use to subtract 120 days from todays date
[int64]$Days = "-120"

#today's date minus 120 days
[DateTime]$Date = Get-Date ((Get-Date).AddDays($Days)).touniversaltime() -Format o

#We are going to only look for orphaned flows in the default Power Platform environment or provide ID for another environment
$env = "DefaultEnvironmentID"

#We will get a list of all flows which have not been modified in the past 120 days
Write-Log "Searching flows not modified in the past 120 days in the Cerner Default Environment"
#$flows = Get-AdminFlow -EnvironmentName $env
$flows = Get-AdminFlow -EnvironmentName $env |  Where-Object {[DateTime]$_.LastModifiedTime -lt $Date}
$filteredflows = $flows | Where-Object {$_.DisplayName -Notlike "Request sign-off"}
$count = $filteredflows.count
Write-Log "$count have been identified which have not been modified in the past 120 days and are not Request Sign-off flows"

$Abandonedflows = @()
$i = 0

#This section will look at list of flows and check if these have owners or co-owners; if no owners flow will be deleted and added to csv file for record purposes
Write-Log "Looking for abandoned flows with no owners"
foreach ($flow in $filteredflows)
{
    $i++
    $hasValidOwner = $false
    $permissions = $null
    #Write-Progress -Activity "Looking for abandoned flows" -Status "Working on $($flow.FlowName)" -PercentComplete (($i / $count) * 100)
    Write-Log "Working on $($flow.FlowName) - #$($i)"
    
    Try{$permissions = Get-AdminFlowOwnerRole -EnvironmentName $env -FlowName $flow.FlowName}
    Catch
        {
    
        Write-Log "Error: [$($_.exception.message)]"
        #Reauthenticate
        Write-Log "Session to Power Platform timedout... reconnecting now"
        Connect-PowerApps
        $permissions = Get-AdminFlowOwnerRole -EnvironmentName $env -FlowName $flow.FlowName
    
        }

    If ($Permissions -ne $null){

        foreach ($permission in $permissions) 
        {
            $users = $null
            $roleType = $permission.RoleType
        
            if ($roleType -ne $null){

                if ($roleType.ToString() -eq "Owner" -or $roleType.ToString() -eq "CanEdit" -or $roleType.ToString() -eq "CanView")
                {
                    $userId = $permission.PrincipalObjectId
                    Try{$users = Get-AzureADUser -Filter "ObjectId eq '$userId'"}
                    Catch [Microsoft.Open.AzureAD16.Client.ApiException]
                        {
                 
                         Write-Log "Error: [$($_.exception.message)]"
                         #Reauthenticate
                         Write-Log "Session to AzureAD timedout... reconnecting now"
                         Connect-AAD
                         $users = Get-AzureADUser -Filter "ObjectId eq '$userId'"

                        }

                    if ($users.Length -gt 0)
                    {
                        $hasValidOwner = $true
                        Write-log "$($flow.FlowName) has owners/editors/viewers, skipping..."
                        break
                    }
                }

            }
        }
     }

    if ($hasValidOwner -eq $false)
    {
        #$flow | select FlowName,Enabled,DisplayName,CreatedTime,LastModifiedTime,EnvironmentName,@{Name="DateDeleted"; Expression={get-date -Format MM/dd/yyyy}} | export-csv c:\temp\abandonedflowstest.csv -Append -NoTypeInformation
        Write-Log "$($flow.FlowName) will be deleted"
        $Deletedon = get-date -Format "MM/dd/yyyy"

        $Result = "" | Select-Object FlowName,Enabled,DisplayName,CreatedTime,LastModifiedTime,EnvironmentName,CreatedBy,DateDeleted
        $Result.FlowName = $flow.FlowName
        $Result.Enabled = $flow.Enabled
        $Result.DisplayName = $flow.DisplayName
        $Result.CreatedTime = $flow.CreatedTime
        $Result.LastModifiedTime = $flow.LastModifiedTime
        $Result.EnvironmentName = $flow.EnvironmentName
        $Result.CreatedBy = $flow.CreatedBy.ObjectID
        $Result.DateDeleted = $Deletedon
        $Abandonedflows += $Result

        #abandoned flow is disabled
        #Try{Disable-AdminFlow -EnvironmentName $env -FlowName $flow.FlowName}
        #Catch
        #{
    
        #Write-Log "Error: [$($_.exception.message)]"
        #Reauthenticate
        #Write-Log "Session to Power Platform timedout... reconnecting now"
        #Connect-PowerApps
        #Disable-AdminFlow -EnvironmentName $env -FlowName $flow.FlowName
    
        #}

        #abandoned flow is deleted - feel free to comment this section out if you are not ready to delete just yet
        Try{Remove-AdminFlow -EnvironmentName $env -FlowName $flow.FlowName}
        Catch
        {
    
        Write-Log "Error: [$($_.exception.message)]"
        #Reauthenticate
        Write-Log "Session to Power Platform timedout... reconnecting now"
        Connect-PowerApps
        Remove-AdminFlow -EnvironmentName $env -FlowName $flow.FlowName
    
        }
    }
}

$Countofabandonedflows = $Abandonedflows.count
write-log "$Countofabandonedflows are abandoned and have been deleted"

$Abandonedflows | export-csv C:\Temp\abandonedflows.csv -Append -NoTypeInformation
Write-Log "Abandonedflows.csv has been updated with list of deleted flows"

Write-Log "**************************************** Script has Finished ****************************************"