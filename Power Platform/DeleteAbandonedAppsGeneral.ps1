#Script to find and delete abandoned apps in any given Power Platform Environment

#Import modules
Import-Module Microsoft.PowerApps.Administration.PowerShell

#write log function
[string]$LogFileDate = (get-date -f d).replace("/","-")
$LogFile = "C:\Temp\Delete_Abandoned_Apps_$LogFileDate.log"
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

#We will get a list of all apps which have not been modified in the past 120 days
Write-Log "Searching for apps not modified in the past 120 days in the Default Environment"
#$apps = Get-AdminFlow -EnvironmentName $env
$apps = Get-AdminPowerApp -EnvironmentName $env |  Where-Object {[DateTime]$_.LastModifiedTime -lt $Date}
$filteredapps = $apps | Where-Object {$_.internal.properties.embeddedApp.type -Notlike "SharepointFormApp"}

$count = $filteredapps.count
Write-Log "$count have been identified which have not been modified in the past 120 days."

$Abandonedapps = @()
$i = 0

#This section will look at list of app and check if these have owners, co-owners or view permisisons; if no permissions, app will be deleted and added to csv file for record purposes
Write-Log "Looking for abandoned apps with no owners"
foreach ($app in $filteredapps)
{
    $i++
    $hasValidOwner = $false
    $permissions = $null
    #Write-Progress -Activity "Looking for abandoned apps" -Status "Working on $($app.DisplayName)" -PercentComplete (($i / $count) * 100)
    Write-Log "Working on $($app.DisplayName) - #$($i)"

    Try{$permissions = Get-AdminPowerAppRoleAssignment -AppName $app.AppName -environment $env}
    Catch
        {
    
        Write-Log "Error: [$($_.exception.message)]"
        #Reauthenticate
        Write-Log "Session to Power Platform timedout... reconnecting now"
        Connect-PowerApps
        $permissions = Get-AdminPowerAppRoleAssignment -AppName $app.AppName -environment $env
    
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
                        Write-log "$($app.DisplayName) has owners/editors/viewers, skipping..."
                        break
                    }
                }

            }
        }
     }

    if ($hasValidOwner -eq $false)
    {
        
        Write-Log "$($app.DisplayName) will be deleted"
        $Deletedon = get-date -Format "MM/dd/yyyy"

        $Result = "" | Select-Object AppName,DisplayName,CreatedTime,LastModifiedTime,EnvironmentName,CreatedBy,DateDeleted
        $Result.AppName = $app.AppName
        $Result.DisplayName = $app.DisplayName
        $Result.CreatedTime = $app.CreatedTime
        $Result.LastModifiedTime = $app.LastModifiedTime
        $Result.EnvironmentName = $app.EnvironmentName
        $Result.CreatedBy = $app.Owner.email
        $Result.DateDeleted = $Deletedon
        $Abandonedapps += $Result

        #abandoned app is deleted - you can commented this section out if you don't want to delete just yet
        Try{Remove-AdminPowerApp -AppName $app.AppName -EnvironmentName $env}
        Catch
        {
    
        Write-Log "Error: [$($_.exception.message)]"
        #Reauthenticate
        Write-Log "Session to Power Platform timedout... reconnecting now"
        Connect-PowerApps
        Remove-AdminPowerApp -AppName $app.AppName -EnvironmentName $env
    
        }
    }
}

$Countofabandonedapps = $Abandonedapps.count
write-log "$Countofabandonedapps are abandoned and have been deleted"

$Abandonedapps | export-csv C:\Temp\AbandonedApps.csv -Append -NoTypeInformation
Write-Log "AbandonedApps.csv has been updated with list of deleted apps"

Write-Log "**************************************** Script has Finished ****************************************"