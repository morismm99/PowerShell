#This script uses a new cmmdlet in the Power Platform Admin module to get a list of all flows, across all environments
#at risk of suspension
#More info here - https://learn.microsoft.com/en-us/power-platform/admin/power-automate-licensing/faqs#how-can-i-identify-flows-that-need-premium-licenses-to-avoid-interruptions
#Created on 6/1/2023 by Moris Montejo, Microsoft

#Import modules
Import-Module Microsoft.PowerApps.Administration.PowerShell

# Connect to the Power Platform
Add-PowerAppsAccount

# Get all environments
$environments = Get-PowerAppEnvironment

# Iterate through each environment
foreach ($environment in $environments) {
    $environmentName = $environment.EnvironmentName

    # Run the command for each environment
    Get-AdminFlowAtRiskOfSuspension -EnvironmentName $environmentName -ApiVersion '2016-11-01' | Export-Csv -Path C:\temp\flowsuspensionList.csv -NoTypeInformation -Append
}