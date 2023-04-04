#How to install MS Graph PS Module - https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0
#How to use MS Graph PS Module - https://learn.microsoft.com/en-us/powershell/microsoftgraph/get-started?view=graph-powershell-1.0
#MS Graph - List apps for users API - https://learn.microsoft.com/en-us/graph/api/userteamwork-list-installedapps?view=graph-rest-1.0&tabs=powershell

# Connect to Microsoft Graph API
Connect-MSGraph -Scopes "User.Read.All","Group.ReadWrite.All","TeamsAppInstallation.ReadForUser"

# Specify the user's email address
$userGUID = "<userGUID>"

# Get the list of apps and their details used by the user
$userApps = Get-MgUserTeamworkInstalledApp -UserId $userGUID -ExpandProperty "teamsAppDefinition"

# Create a table of the apps and their details
$appTable = @()
foreach ($app in $userApps) {
    $appTable += [pscustomobject]@{
        AppName = $app.TeamsAppDefinition.DisplayName
        AppId = $app.TeamsAppDefinition.Id
        PackageId = $app.TeamsAppDefinition.TeamsAppId
    }
}

# Export the app table to a CSV file
$appTable | Export-Csv -Path "C:\temp\user_apps.csv" -NoTypeInformation