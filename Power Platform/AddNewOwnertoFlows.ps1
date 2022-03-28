#Quick script to assign ownership/can edit permissions to existing flows
#You can do a Get-AdminFlow | Export-Csv -Path C:\temp\FlowExport2172021.csv to get a list of all flows, then filter to necessary flows
#Then, you can copy those flows you want to add a new owner to into a temp csv file and import it per this script

#authenticate to Power Apps/Power Automate
Add-PowerAppsAccount

#Import list of flows needing a new owner
$flows = Import-CSV -Path C:\Temp\flowstoassign.csv
#Add object ID for new owner to be given permissions
$newowner = "5942757b....."

#bulk add that user as an owner of the imported flows
foreach ($flow in $flows){

Set-AdminFlowOwnerRole -PrincipalType User -PrincipalObjectId $newowner -RoleName CanEdit -FlowName $flow.FlowName -EnvironmentName $flow.EnvironmentName

}