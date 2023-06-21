#This script sample gets all SPO sites and adds a security group to site collection administrators
#It also removes the M365 Group Owners from the site collection administrators

Import-Module Microsoft.Online.Sharepoint.PowerShell -DisableNameChecking
 
#Variables for processing - update the $group variable with the correct group object ID from azure
$AdminURL = "https://<domain>-admin.sharepoint.com/"
$Group = "c:0t.c|tenant|<group object ID>"
 
#Connect to SharePoint Online and authenticate
Connect-SPOService -url $AdminURL 
 
#Get All Site Collections
$Sites = Get-SPOSite -Limit ALL
 
#Loop through each site, add group as site colleciton admin and remove 365 group owners from site collection admins
Foreach ($Site in $Sites)
{
    Write-host "Adding group as site collection admin for:"$Site.URL -f Green
    Set-SPOUser -Site $SiteUrl -LoginName $Group -IsSiteCollectionAdmin $true
    Write-host "Scanning site:"$Site.Url -f Yellow
    #Get All Site Collection Administrators
    $Admins = Get-SPOUser -Site $site.Url | Where-Object {$_.IsSiteAdmin -eq $true}
 
    #Iterate through each admin
    Foreach($Admin in $Admins)
    {
        #Check if the Admin Name contains "Owners"
        If($Admin.DisplayName -match "Owners")
        {
            #Remove Site O365 Groups user from site collection Administrator
            Write-host "Removing Site Collection Admin from:"$Site.URL -f Green
            Set-SPOUser -site $Site -LoginName "c:0o.c|federateddirectoryclaimprovider|$($Admin.LoginName)" -IsSiteCollectionAdmin $False
        }
    }
}


#Read more: https://www.sharepointdiary.com/2017/02/sharepoint-online-remove-site-collection-administrator-using-powershell.html#ixzz85IRqlxCA
#Read more: https://sposcripts.com/add-site-collection-administrator/