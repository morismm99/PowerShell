This script queries the /subscribedskus MS Graph endpoint:

https://docs.microsoft.com/en-us/graph/api/subscribedsku-get?view=graph-rest-1.0&tabs=http

The script needs an Azure App Registration with organization.read.all, directory.read.all application or delegated type permissions AND selected.sites permissions. This is because we are going to use the premium HTTP action to do REST API calls to the MS Graph.

To use this script in your environment follow this steps:

## Create a SharePoint Online list with the following columns:

1. Title
2. SkuID (single line of text)
3. ConsumedUnits (single line of text)
4. PrepaidUnits (single line of text)
5. RemainingUnits (single line of text)
6. Date (Date and time - date only)

## Make sure App Registration has write permissios to the SPO site being used - more info here:

https://ashiqf.com/2021/03/15/how-to-use-microsoft-graph-sharepoint-sites-selected-application-permission-in-a-azure-ad-application-for-more-granular-control/

## Replace information in the script accordingly with information for your app registration, SPO site ID in Graph calls, send an email, etc.