This script queries the /subscribedskus MS Graph endpoint:

https://docs.microsoft.com/en-us/graph/api/subscribedsku-get?view=graph-rest-1.0&tabs=http

The script needs an Azure App Registration with organization.read.all, directory.read.all application or delegated type permissions AND selected.sites permissions. This is because we are going to use the premium HTTP action to do REST API calls to the MS Graph.

To use this script in your environment follow this steps:

1. Create a SharePoint Online list with the following columns:

Title
SkuID (single line of text)
ConsumedUnits (single line of text)
PrepaidUnits (single line of text)
RemainingUnits (single line of text)
Date (Date and time - date only)

2. Replace information in the script accordingly with information for your app registration, SPO site ID in Graph calls, send an email, etc.
