# MicrosoftGraph_scripts
Collection of scripts used to manage various aspects of Microsoft 365 services through Microsoft Graph API

group_events_cleanup.ps1 - script which could be used to clean-up events from groups' calendars. Script uses device code authentication flow. Please read the overview article for Microsoft's implementation of Oauth 2.0 device code authentication flow
https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code

It was required to add user who ran the script into the each group to work on to avoid the 403 errors. While this is not described in the API documentation https://docs.microsoft.com/en-us/graph/api/group-update-event?view=graph-rest-1.0&tabs=http


