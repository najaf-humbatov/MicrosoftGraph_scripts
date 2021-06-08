# MicrosoftGraph_scripts
Collection of scripts used to manage various aspects of Microsoft 365 services through Microsoft Graph API

Disclaimer
The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts are provided AS IS without warranty of any kind. 
Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. 
The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever 
(including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) 
arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.


group_events_cleanup.ps1 - script which could be used to clean-up events from groups' calendars. Script uses device code authentication flow. Please read the overview article for Microsoft's implementation of Oauth 2.0 device code authentication flow
https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code

It was required to add user who ran the script into the each group to work on to avoid the 403 errors. While this is not described in the API documentation https://docs.microsoft.com/en-us/graph/api/group-update-event?view=graph-rest-1.0&tabs=http


