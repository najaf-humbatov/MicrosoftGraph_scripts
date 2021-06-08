<#
Sample script to check and report or delete events in unified groups calendars.
v. 0.9

Written by Najaf Humbatov (najaf.humbatov@microsoft.com)
January 4th, 2020

Disclaimer
The sample scripts are not supported under any Microsoft standard support program or service. The sample scripts are provided AS IS without warranty of any kind. 
Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. 
The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever 
(including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) 
arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

Please read the overview article for Microsoft's implementation of Oauth 2.0 device code authentication flow
https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code
#>

<#
This script to operate with delegated user permissions
#>
<#
script to use Power Shell Az module to extract the logged-in user OnjectID to be operated further down the workflow
#>
<#
#Credentials to be used in script run.
$creds=Get-Credential
Connect-AzAccount -Credential $creds
$ObjId=(Get-AzAdUser -UserPrincipalName $creds.UserName).Id
#>
#========Authentication section===========
Write-Host "PowerShell wrapper for AAD Graph reuqests to work with Unified groups calendars"

$TenantId = Read-Host "Please enter the Tenant ID" 
$ClientId = Read-Host "Please enter the Client ID" 
$scope = 'Group.ReadWrite.All GroupMember.ReadWrite.All User.Read'
$resource = 'https://graph.microsoft.com/'

$codeBody = @{
    'resource' = $resource
    'client_id' = $ClientId
    'scope' = $scope
}

# debug output

#Write-Host "CodeBody"
#$codeBody

# Get OAuth Code
$codeRequest = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$TenantId/oauth2/devicecode" -Body $codeBody

# Print Code to console

#Write-Host "CodeRequest"
#$codeRequest | fl
Write-Host "`n$($codeRequest.message)"

$tokenBody = @{

    grant_type = "urn:ietf:params:oauth:grant-type:device_code"
    code       = $codeRequest.device_code
    client_id  = $clientId

}
#Write-Host "Token Body"
#$tokenBody | fl

# Get OAuth Token
while ([string]::IsNullOrEmpty($tokenRequest.access_token)) {

    $tokenRequest = try {

        Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/token" -Body $tokenBody

    }
    catch {

        $errorMessage = $_.ErrorDetails.Message | ConvertFrom-Json

        # If not waiting for auth, throw error
        if ($errorMessage.error -ne "authorization_pending") {
            throw

        }

    }

}

# debug output
Write-Host "Token Request"
$tokenRequest | Format-List

#$tokenRequest.scope
#=======================Main logic section===================================

# Auth header for Invoke-RestMethod

$Headers = @{
    'Authorization' = "Bearer $($tokenRequest.access_token)"
}

# To turn the recurring calendar events into the single occurence event we need to make PATCH request and update recurrence property to $Null and event type to singleInstance

$UpdatedValues = ConvertTo-JSON @{
    "type" = "singleInstance"
    "recurrence" = $null
}
<# 
The script logic would be that script will go through the groups  whih the admin is not a member of and check for recurring events and single them out.
Get the logged in user id to be used to access the private groups' calendars. 
#>

$user_object_id_url='https://graph.microsoft.com/v1.0/me'
$user_object_id = (Invoke-RestMethod -Uri $user_object_id_url -Headers $Headers).Id

Write-Host "running user id - ", $user_object_id
$user_object_id

$user_object_json = ConvertTo-Json @{
    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$user_object_id"
}

Write-Host "User Object representation for POST: ", $user_object_json

<#
Get list of Unified group in the tenant
Microsoft Graph clean request looks like:
https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'Unified')&$select=id,displayName
#>


$group_list_url = 'https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c+eq+'+"`'"+'Unified'+"`'"+')&$select=id'

$group_list = Invoke-RestMethod -Uri $group_list_url -Headers $Headers

# Optional output of groups list

Write-Host "Group List"
$group_list.value.id | Format-List

for ( $i = 0; $i -lt ($group_list.value.id).Count; $i++) {

   # List all the given group members to check if running user is a member
    $group_members_id_list_url = 'https://graph.microsoft.com/v1.0/groups/'+$group_list.value.id[$i]+'/members?$select=id'
    $group_members_id_list = Invoke-RestMethod -Uri $group_members_id_list_url -Headers $Headers
    Write-Host " Groump members list is: " , ($group_members_id_list.value.id).Count
    $group_members_id_list.value.id | Format-List
    
    <#
    Block to check if running user is member of the unified group. if yes we consider this group as home group and
    will not take any actions. If runner wants to check against the home group calendar copy / paste from other route will work.
    #>
    # existance flag to report if user exists in the group

    $if_exists = 0

    for ($k = 0; $k -lt ($group_members_id_list.value.id).Count; $k++) {
        if ($user_object_id -eq $group_members_id_list.value.id[$k]) {
            $if_exists = 1
            Write-Host "Admin group, no action needed", $group_list.value.id[$i], $if_exists
        }
    }
    if ($if_exists -eq 0) {
       Write-Host "Action we need with group id:", $group_list.value.id[$i]
       
       #Adding member to the group
       $group_to_add_member_url = 'https://graph.microsoft.com/v1.0/groups/'+$group_list.value.id[$i]+'/members/$ref'
       $group_to_add_member_url

       #Adding owner to the group
       #$group_to_add_owner_url = 'https://graph.microsoft.com/v1.0/groups/'+$group_list.value.id[$i]+'/owners/$ref'
       #$group_to_add_owner_url

       $Member_to_add = Invoke-WebRequest -Method Post -Uri $group_to_add_member_url -Headers $Headers -ContentType "application/json" -Body $user_object_json -UseBasicParsing
       Write-Host "Memeber to add Status Code: " 
       $Member_to_add.StatusCode

       #$Owner_to_add = Invoke-WebRequest -Method Post -Uri $group_to_add_owner_url -Headers $Headers -ContentType "application/json" -Body $user_object_json -UseBasicParsing
       #Write-Host "Memeber to add Status Code: " 
       #$Owner_to_add.StatusCode
       
       #$Member_to_add | Format-List

       #Write-Host "Group ID : ", $group_list.value.id[$i] # debug output
       #if we add &$filter=type eq 'seriesMaster' at the end of our url we can filter recurring events directly in the body of Graph request
       $events_list_check_uri = 'https://graph.microsoft.com/v1.0/groups/'+$group_list.value.id[$i]+'/events?$select=id,type'

       Write-Host "Resulting request URI is: ", $events_list_check_uri

       $events_info = Invoke-RestMethod -Uri $events_list_check_uri -Headers $Headers

       for ($j=0; $j -lt ($events_info.value).Count; $j++) {
           #Write-Host "Printing id and type values for each of event"
           #$events_info.value[$j].id
           #$events_info.value[$j].type
        
           <# here we form URL for PATCH request for Invoke-RestMethod call. I decided not to check if event is the single instance or recurring one 
           but update properties in bulk for possible performance gain. This check can be easily organized by wrapping construction below into
           if statement similar to the example
           if ( $events_info.value[$j].type -eq 'seriesMaster') { our url construction and PATCH}         
           #>
           $event_edit_uri = 'https://graph.microsoft.com/v1.0/groups/'+$group_list.value.id[$i]+'/events/'+$events_info.value[$j].id
           #$event_test_uri = 'https://graph.microsoft.com/v1.0/groups/'+$group_list.value.id[$i]+'/events/'+$events_info.value[$j].id+'?$select=id'
           Write-Host "Resulting URL for PATCH request", $event_edit_uri # Debug output
           Invoke-RestMethod -Uri $event_edit_uri -Method Patch -Headers $Headers -ContentType "application/json" -Body $UpdatedValues 
           #Invoke-RestMethod -Uri $event_test_uri -Method Get -Headers $Headers

        }

        #Removing added member from the target group after calendar actions.
        $group_to_remove_member_url = 'https://graph.microsoft.com/v1.0/groups/'+$group_list.value.id[$i]+'/members/'+$user_object_id+'/$ref'
        $group_to_remove_member_url
        
        #Removing addedd owner from the target group after calendar actions
        #$group_to_remove_owner_url = 'https://graph.microsoft.com/v1.0/groups/'+$group_list.value.id[$i]+'/owners/'+$user_object_id+'/$ref'
        #$group_to_remove_owner_url
        
        #Owner and member removal
        #$Owner_to_remove = Invoke-WebRequest -Method Delete -Uri $group_to_remove_owner_url -Headers $Headers -UseBasicParsing
        #Write-Host "Owner to remove Status Code: "
        #$Owner_to_remove.StatusCode

        $Member_to_remove = Invoke-WebRequest -Method Delete -Uri $group_to_remove_member_url -Headers $Headers -UseBasicParsing
        Write-Host "Memeber to remove Status Code: "
        $Member_to_remove.StatusCode
        #$Member_to_remove | Format-List
    }

    
    else {
        Write-Host "Everything is quiet in Baghdad city:", $group_list.value.id[$i]
    }
}
    