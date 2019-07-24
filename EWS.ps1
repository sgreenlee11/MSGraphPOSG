$clientid = "7e2e61e2-3004-4141-9426-74a5452bcdcb"
$secret = ".ot6v*wSj5U4Nr9-a:oxkdoJ?H/]u2OK"
$resource = "https://outlook.office365.com"
$tenantid = "6c8aa3a9-a08e-4e97-8e75-8f9868f25fa3"
$token = Get-AADToken -ClientID $clientid -ClientSecret $secret -RedirectUri "https://localhost" -TenantID $tenantid -Resource $resource -Scopes "full_access_as_app"

$primaryemail = "admin@M365x894133.onmicrosoft.com"
$smtpuser = "MiriamG@M365x894133.OnMicrosoft.com"
$otheruser = "IsaiahL@M365x894133.OnMicrosoft.com"


#Load EWS Managed API DLL
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
#Create EWS Exchange Object
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2  

$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$service.UseDefaultCredentials = $false
$creds = New-Object System.Net.NetworkCredential($primaryemail,"")   
$service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]::new($token.Result.AccessToken)
#$service.Credentials = $creds      
#Use Autodiscover to find EWS endpoint - Can add static option if required
$service.AutodiscoverUrl($primaryemail, {$true})

$mbx = (Get-mailbox).primarysmtpaddress
foreach($m in $mbx){
$response = [Microsoft.Exchange.WebServices.Autodiscover.GetUserSettingsResponse]::new()
$t = [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverService]::new("M365x894133.OnMicrosoft.com")
$t.url = "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"
$t.Credentials = $creds
$response = $t.GetUserSettings($m,"GroupingInformation")
Write-Host $response.Settings
}

$Service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $smtpuser)
$inboxid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $smtpuser)
$InboxItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
$InboxItems = $service.FindItems($inboxid, $InboxItemView)
$service.HttpHeaders.Add("X-AnchorMailbox", $smtpuser);
$service.HttpHeaders.Add("X-PreferServerAffinity", $true);
$inboxbind =  [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$inboxid)
$newevent = [Microsoft.Exchange.WebServices.Data.EventType]::NewMail
$createdevent = [Microsoft.Exchange.WebServices.Data.EventType]::Created
$deletedevent = [Microsoft.Exchange.WebServices.Data.EventType]::Deleted
$modifiedevent = [Microsoft.Exchange.WebServices.Data.EventType]::Modified
$movedevent = [Microsoft.Exchange.WebServices.Data.EventType]::Moved


$appscopes = New-Object System.Collections.ObjectModel.Collection["Microsoft.Exchange.WebServices.Data.FolderId"]
$appscopes.Add($inboxbind.Id)
$subscription = $service.SubscribeToPullNotifications($appscopes,45,$null,$createdevent,$deletedevent,$modifiedevent,$movedevent)
$events = $subscription.GetEvents()




$Service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $otheruser)
$inboxid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $otheruser)
$InboxItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
$InboxItems = $service.FindItems($inboxid, $InboxItemView)
$inboxbind =  [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$inboxid)
$service.HttpHeaders.Add("X-AnchorMailbox", $smtpuser);
$service.HttpHeaders.Add("X-PreferServerAffinity", $true);

$appscopes = New-Object System.Collections.ObjectModel.Collection["Microsoft.Exchange.WebServices.Data.FolderId"]
$appscopes.Add($inboxbind.Id)


$othersubscription = $service.SubscribeToPullNotifications($appscopes,45,$null,$createdevent,$deletedevent,$modifiedevent,$movedevent)

$events = $othersubscription.GetEvents()

    
