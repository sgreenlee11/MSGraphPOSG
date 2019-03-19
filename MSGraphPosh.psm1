function Get-AADToken
{
	param(
		[String[]]$Scopes,
		$ClientID,
		$AppName,
		$UserName,
		$Credential,
		$ClientSecret,
		$RedirectUri,
		$TenantID,
		$Resource
	)

	#Use Organizations endpoint when sending credentials as part of request, otherwise common
	If($username -or $Credential)
	{
		[uri]$authority = "https://login.microsoftonline.com/organizations/oauth2/authorize"
	}
	ElseIf($ClientSecret)
	{
		[uri]$authority = "https://login.microsoftonline.com/" + $TenantID

	}
	else
	{
		[uri]$authority = "https://login.microsoftonline.com/common/oauth2/authorize"
	}
	if($ClientSecret)
	{
		
		$clientcred = [Microsoft.Identity.Client.ClientCredential]::new($ClientSecret)
		$clientapp = [Microsoft.Identity.Client.ConfidentialClientApplication]::new($clientid,$authority,$RedirectUri,$clientcred,$null,$null)
		Write-Warning "Changing scope to https://graph.microsoft.com/.default for Client Credentials Flow. Will not work for other APIs"
		if($Resource)
		{
			$scopes = $resource + "/.default"
		}
		else{
			Write-Warning "Changing scope to https://graph.microsoft.com/.default for Client Credentials Flow. Will not work for other APIs. Use Resource Parameter if not using Graph"
			$scopes = "https://graph.microsoft.com/.default"
		}

	}
	else{
		$clientapp = [Microsoft.Identity.Client.PublicClientApplication]::new($clientid,$authority)
	}
	#Build scopes array
	$appscopes = New-Object System.Collections.ObjectModel.Collection["string"]
	foreach($s in $scopes)
	{
		$appscopes.Add($s)
	}
	#If Username parameter is specified, use WIA authentication with logged on account
	If($username)
	{
		$authresult = $clientapp.AcquireTokenByIntegratedWindowsAuthAsync($appscopes,$username)
	}
	elseif ($Credential) {
		$authresult = $clientapp.AcquireTokenByUsernamePasswordAsync($appscopes,$Credential.Username,(ConvertTo-SecureString ($Credential.GetNetworkCredential().Password) -AsPlainText -Force))
	}
	elseif($ClientSecret)
	{
		$authresult = $clientapp.AcquireTokenForClientAsync($appscopes)
	}
	else{
		$authresult = $clientapp.AcquireTokenAsync($appscopes)
	}
	
	return $authresult

	
}

function Get-RESTMailContacts
{
    param(
        $Authentication,
        $Resource,
        $Username
    )
    
    #Create array for output
    $contactarray = New-Object System.Collections.ArrayList

    #Grab first 10 contacts

    if($username)
    {

    }
    else
    {
    $contacturl = "https://graph.microsoft.com/v1.0/me/contacts"
    }
    $allcontacts = Invoke-RestMethod -Uri $contacturl -Method GET -Headers @{Authorization = $Authentication.result.CreateAuthorizationHeader()}

    #Loop through returned contacts and 
    foreach($c in $allcontacts.value)
    {
        $conobj = New-Object psobject
        $conobj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $c.DisplayName
        $conobj | Add-Member -MemberType NoteProperty -Name "Title" -Value $c.jobTitle
        $conobj | Add-Member -MemberType NoteProperty -Name "Manager" -Value $c.manager
        $conobj | Add-Member -MemberType NoteProperty -Name "Address" -Value ($c.emailaddresses | Select-Object address).address
        [void]$contactarray.Add($conobj)
    }

    #Check for more than 10 contacts
    if($allcontacts.'@odata.nextLink' -ne $null)
    {
        $morecontacts = $true
        $nextlink = $allcontacts.'@odata.nextLink'
    }
    while($morecontacts -eq $true)
    {
        $nextcontacts = Invoke-RestMethod -Uri $nextlink -Method GET -Headers @{Authorization = $Authentication.result.CreateAuthorizationHeader()}

        #Loop through returned contacts and 
        foreach($c in $nextcontacts.value)
        {
            $conobj = New-Object psobject
            $conobj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $c.DisplayName
            $conobj | Add-Member -MemberType NoteProperty -Name "Title" -Value $c.jobTitle
            $conobj | Add-Member -MemberType NoteProperty -Name "Manager" -Value $c.manager
            $conobj | Add-Member -MemberType NoteProperty -Name "Address" -Value ($c.emailaddresses | Select-Object address).address
            [void]$contactarray.Add($conobj)
        }
        if($nextcontacts.'@odata.nextLink' -eq $null)
        {
            $morecontacts = $false
        }
        else {
            $nextlink = $nextcontacts.'@odata.nextLink' 
        }
    }


    return $contactarray
}

function Add-RESTGuest
{
    param(
        $Authentication,
        $GuestEmail,
        $RedirectUrl,
        $SendInviteMessage,
        $GuestDisplayName
    )
    $inviteurl = $groupurl = "https://graph.microsoft.com/v1.0/invitations"

    $inviteurl = $groupurl = "https://graph.microsoft.com/v1.0/invitations"
    $newinvitebody = @{
    "invitedUserEmailAddress" = $GuestEmail;
    "inviteRedirectUrl"=$GuestEmail;
    "sendInvitationMessage"=$SendInviteMessage;
    "invitedUserDisplayName"=$GuestDisplayName
    }
$newinvitebody = $newinvitebody | ConvertTo-Json -Depth 10
Invoke-RestMethod -Uri $inviteurl -Method POST -Headers @{Authorization = $authentication.result.CreateAuthorizationHeader()} -ContentType application/json -Body $newinvitebody

}