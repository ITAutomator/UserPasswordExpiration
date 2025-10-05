## To enable scrips, Run powershell as admin then type Set-ExecutionPolicy RemoteSigned
#region    --- Transcript Open
$TranscriptTemp = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $TranscriptTemp | Out-Null
#endregion --- Transcript Open
#region    --- Main function header
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
#$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
#$scriptCSV      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".csv"  ### replace .ps1 with .csv
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
$psm1="$($scriptDir)\ITAutomator.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {Write-Host "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
#$psm1="$($scriptDir)\ITAutomator M365.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {Write-Host "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
#endregion  --- Main function header
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Use this script to: "
Write-Host "[O] Org settings: View or change the organization's password expiration settings"
Write-Host "     Note: There is no UI in Entra to view or set these options."
Write-Host ""
Write-Host "[R] Report user password expiration: Report on all users' password expiration status to a CSV file"
Write-Host "    Note: This report can be useful before changing the org settings."
Write-Host ""
Write-Host "[U] Update user password expiration: Change individual users' password expiration settings from a CSV file"
Write-Host "    Note: This sets or clears the DisablePasswordExpiration policy"
Write-Host ""
Write-Host "Note: If a user is about to hit the password change date"
Write-Host "and you change their policy to DisablePasswordExpiration=TRUE, "
Write-host "they will NOT be prompted to change their password at next login."
Write-Host ""
Write-Host "-----------------------------------------------------------------------------"
#region Connections
if ($true) { # Connect-MgGraph
    if (-not (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue)) {
        Write-Host "Connect-MgGraph is NOT available. You may need to install the Microsoft Graph module:"
        Write-Host 'Install-Module Microsoft.Graph -Scope CurrentUser'
        PressEnterToContinue
        exit
    }
    # Check if we are already connected
    while ($true) {
        # Check if already connected to Microsoft Graph
        $mgContext = Get-MgContext
        if ($mgContext -and $mgContext.Account -and $mgContext.TenantId) {
            Write-Host "Already connected to Microsoft Graph."
            Write-Host " Connected as: $($mgContext.Account)"
            Write-Host "    Tenant ID: $($mgContext.TenantId)"
            $tenantDomain = (Get-MgDomain | Where-Object { $_.IsDefault }).Id
            Write-Host "Tenant Domain: " -NoNewline
            Write-Host $tenantDomain -ForegroundColor Green
            $response = AskForChoice "Choice: " -Choices @("&Use this connection","&Disconnect and try again","E&xit") -ReturnString
            # If the user types 'exit', break out of the loop
            if ($response -eq 'Disconnect and try again') {
                Write-Host "Disconnect-MgGraph..."
                Disconnect-MgGraph | Out-Null
                PressEnterToContinue "Done. Press <Enter> to connect again."
                Continue # loop again
            }
            elseif ($response -eq 'exit') {
                return
            }
            else { # on to next step
                break
            }
        } else {
            Write-Host "Not connected. Connecting now..."
            Write-Host "We will try 'Connect-MgGraph' to authenticate. Before we do, open a browser to an admin session on the desired tenant."
            PressEnterToContinue
            Connect-MgGraph -Scopes "User.ReadWrite.All", "Mail.ReadWrite", "Directory.ReadWrite.All" -NoWelcome
            # Confirm connection
            $mgContext = Get-MgContext
            if ($mgContext) {
                Write-Host "Now connected to Microsoft Graph as $($mgContext.Account)"
                #Write-Host "Tenant Domain: $($mgContext.TenantDomain)"
                # Make sure you're connected with Directory.Read.All or Directory.ReadWrite.All
                $tenantDomain = (Get-MgDomain | Where-Object { $_.IsDefault }).Id
                Write-Host "Tenant Domain: $tenantDomain"
            } else {
                Write-Error "Failed to connect to Microsoft Graph."
            }
        }
    } # while true forever loop
    Write-Host
} # Connect-MgGraph
if ($false) { # Connect-ExchangeOnline
    # Check if Connect-ExchangeOnline is available
    if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        Write-Host "ERROR: 'Connect-ExchangeOnline' command was not found."
        Write-Host "Please install the ExchangeOnlineManagement module using:"
        Write-Host "   Install-Module ExchangeOnlineManagement"
        Write-Host "Or load the module if it is already installed, then try again."
        Write-Host "Press any key to exit..."
        PressEnterToContinue
        exit
    }
    # Check if we are already connected
    while ($true) {
        try {
            $orgConfig = Get-OrganizationConfig -ErrorAction Stop
            # The Identity property typically shows your tenant's name or domain
            $tenantNameOrDomain = $orgConfig.Identity
            Write-Host "You are currently connected to tenant: " -NoNewline
            Write-host $tenantNameOrDomain -ForegroundColor Green
            $response = AskForChoice "Choice: " -Choices @("&Use this connection","&Disconnect and try again","E&xit") -ReturnString
            # If the user types 'exit', break out of the loop
            if ($response -eq 'Disconnect and try again') {
                Write-Host "Disconnect-ExchangeOnline..."
                $null = Disconnect-ExchangeOnline -Confirm:$false
                PressEnterToContinue "Done. Press <Enter> to connect again."
                Continue # loop again
            }
            elseif ($response -eq 'exit') {
                return
            }
            else { # on to next step
                break
            }
        } # try steps
        catch {
            Write-Host "ERROR: Not connected to Exchange Online or invalid session."
            Write-Host "We will try 'Connect-ExchangeOnline' to authenticate. Before we do, open a browser to an admin session on the desired tenant."
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            Write-Host "Connect-ExchangeOnline ... " -ForegroundColor Yellow
            Connect-ExchangeOnline -ShowBanner:$false
            Write-Host "Done" -ForegroundColor Yellow
            Continue # loop again
        } # catch error
    } # while true forever loop
    Write-Host
} # Connect-ExchangeOnline
#endregion Connections
do
{
    $choice = AskForChoice "Choices" -Choices @("&Org settings","&Report user password expiration","&Update user password expiration","E&xit") -Default 3 -ReturnString -ShowMenu
    if ($choice -eq "Exit") { Continue }
    if ($choice -eq "Org settings") { # Org settings
        #Connect-MgGraph -Scopes "Domain.Read.All"
        do {
            Write-Host "------------------- Organization Domain-level Password Expiration Settings ------------------"
            $mgdomains = Get-MgDomain | Select-Object Id, PasswordValidityPeriodInDays, PasswordNotificationWindowInDays
            $domchoice = ChooseFromList $mgdomains -title "Domain to change. Note: 2147483647 means NEVER (and user-level settings are ignored)" 
            if ($domchoice -ne -1) {
                Write-Host "                         Domain: " -NoNewline
                Write-Host $mgdomains[$domchoice].Id -ForegroundColor Green
                if (($null -eq $mgdomains[$domchoice].PasswordValidityPeriodInDays) -or ($mgdomains[$domchoice].PasswordValidityPeriodInDays -eq 2147483647))
                { # blank or max
                    Write-Host "The domain doesn't enforce password expiration." -ForegroundColor Green
                    Write-Host "(Setting DisablePasswordExpiration=TRUE on a user is ignored)"
                } # blank or max
                else { # has a non-max number
                    Write-Host "Expires passwords every n days: " -NoNewline
                    Write-Host $mgdomains[$domchoice].PasswordValidityPeriodInDays -ForegroundColor Green
                    Write-Host "   Notification window in days: " -NoNewline
                    Write-Host $mgdomains[$domchoice].PasswordNotificationWindowInDays -ForegroundColor Green
                    Write-Host "(Setting DisablePasswordExpiration=TRUE on a user will override the domain.)"
                } # has a non-max number
                #region Change PasswordValidityPeriodInDays
                $newdays = Read-Host "Enter new number of days to expire passwords (use 0 for NEVER) (or press <Enter> to skip)"
                if ($newdays -match '^\d+$') {
                    Write-Host "Current  PasswordValidityPeriodInDays is $($mgdomains[$domchoice].PasswordValidityPeriodInDays)" -ForegroundColor Yellow
                    if ($newdays -eq 0) {$newdays = 2147483647} # convert 0 to max
                    if ($null -eq $mgdomains[$domchoice].PasswordValidityPeriodInDays) {$mgdomains[$domchoice].PasswordValidityPeriodInDays = 2147483647} # convert 0 to max
                    if ($newdays -eq $mgdomains[$domchoice].PasswordValidityPeriodInDays) {
                        Write-Host "No change needed. It is already set to $newdays days." -ForegroundColor Green
                    } else {
                        Write-Host "Changing PasswordValidityPeriodInDays to $newdays ..." -NoNewline
                        try {
                            Connect-MgGraph -Scopes "Domain.ReadWrite.All","Directory.AccessAsUser.All" -NoWelcome
                            Update-MgDomain -DomainId $mgdomains[$domchoice].Id -PasswordValidityPeriodInDays $newdays
                            Write-Host "Updated." -ForegroundColor Green
                        } catch {
                            Write-Host "Error updating domain: $_" -ForegroundColor Red
                        }
                    }
                } else {
                    Write-Host "Skipped changing PasswordValidityPeriodInDays."
                }
                #endregion PasswordValidityPeriodInDays
                #region Change PasswordNotificationWindowInDays
                $newdays = Read-Host "Enter new number of days for the notification window (14 for 2 weeks) (or press <Enter> to skip)"
                if ($newdays -match '^\d+$') {
                    if ($newdays -eq $mgdomains[$domchoice].PasswordNotificationWindowInDays) {
                        Write-Host "No change needed. It is already set to $newdays days." -ForegroundColor Green
                    } else {
                        Write-Host "Current  PasswordNotificationWindowInDays is $($mgdomains[$domchoice].PasswordNotificationWindowInDays)" -ForegroundColor Yellow
                        Write-Host "Changing PasswordNotificationWindowInDays to $newdays ..." -NoNewline
                        try {
                            Connect-MgGraph -Scopes "Domain.ReadWrite.All","Directory.AccessAsUser.All" -NoWelcome
                            Update-MgDomain -DomainId $mgdomains[$domchoice].Id -PasswordNotificationWindowInDays $newdays
                            Write-Host "Updated." -ForegroundColor Green
                        } catch {
                            Write-Host "Error updating domain: $_" -ForegroundColor Red
                        }
                    }
                } else {
                    Write-Host "Skipped PasswordNotificationWindowInDays."
                }
                #endregion Change PasswordNotificationWindowInDays
            }
        } until ($domchoice -eq -1) 
    } # Org settings
    if ($choice -eq "Report user password expiration") { # Report user password expiration
        $rows=@()
        Write-Host "Get-MgDomain..."
        $mgdomains = Get-MgDomain | Select-Object Id, PasswordValidityPeriodInDays, PasswordNotificationWindowInDays
        Write-Host "Get-MgUser... " -noNewline
        $mgUsers = Get-MgUser -All -Property "displayName,userPrincipalName,passwordPolicies,AccountEnabled,LastPasswordChangeDateTime,CreatedDateTime,Mail" | Select-Object displayName,userPrincipalName,passwordPolicies,AccountEnabled,LastPasswordChangeDateTime,CreatedDateTime,Mail | Sort-Object displayName
        Write-Host "$($mgUsers.count) users"
        $i=0
        ForEach ($mgUser in $mgUsers)
        { # each user
            $Warnings = @()
            Write-Host "$((++$i)) of $($mgUsers.Count): $($mgUser.DisplayName) <$($mgUser.UserPrincipalName)>"
            $PasswordNeverExpires= $mguser.PasswordPolicies -match "DisablePasswordExpiration"
            $userdomain = $mgUser.UserPrincipalName.Split("@")[1]
            $dom_info = $mgdomains | Where-Object { $_.Id -ieq $userdomain }
            $dom_PasswordValidityPeriodInDays     = if ($dom_info.PasswordValidityPeriodInDays -eq 2147483647) { 0 } else { $dom_info.PasswordValidityPeriodInDays} 
            $dom_PasswordNotificationWindowInDays = $dom_info.PasswordNotificationWindowInDays
            if ($mguser.AccountEnabled)
            { # Account is enabled
                $PasswordExpiryDisabled       = $mguser.PasswordPolicies -match "DisablePasswordExpiration"
                if ($PasswordExpiryDisabled)
                { # user has DisablePasswordExpiration
                    $PasswordExpiryInDays = "Never (User has DisablePasswordExpiration policy)"
                    $PasswordExpiryDate   = "Never (User has DisablePasswordExpiration policy)"
                } # user has DisablePasswordExpiration
                else
                { # user has not DisablePasswordExpiration
                    if ($dom_PasswordValidityPeriodInDays -gt 0)
                    { # domain has a password expiry policy
                        if ($mguser.LastPasswordChangeDateTime)
                        { # user has changed password at least once
                            $PasswordExpiryInDays = [math]::Round((($mguser.LastPasswordChangeDateTime.AddDays($dom_PasswordValidityPeriodInDays)) - (Get-Date)).TotalDays,0)
                            # if ($PasswordExpiryInDays -lt 0) {$PasswordExpiryInDays = 0}
                            $PasswordExpiryDate = $mguser.LastPasswordChangeDateTime.AddDays($dom_PasswordValidityPeriodInDays)
                            if ($PasswordExpiryInDays -lt $dom_PasswordNotificationWindowInDays)
                            {
                                if ($PasswordExpiryInDays -lt 0) {
                                    $Warnings+="Password expired"
                                }
                                else {
                                    $Warnings+="Password expiry notification active [$($dom_PasswordNotificationWindowInDays) days]"
                                }
                            }
                        } # user has changed password at least once
                        else
                        { # user has never changed password
                            # use created date as a proxy for last pwd change date
                            if ($mguser.CreatedDateTime)
                            { # has created date
                                $PasswordExpiryInDays = [math]::Round((($mguser.CreatedDateTime.AddDays($dom_PasswordValidityPeriodInDays)) - (Get-Date)).TotalDays,0)
                                # if ($PasswordExpiryInDays -lt 0) {$PasswordExpiryInDays = 0}
                                $PasswordExpiryDate = $mguser.CreatedDateTime.AddDays($dom_PasswordValidityPeriodInDays)
                                if ($PasswordExpiryInDays -lt $dom_PasswordNotificationWindowInDays)
                                {
                                    if ($PasswordExpiryInDays -lt 0) {
                                        $Warnings+="Password expired"
                                    }
                                    else {
                                        $Warnings+="Password expiry notification active [$($dom_PasswordNotificationWindowInDays) days]"
                                    }
                                }
                            } # has created date
                            else
                            { # no created date - can't determine expiry date
                                $Warnings+="User has never changed their password and there's no CreatedDate to use as a proxy for pwd change date"
                                $PasswordExpiryInDays     = "Unknown"
                                $PasswordExpiryDate       = "Unknown"
                            } # no created date - can't determine expiry date
                        } # user has never changed password
                    } # domain has a password expiry policy
                    else
                    { # domain has no password expiry policy
                        $PasswordExpiryInDays     = "Never (Domain has no password expiry policy)"
                        $PasswordExpiryDate       = "Never (Domain has no password expiry policy)"
                    } # domain has no password expiry policy
                } # user has not DisablePasswordExpiration
            } # Account is enabled
            else 
            { # Account is disabled
                $PasswordExpiryInDays = "N/A (Account is disabled)"
                $PasswordExpiryDate   = "N/A (Account is disabled)"
            } # Account is disabled
            $row_obj=[pscustomobject][ordered]@{
                UserPrincipalName             = $mguser.UserPrincipalName
                Mail                          = $mguser.Mail
                DisplayName                   = $mguser.DisplayName
                CreatedDate                   = $mguser.CreatedDateTime
                LastPasswordChangeDate        = $mguser.LastPasswordChangeDateTime
                PasswordExpiryDate            = if ($PasswordExpiryDate -is [datetime]) {$PasswordExpiryDate} else {$PasswordExpiryDate}
                PasswordExpiryInDays          = $PasswordExpiryInDays
                AccountEnabled                = $mguser.AccountEnabled
                PasswordNeverExpires          = $PasswordNeverExpires
                Domain                        = $userdomain
                DomainPasswordExpiryInDays    = if ($dom_PasswordValidityPeriodInDays -eq 0) {"Never"} else {$dom_PasswordValidityPeriodInDays}
                DomainPasswordNotificationInDays = $dom_PasswordNotificationWindowInDays
                Warnings                      = $Warnings -join "; "
            }
            #### Add to results
            $rows += $row_obj
        } # each user
        # Export results
        $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $ReportTarget = "$(Split-Path $scriptFullname -Parent)\Reports\$($scriptBase)_report_$($date).csv"
        New-Item (Split-Path $ReportTarget -Parent) -ItemType Directory -Force | Out-Null # Make folder
        $rows | Export-Csv -Path $ReportTarget -NoTypeInformation -Encoding UTF8
        Write-Host "Report exported to: " -NoNewline
        Write-Host (Split-Path $ReportTarget -Leaf) -ForegroundColor Green
        if (AskForChoice "Open the report now?") {
            Invoke-Item $ReportTarget
        }
    } # Report user password expiration
    if ($choice -eq "Update user password expiration") { # Change user password expiration
        #region --- read / create CSV file of users
        $UpdateFolder = "$(Split-Path $scriptFullname -Parent)\Updates"
        New-Item $UpdateFolder -ItemType Directory -Force | Out-Null # Make folder
        # Search updates folder for most recent CSV file
        $UpdateFile = Get-ChildItem -Path $UpdateFolder -Filter "*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($UpdateFile) {
            $UpdateCSV = $UpdateFile.FullName
            Write-Host "Using most recent CSV file found in 'Updates' folder: " -NoNewline
            Write-Host (Split-Path $UpdateCSV -Leaf) -ForegroundColor Green
        } else {
            $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
            $UpdateCSV = "$($UpdateFolder)\$($scriptBase)_update_$($date).csv"
            Write-Host "No CSV files found in 'Updates' folder. Will use default file: " -NoNewline
            Write-Host (Split-Path $UpdateCSV -Leaf) -ForegroundColor Green
            ######### Template
            "UserPrincipalName,DisablePasswordExpiration" | Add-Content $UpdateCSV
            "jsmith@domain.com,TRUE" | Add-Content $UpdateCSV
            ######### Template
            Write-Host "Template CSV created." -ForegroundColor Yellow
            PressEnterToContinue "Press <Enter> to open the CSV file for editing."
            Invoke-Item $UpdateCSV
            PressEnterToContinue "When done editing the CSV file, press <Enter> to continue."
        }
        ## ----------Fill $rows with contents of file
        $rows=@(import-csv $UpdateCSV)
        $rowscount = $rows.count
        Write-Host "CSV: $(Split-Path $UpdateCSV -leaf) ($($rowscount) entries)"
        $rows | Format-Table
        #endregion --- read / create CSV file of users
        #region Get domain from first entry (optional)
        $domain_col=@($rows[0].psobject.Properties.name)[0] # property name of first column in csv
        $domain=$rows[0].$domain_col # contents of property
        $domain=$domain.Split("@")[1]   # domain part
        Write-Host "Connected to Microsoft Graph domain: " -NoNewline
        Write-Host $tenantDomain -ForegroundColor Green
        Write-Host "    Default domain from first entry: " -NoNewline
        Write-Host $domain -ForegroundColor Green
        if ($domain -ne $tenantDomain)
        {
            Write-Host "WARNING: The domain from the first entry doesn't match the connected tenant domain!" -ForegroundColor Yellow
            if (-not(AskForChoice "Continue despite the domain difference?")){
                Write-Host "Aborting."
                Stop-Transcript | Out-Null
                if (Test-Path $TranscriptTemp) {Remove-Item $TranscriptTemp -Force}
                Start-Sleep 2
                continue
            }
        }
        #endregion Get domain from first entry (optional)
        $processed=0
        if ($true)
        { ## continue choices
            $choiceLoop=0
            $i=0        
            foreach ($row in $rows)
            { # each row
                $i++
                write-host "----- $i of $rowscount $row"
                if ($choiceLoop -ne "Yes to All")
                {
                    $choiceLoop = AskForChoice "Process entry $($i) ?" -Choices @("&Yes","Yes to &All","&No","No and E&xit") -Default 1 -ReturnString
                }
                if (($choiceLoop -eq "Yes") -or ($choiceLoop -eq "Yes to All"))
                { # choiceloop
                    $processed++
                    #region    --------- Custom code for object $row
                    # Find user
                    $user = Get-MgUser -UserId $row.UserPrincipalName -Property "displayName,userPrincipalName,passwordPolicies" -ErrorAction SilentlyContinue
                    if (-not $user) {
                        Write-Host "User with UserPrincipalName '$($row.UserPrincipalName)' not found. [And will be skipped]" -ForegroundColor Yellow
                        PressEnterToContinue
                        Continue
                    }
                    # Check domain
                    # Check existing password policy
                    $orgExpiresDomain = $false
                    $userdomain = $user.UserPrincipalName.Split("@")[1]
                    $mgdomain = $mgdomains | Where-Object { $_.Id -ieq $userdomain }
                    if ($mgdomain) {
                        Write-Host "User domain: $userdomain" -NoNewline
                        if (($null = $mgdomain.PasswordValidityPeriodInDays) -or ($mgdomain.PasswordValidityPeriodInDays -ne 2147483647)) {
                            Write-Host " (Org expires password every: $($mgdomain.PasswordValidityPeriodInDays) days)" -ForegroundColor Yellow -NoNewline
                            $orgExpiresDomain = $true
                        } else {
                            Write-Host " (Org doesn't enforce password expiration)" -ForegroundColor Green -NoNewline
                            $orgExpiresDomain = $false
                        }
                    } else {
                        Write-Host "User domain: $userdomain (not found in tenant!)" -ForegroundColor Yellow -NoNewline
                        $orgExpiresDomain = $false
                    }
                    # Org setting
                    if ($orgExpiresDomain) {
                        Write-Host " (User setting can override this)" -ForegroundColor Green
                    } else {
                        Write-Host " (User settings set by this code are ignored)" -ForegroundColor Yellow
                    }
                    $OldDisablePasswordExpiration = $user.PasswordPolicies -match "DisablePasswordExpiration"
                    $SetDisablePasswordExpiration = $row.DisablePasswordExpiration -eq "TRUE"
                    Write-Host "$($i): $($row.UserPrincipalName) [$($OldDisablePasswordExpiration)] "
                    Write-Host "  Desired DisablePasswordExpiration: " -NoNewline
                    Write-Host $SetDisablePasswordExpiration -ForegroundColor Green
                    Write-Host "  Current DisablePasswordExpiration: " -NoNewline
                    # See if they already match
                    if ($SetDisablePasswordExpiration -eq $OldDisablePasswordExpiration)
                    { # They match
                        Write-Host $OldDisablePasswordExpiration -ForegroundColor Green
                        Write-Host "[OK] No change needed." -ForegroundColor Green
                    } else {
                        Write-Host $OldDisablePasswordExpiration -ForegroundColor Yellow
                        If ($SetDisablePasswordExpiration) {
                            $newPolicies = "DisablePasswordExpiration"
                        } else {
                            $newPolicies = "None"
                        }
                        Write-Host "Changing from [$($user.PasswordPolicies)] to [$($newPolicies)] " -ForegroundColor Yellow -NoNewline
                        # Update user
                        try {
                            Update-MgUser -UserId $row.UserPrincipalName -PasswordPolicies $newPolicies
                            Write-Host "Updated." -ForegroundColor Green
                        } catch {
                            Write-Host "Error updating user: $_" -ForegroundColor Red
                        }
                    } # They don't match
                    #endregion --------- Custom code for object $row
                } # choiceloop
                if ($choiceLoop -eq "No")
                {
                    write-host ("Entry ($i) skipped.")
                }
                if ($choiceLoop -eq "No and Exit")
                {
                    write-host "Aborting."
                    Break
                }
            } # each row
        } ## continue choices
        Write-Host "------------------------------------------------------------------------------------"
        Write-Host "Done. $($processed) of $($rowscount) entries processed. Press [Enter] to exit."
        Write-Host "------------------------------------------------------------------------------------"
    } # Change user password expiration
} until ($choice -eq "Exit") # loop until exit
#region    ---- Transcript Save
Stop-Transcript | Out-Null
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$TranscriptTarget = "$(Split-Path $scriptFullname -Parent)\Logs\$($scriptBase)_$($date)_log.txt"
New-Item (Split-Path $TranscriptTarget -Parent) -ItemType Directory -Force | Out-Null # Make Logs folder
Move-Item $TranscriptTemp $TranscriptTarget -Force
#endregion ---- Transcript Save