<#

.SYNOPSIS
Connect to Office 365 services using modern or legacy authentication

.DESCRIPTION
Connect to Office 365 services using modern or legacy authentication

.EXAMPLE
Connect-Tenant -Username user@contoso.com
Connects to AzureAD, Exchange Online, and MS Online Services.

.EXAMPLE
Connect-Tenant -Username user@contoso.com -Sharepoint -orgname "contoso"
Connects to AzureAD, Exchange Online, MS Online, and Sharepoint Online

.EXAMPLE
Connect-Tenant -Username user@contoso.com -SkypeOnline
Connects to Skype for Business Online, MS Online, Exchange Online, and AzureAD

.EXAMPLE
Connect-Tenant -Username user@contoso.com -SecurityCenter
Connects to Exchange Online, AzureAD, MS Online, and Security and Compliance Center.

.EXAMPLE
Connect-Tenant -LegacyAuth
Connects to AzureAD, Exchange Online, and MS Online Services using legacy authentication

.Notes
Updated 9/1/2021 by Chaim Black.
#>

function Connect-Tenant {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [String]$UserName,
        [Parameter()]
        [switch]$AzureAD,
        [Parameter()]
        [switch]$SelectConnections,
        [Parameter()]
        [switch]$MSOnline,
        [Parameter()]
        [switch]$ExchangeOnline,
        [Parameter()]
        [switch]$SecurityCenter,
        [Parameter()]
        [switch]$SkypeOnline,
        [Parameter()]
        [switch]$SharePoint,
        [Parameter()]
        [string]$OrgName,
        [Parameter()]
        [switch]$LegacyAuth
    )

    if (!($SelectConnections)) {
        $AzureAD         = $true
        $ExchangeOnline  = $true
        $MSOnline        = $True
    }

    if ($LegacyAuth) {
        $userCredential = Get-Credential
        if (!($userCredential)) {Write-Verbose "No credentials detected."; return}
        else {Write-Verbose "Credentials detected."}
    }
    Else {
        if (!($Username)) {
            $username = Read-Host "Please enter a username"
        }
    }

    $Modules = Get-Module -ListAvailable

    If (!($Modules.name -like "*ExchangeOnlineManagement*")) {
        Install-Module -Name ExchangeOnlineManagement -Force:$true
    }

    Import-Module 'ExchangeOnlineManagement'    
    $Module = Get-Module -Name 'ExchangeOnlineManagement'
    if ($Module.Version -lt '2.0.5') {
        Update-Module -Name 'ExchangeOnlineManagement' -Force:$true
        Remove-Module 'ExchangeOnlineManagement'
        Import-Module 'ExchangeOnlineManagement'
    }
    

    #Connect to AzureAD
    If ($AzureAD) {        
        #Check for/Install AzureAD
        If (!($Modules.name -like "*Azuread*")) {
            Write-Verbose "AzureAD module is not installed."    
            If (!(Get-PackageProvider | Where-Object {$_.name -like "*nuget*"})) {
                Write-Verbose "Nuget Package Provider is not installed."
                Write-Verbose "Installing the Nuget Package Provider."
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -force
                if ($?) {Write-Verbose "Installation of Nuget Package Provider succeeded."} else {Write-Verbose "Installation of Nuget Package Provider failed."}    
            }

            Write-Verbose "Installing AzureAD Module."            
            Install-Module -Name AzureAD -Force
            if ($?) {Write-Verbose "Installation of AzureAD module succeeded."} else {Write-Verbose "Installation of AzureAD module failed."}
            Start-Sleep -Seconds 2
            Import-Module AzureAD 
        }

        if ($LegacyAuth) {
            Write-Host "Attempting connection to AzureAD" -ForegroundColor White
            Try {$AzureADConnect = Connect-AzureAD -Credential $userCredential -ErrorAction SilentlyContinue -InformationAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null}
            Catch {Write-Verbose "Attempting Connection to AzureAD"}
        }
        Else {
            Write-Host "Attempting connection to AzureAD" -ForegroundColor White
            $AzureADConnect = Connect-AzureAD -ErrorAction SilentlyContinue -InformationAction SilentlyContinue -WarningAction SilentlyContinue
        }

        try {$AzureADResult = Get-AzureADDomain}
        Catch {Write-verbose 'Testing Connection to AzureAD'}

        if ($AzureADResult) {
            Write-Host 'Connection to AzureAD: Success' -ForegroundColor Green
        }
        Else { Write-Host 'Connection to AzureAD: Failed' -ForegroundColor Red}
    }

   
    If ($MSOnline) {
        #Connect to MS Online:
        If (!($Modules.name -like "*msonline*")) {
            If (!(Get-PackageProvider | Where-Object {$_.name -like "*nuget*"})) {
                Write-Verbose "Nuget Package Provider is not installed."
                Write-Verbose "Installing the Nuget Package Provider."
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -force
                if ($?) {Write-Verbose "Installation of Nuget Package Provider succeeded."} else {Write-Verbose "Installation of Nuget Package Provider failed."}
            }

            Write-Verbose "Microsoft Online Services module is not installed."
            Write-Verbose "Installing Microsoft Online Services Module."
            Install-Module -Name MSOnline -Force
            if ($?) {Write-Verbose "Installation of Microsoft Online Services module succeeded."} else {Write-Verbose "Installation of Microsoft Online Services module failed."}

            Import-Module MSOnline
        }

        if ($LegacyAuth) {
            Write-Host "Attempting connection to MSOnline" -ForegroundColor White
            $MsonlineConnect = Connect-MsolService -Credential $userCredential -ErrorAction SilentlyContinue -InformationAction SilentlyContinue -WarningAction SilentlyContinue
        }
        Else {
            Write-Host "Attempting connection to MSOnline" -ForegroundColor White
            $MsonlineConnect = Connect-MsolService -ErrorAction SilentlyContinue -InformationAction SilentlyContinue -WarningAction SilentlyContinue
        }

        try {$MSOnlineResult = Get-MsolCompanyInformation -ErrorAction SilentlyContinue }
        Catch {Write-verbose 'Testing Connection to MSOnline'}

        if ($MSOnlineResult) {
            Write-Host 'Connection to MSOnline: Success' -ForegroundColor Green
        }
        Else { Write-Host 'Connection to MSOnline: Failed' -ForegroundColor Red}

        if ($MSOnlineResult) {
            #Verify account is global admin
            $role = Get-MsolRole -RoleName "Company Administrator"
            if ($LegacyAuth) { $gadmin = $userCredential.UserName}
            Else {$gadmin = $UserName}
            If (((Get-MsolRoleMember -RoleObjectId $role.ObjectId).EmailAddress) -like "$gadmin") {}
            Else {
            Write-Host 'Error: User is not a global administrator.' -ForegroundColor Red   
            }
        }
    }

    if ($ExchangeOnline) {
        if ($LegacyAuth) {
            Write-Host "Attempting connection to ExchangeOnline" -ForegroundColor White
            Connect-ExchangeOnline  -Credential $userCredential -ShowProgress:$true -ShowBanner:$False
        }
        Else {
            Write-Host "Attempting connection to ExchangeOnline" -ForegroundColor White
            Connect-ExchangeOnline -UserPrincipalName $Username -ShowProgress:$true -ShowBanner:$False
        }
        
        try {
            $ExchangeResult = get-command get-mailbox -ErrorAction SilentlyContinue
        }
        Catch {Write-verbose 'Testing Connection to Exchange'}

        if ($ExchangeResult) {
            Write-Host 'Connection to Exchange Online: Success' -ForegroundColor Green
        }
        Else { Write-Host 'Connection to Exchange Online: Failed' -ForegroundColor Red}
    }
    
    if ($SecurityCenter) {         
        if ($LegacyAuth) {
            Write-Host "Attempting connection to Security Center" -ForegroundColor White
            Connect-IPPSSession  -Credential $userCredential -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        }
        Else {
            Write-Host "Attempting connection to Security Center" -ForegroundColor White
            Connect-IPPSSession -UserPrincipalName $Username -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        }

        try {$SecurityResult = get-command Get-ProtectionAlert -ErrorAction SilentlyContinue}
        Catch {Write-Verbose 'Testing Connection to Security Center'}

        if ($SecurityResult) {
            Write-Host 'Connection to Security Center: Success' -ForegroundColor Green
        }
        Else { Write-Host 'Connection to Security Center: Failed' -ForegroundColor Red}        
    }

    if ($SkypeOnline) {
        If (!(Get-WmiObject -Class Win32_Product | Where-Object {$_.name -like "*skype*"})) {
            Write-Host 'Connection to Skype: Failed - Please manually install Skype module and reboot your computer.' -ForegroundColor Red
            Write-Host 'Link: https://www.microsoft.com/en-us/download/details.aspx?id=39366' -ForegroundColor Red
        }
        Else {
            Import-Module SkypeOnlineConnector
            if ($LegacyAuth) {
                Write-Host "Attempting connection to SFB" -ForegroundColor White
                $sfbSession = New-CsOnlineSession -Credential $userCredential
            }
            Else {
                Write-Host "Attempting connection to SFB" -ForegroundColor White
                $sfbSession = New-CsOnlineSession -UserName $Username
            }
            $sfbresultSession = Import-PSSession $sfbSession -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue

            try {$SFBResult = get-command Get-CsOnlineUser -ErrorAction SilentlyContinue}
            Catch {Write-Verbose 'Testing Connection to SFB'}

            if ($SFBResult) {
                Write-Host 'Connection to Skype: Success' -ForegroundColor Green
            }
            Else { Write-Host 'Connection to Skype: Failed' -ForegroundColor Red}
        }
    }

    if ($SharePoint) {
        if (!($OrgName)) {Write-host 'Missing orginization name - add -orgname variable'}
        if (!(Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable)) {
            $Result = Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force:$True -ErrorAction SilentlyContinue -WarningAction SilentlyContinue -InformationAction SilentlyContinue
        }
        if (!(Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable)) {
            write-host "Failed to install SharePoint PowerShell Module"
        }
        Write-Host "Attempting connection to Sharepoint Online" -ForegroundColor White
        Connect-SPOService -Url https://$orgName-admin.sharepoint.com

        try {$SharePointResult = get-command Get-SPOHomeSite -ErrorAction SilentlyContinue}
        Catch {Write-verbose 'Testing Connection to SFBSharPoint'}

        if ($SharePointResult) {
            Write-Host 'Connection to Sharepoint: Success' -ForegroundColor Green
        }
        Else { Write-Host 'Connection to Sharepoint: Failed' -ForegroundColor Red}
    }
}