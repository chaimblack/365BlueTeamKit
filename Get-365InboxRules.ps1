<#
.SYNOPSIS
Search and audit all inbox rules for an entire Office 365 Tenant.
 
.DESCRIPTION
This script captures all user inbox rules, looks at several attributes which are often misused by attackers, and exports it to a xlsx (default) or csv. It also outputs the raw results to json.

.PARAMETER NoLaunch
Prevents the output folder from launching after completion of command. Default is set to open folder.

.PARAMETER Path
Sets the path of the output. Default path is C:\PSOutput\Get-365InboxRules\

.PARAMETER ExportCSV
By default, this script exports as an Excel file. Setting this switch exports as csv instead.

.PARAMETER OutObject
Displays all rules as an array. Can be used if you want to pipe it to something else.

.PARAMETER OutRawObject
Displays all raw rules as an array. Can be used if you want to pipe it to something else.

.PARAMETER Username
By default, this script searches an entire Office 365 tenant. You can use this to specify a single user to search instead.

.PARAMETER DisplayWrapped
Displays all rules wrapped in note format.

.PARAMETER RemoveUserRules
Removes all user inbox rules. Can only be used if username is specified.

.EXAMPLE 
Get-365InboxRules
Verifies connection to Office 365 Exchange Online, runs foreach loop on all mailboxes to capture and output all user inbox rules. Looks at several attributes which are often misused by attackers. Exports to XLSX and JSON.

.NOTES
Created by Chaim Black on 1/8/2021.
Last updated: 9/13/2021

Get-365InboxRules was created to audit user inbox rules in Office 365. 

Attackers typically create malicious inbox rules after compromising an account, often redirecting, forwarding, moving, or deleting mail. 
This script captures all user inbox rules, looks at several attributes which are often misused by attackers, and exports it to a xlsx (default) or CSV.
It also outputs the raw results to json.

As a reminder, always check user inbox rules when investigating an email compromise.
When running this script in larger companies, I often find either active email compromises, or remnants of prior compromises which was not fully cleaned up.
#>
Function Get-365InboxRules {
    
    [CmdletBinding()]
    Param(
        [Parameter()]
        [switch]$NoLaunch,
        [Parameter()]
        [string]$Path,
        [Parameter()]
        [switch]$ExportCSV,
        [Parameter()]
        [switch]$OutObject,
        [Parameter()]
        [switch]$OutRawObject,
        [Parameter()]
        [string]$Username,
        [Parameter()]
        [switch]$DisplayWrapped,
        [Parameter()]
        [switch]$RemoveUserRules
    )
    

    <########################################
            Prerequisites
    ########################################>

    If (!(Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue)) {
        Write-host 'Error: Not Connected to Exchange Online.' -ForegroundColor Red
        Write-Host 'Please see https://www.powershellgallery.com/packages/ExchangeOnlineManagement for installation instructions' -ForegroundColor Green
        break
    }

    $Date = (Get-Date).toString("MM-dd-yy hh-mm-ss")

    if (!($path)) {
        $Company = ((Get-MsolCompanyInformation).displayname).trim()
        $SaveLocation = "C:\PSOutput\Get-365InboxRules\$Company"
        If (!(test-path $SaveLocation)) {
            New-Item -ItemType Directory -Force -Path $SaveLocation | Out-Null
        }
    }
    Else {
        $SaveLocation = $Path
    }

    If (!($ExportCSV)) {
        If (!(Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
            Write-Host "Missing prerequisites. Installing ImportExcel PS module now..."
            Install-Module ImportExcel -Force -AllowClobber

            If ($?) {Write-Host "Reporting module installed successfully." ; Import-Module ImportExcel}
            else {Write-Host "Reporting module failed to install." -ForegroundColor Red ; return }
        }
        Else {Import-Module ImportExcel}
    }

    <########################################
                Script
    ########################################>

    if ($Username) {
        If (!(Get-Mailbox -Identity $username)) {
            Write-Host 'Error: Invalid mailbox listed.' -ForegroundColor Red
            break
        }
        $Users = Get-Mailbox -Identity $Username
    }
    Else {
        $Users = Get-Mailbox -ResultSize unlimited
    }

    $AllIrules = foreach ($user in $users) {
            
        $rules = Get-InboxRule -Mailbox $user.UserPrincipalName | Select-Object *
        
        [array]$InboxRules = foreach ($rule in $rules) {

            $IRule = (($rule[0]).Description).Replace("`r",' ').Replace("`n",' ').Replace("`t",' ') 

            if ($rule.DeleteMessage) {$Delete = $True} Else {$delete = $False}

            if ($Rule.MoveToFolder) {$Move = $True} Else {$Move = $False}

            if ($Rule.MarkAsRead) {$MarkAsRead = $True} Else {$MarkAsRead = $False}

            if (
                $Rule.ForwardTo  `
                -or $Rule.ForwardAsAttachmentTo  `
                -or $Rule.RedirectTo
            ) {$Forward = $True} Else {$Forward = $False}

            if (
                $Rule.MyNameInCcBox  `
                -or $Rule.MyNameInToBox  `
                -or $Rule.MyNameInToOrCcBox `
                -or $Rule.SentOnlyToMe `
                -or $Rule.Description -notlike "*If the message:*"
            ) {$ApplyAll = $True} Else {$ApplyAll = $False}

            if (
                $Rule.ReceivedAfterDate  `
                -or $Rule.ReceivedBeforeDate
            ) {$DateBefOrAf = $True} Else {$DateBefOrAf = $False}

            if (
                $Rule.WithinSizeRangeMaximum  `
                -or $Rule.WithinSizeRangeMinimum
            ) {$Size = $True} Else {$Size = $False}

            $ForwardingTo = $null; $AllForwardExt = $null; $ForwardExt = $null; $CanFwdExt = $null; $AllFwd = $null
            $AllFwd = if ($Forward) {
                foreach ($Fwd in ($Rule.ForwardTo)) {
                    ($Fwd.Split('"'))[1]
                }
            
                foreach ($Fwd in ($Rule.ForwardAsAttachmentTo)) {
                    ($Fwd.Split('"'))[1]
                }
            
                foreach ($Fwd in ($Rule.RedirectTo)) {
                    ($Fwd.Split('"'))[1]
                }
            }
            $ForwardingTo = $AllFwd | Select-Object -Unique
            $AllForwardExt = $ForwardingTo | Where-Object {$_ -like "*@*"}
            $ForwardExt = foreach ($External in $AllForwardExt) {
                if ( ($External -split '@')[1] -notin (Get-AcceptedDomain).DomainName) {$External}
            }
            if ($ForwardExt) {$CanFwdExt = $true} 
            Else {$CanFwdExt = $false}
            if (!($Forward)) {$CanFwdExt = $false}

            [PSCustomObject]@{
                'User'            = $user.UserPrincipalName
                'Name'            = $rule.name
                'Enabled'         = $rule.enabled
                'Delete'          = $Delete
                'ApplyAll'        = $ApplyAll
                'Date'            = $DateBefOrAf
                'Size'            = $Size
                'Move'            = $Move
                'MarkAsRead'      = $MarkAsRead
                'FwdorRedir'      = $Forward                    
                'FwdorRedirExt'   = $CanFwdExt
                'Rule'            = $IRule
                'FwdorRedirTo'    = $ForwardingTo -join '; '
                'FwdorRedirExtTo' = $ForwardExt -join '; '
                'UserObjectID'    = $User.ExternalDirectoryObjectId
                'RuleIdentity'    = $Rule.Identity
                'RawRule'         = $Rule
            }
        }
        
        
        [PSCustomObject]@{
            'InboxRules' = $InboxRules
            'RawExport'  = $rules
        }
    }

    $AllRules1   = ($AllIrules | Where-Object {$_.inboxrules}).InboxRules
    $AllRules    = $AllRules1 | Select-Object * -ExcludeProperty RawRule
    $AllRawRules = ($AllIrules | Where-Object {$_.RawExport}).RawExport

    $Suspicious = $Allrules | Where-Object {            
        $_.FwdorRedirExt `
        -or $_.Rule -like "*hacked*" `
        -or $_.Rule -like "*move the message to folder 'RSS Feeds'*" `
        -or $_.Rule -like "*If the message:   the message was sent only to me.  Take the following actions:   delete the message   and stop processing more rules on this message*" `
        -or $_.Rule -like "*If the message:   my name is in the To or Cc box  Take the following actions:   delete the message   and stop processing more rules on this message*" `
        -or $_.Rule -like "*If the message:   my name is in the To box  Take the following actions:   delete the message   and stop processing more rules on this message*" `
        -or $_.Rule -like "Take the following actions:   move the message to folder 'Junk Email'   and stop processing more rules on this message*" `
        -or $_.Rule -like "Take the following actions:   delete the message   and stop processing more rules on this message*" `
        -or $_.Name -like "*..*"
    }

    <########################################
              Export | Output
    ########################################>

    if ($AllRules) {

        if ($ExportCSV) {
            $OutCSVPath = "$SaveLocation\User Inbox Rules - $date.csv"
            $HashOutput = "$SaveLocation\User Inbox Rules - $date.csv.hash.csv"

            $Allrules |
                Export-Csv -Path $OutCSVPath
            
            Get-ItemHash -Default -FilePath $OutCSVPath

            if ($Suspicious) {
                $SuspiciousOutput = "$SaveLocation\User Inbox Rules - Suspicious - $date.csv"

                $Suspicious |
                    Export-Csv -Path $SuspiciousOutput
                
                Get-ItemHash -Default -FilePath $SuspiciousOutput
            }

            
        }
        Else {
            $ExcelOutput = "$SaveLocation\User Inbox Rules - $date.xlsx"

            $Allrules | Export-ExcelDefault -Path $ExcelOutput -WorksheetName "Inbox Rules"

            if ($Suspicious) {
                $Suspicious | Export-ExcelDefault -Path $ExcelOutput -WorksheetName "Suspicious"
            }
            Get-ItemHash -Default -FilePath $ExcelOutput           
        }

        $AllRawRules | ConvertTo-Json -Depth 100 | Set-Content -Path "$SaveLocation\User Inbox Rules - $date.json" 
        Get-ItemHash -Default -FilePath "$SaveLocation\User Inbox Rules - $date.json"        

        if ($OutObject) {
            $Allrules
        }

        if ($OutRawObject) {
            $AllRawRules
        }

        if (!($NoLaunch)) {
            try {
                Start-Process "$SaveLocation" | Out-Null
            } 
            catch {
                Write-Verbose "No results."
            }
        }

        $AllWrapped = foreach ($IRule in $AllRules1) {

            $Username = $IRUle.User
            $Enabled  = $IRule.Enabled 
            $RuleR    = $IRule.rawrule | Select-Object description -ExpandProperty description
            $RuleID   = $IRule.rawrule.Identity

            [PSCustomObject]@{
                'Data'  = "UserName: $Username"
            }

            [PSCustomObject]@{
                'Data'  = "RuleIdentity: $RuleID"
            }

            [PSCustomObject]@{
                'Data'  = "Enabled: $Enabled"
            }

            [PSCustomObject]@{
                'Data'  = "Rule:"
            }

            [PSCustomObject]@{
                'Data'  = $RuleR
            }

            [PSCustomObject]@{
                'Name'  = ''
                'Value' = ''
            }
        }
        
        if ($DisplayWrapped) {
            $AllWrapped.data | Format-Table -Wrap
        }

        $AllWrapped.data | Out-File -FilePath "$SaveLocation\User Inbox Rules - $date.txt"
        Get-ItemHash -FilePath "$SaveLocation\User Inbox Rules - $date.txt" -Default       
    }

    if ($RemoveUserRules -and ($Username)) {
        if ($Allrules) {

            foreach ($Rule in $AllRules) {            
                Remove-InboxRule -Confirm:$false -Identity $Rule.RuleIdentity
            }

            $VerifyAllRules = Get-InboxRule -Mailbox $Username

            if ($VerifyAllRules) {
                Write-Host "Error: Failed to remove user inbox rules. Please do so manually." -ForegroundColor Red
            }
            Else{
                Write-Host "Removed user inbox rules." -ForegroundColor Green
            }
        }
        Else{
            Write-Host "No inbox rules found." -ForegroundColor Green
        }
    }
}