<#
.SYNOPSIS
Searches Office 365 Audit Logs.

.DESCRIPTION
Searches Office 365 Audit Logs and does various funtions.

.PARAMETER Username
Specifies username. Multiple usernames should be seperated by commas and in quotes.

.PARAMETER StartDate
Defines start date. Default is 7 days prior to current date.

.PARAMETER EndDate
Defines end date.

.PARAMETER SkipUAL
When performing search on mailbox, skips search on Unified Audit Log and only searches mailbox logs

.PARAMETER All
Sets the start date 90 days prior current date. Using this automatically sets the end date to the current date. Update 9/2021: Microsoft allows for searching up to 1 year. Due to the ammount of data which it would capture, the "All" switch still only goes 90 days but you can manually set it to 365 days.

.PARAMETER Days
Defines how many days ago the start date should be. Max is 364.

.PARAMETER IPAddress
Defines IP address to search. Multiple IP addresses should be seperated by commas and in quotes. This command can be combined with -username to only display logins to specific user accounts from specified IP address.

.PARAMETER FailedSignIns
Checks for failed sign ins

.PARAMETER FailedSignInStatistics
Checks for failed sign in statistics for the past 7 days

.PARAMETER Path
Defines the output path. Default saves to C:\PSOutPut\Search-365AuditLogs

.PARAMETER DisplayOutput
Defines if output should be displayed

.PARAMETER IncludeMailboxLogs
Also searchs mailbox logs when searching for login locations. This increases the time of the script.

.PARAMETER NoLaunch
Prevents Windows Explorer from launching after search.

.PARAMETER InboxRules
Searches for inbox rules

.PARAMETER NoLookup
Does not lookup IP information.

.PARAMETER SkipMailboxLogs
Does not search mailbox logs. Only searches Unified Audit Logs.

.PARAMETER MailboxPermissions
Searches 'Add-MailboxPermission' operation and collects data.

.PARAMETER AzureLogs
Defines path of Azure AD log file in JSON format. If defined, will import and include data in results.

.EXAMPLE
Search-365AuditLogs -username John@contoso.com
Checks to see if audit logs are enabled, and produces report. Output location automatically opens upon completion.

.EXAMPLE
Search-365AuditLogs -username John@contoso.com -startdate 01/01/2018 -enddate 1/15/2018
Checks to see if audit logs are enabled, and produces logging report from 01/01/2018 to 1/15/2018

.EXAMPLE
Search-365AuditLogs -username John@contoso.com -startdate 01/01/2018 -enddate 1/15/2018 -IPAddress 8.8.8.8
Checks to see if audit logs are enabled, and produces logging report from 01/01/2018 to 1/15/2018. Only includes data from 8.8.8.8 IP address.

.EXAMPLE
Search-365AuditLogs -username John@contoso.com -All
Checks to see if audit logs are enabled, and produces logging report for 90 days from current date

.EXAMPLE
Search-365AuditLogs -username John@contoso.com -Path "C:\temp"
Checks to see if audit logs are enabled, produces report, and saves to "C:\temp"

.EXAMPLE
Search-365AuditLogs -username John@contoso.com -DisplayOutput
Checks to see if audit logs are enabled, produces report, and displays it in the terminal

.EXAMPLE
Search-365AuditLogs -IPAddress 8.8.8.8
Checks all activity from IP Address 8.8.8.8. Allows custom date range.

.EXAMPLE
Search-365AuditLogs -FailedSignIns
Searches all failed sign in attempts and produces report. Allows custom date range.

.EXAMPLE
Search-365AuditLogs -FailedSignInStatistics
Searches all failed sign in attempts for the past 7 days and produces report with statistics.

.NOTES
Last updated 9/10/2021 Chaim Black.
#>

function Search-365AuditLogs {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [string]$UserName,
        [Parameter()]
        [string]$StartDate,
        [Parameter()]
        [string]$EndDate,
        [Parameter()]
        [int]$Days,
        [Parameter()]
        [switch]$All,
        [Parameter()]
        [switch]$SkipUAL,
        [Parameter()]
        [string]$IPAddress,
        [Parameter()]
        [string]$ObjectIDs,
        [Parameter()]
        [switch]$FailedSignIns,
        [Parameter()]
        [string]$Path,
        [Parameter()]
        [switch]$DisplayOutput,
        [Parameter()]
        [switch]$FailedSignInStatistics,
        [Parameter()]
        [switch]$IncludeMailboxLogs,
        [Parameter()]
        [switch]$SkipMailboxLogs,
        [Parameter()]
        [string]$Inputlog,
        [Parameter()]
        [string]$CompanyName,
        [Parameter()]
        [switch]$NoLaunch,
        [Parameter()]
        [switch]$InboxRules,
        [Parameter()]
        [switch]$NoLookup,        
        [Parameter()]
        [switch]$MailboxPermissions,        
        [Parameter()]
        [string]$AzureLogs
    )

    #########################
    #Mandatory and prerequisites
    #########################

    if ($username) {$Username = $username.ToLower()}       

    if (!($Inputlog)) {
        If (!(Get-Command -Name get-mailbox -ErrorAction SilentlyContinue)) {
            Write-host "Error - Not Connected to Exchange Online. Please run Connect-Tennant" -ForegroundColor Red
            break
        }
        If (!(Get-MsolCompanyInformation -ErrorAction SilentlyContinue)) {
            Write-host "Error - Not Connected to MS Online. Please run Connect-Tennant" -ForegroundColor Red
            break
        }
    }

    if ($Inputlog -and (!($CompanyName))) {
        Write-Host 'Error - Please enter the "-CompanyName" switch with the company name.' -ForegroundColor Red
        break
    }

    if ($Inputlog) {
        If (!(test-path $Inputlog)) {
            Write-Host "Error - Failed to find input file." -ForegroundColor Green
            break
        }
    }

    if ($NoLookup) {
        function Get-IPLookup {
            [CmdletBinding()]
            Param(
                [Parameter()]
                [string]$IPAddress,
                [Parameter()]
                [switch]$OutputObject
            )
        
            if ($OutputObject) {
                [PSCustomObject]@{
                    'IP'               = $IPAddress
                    'City'             = ' '
                    'State'            = ' '
                    'Country'          = ' '
                    'ISP'              = ' '
                    'IPType'           = ' '            
                    'CSV'              = ' '
                }
            }
            else {
                [PSCustomObject]@{
                    ' ' = ' '
                }
            }
        }
    }
    Else {
        if (!(get-command Get-IPLookUp -ErrorAction SilentlyContinue)) {
            Write-Host 'Error: Missing Get-IPLookup module.. You may also add the "NoLookup" switch to ignore.'  -ForegroundColor Red
            Break
        } 
    }

    #Check for Excel Reporting Module; Install if not found
    If (!(Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
        Write-Host "Missing ImportExcel module. Installing now..."
        Install-Module ImportExcel -Force -AllowClobber

        If ($?) {Write-Host "Reporting module installed successfully." ; Import-Module ImportExcel}
        else {Write-Host "Reporting module failed to install." -ForegroundColor Red ; return }
    }
    Else {Import-Module ImportExcel}
    

    #Set date start and end if not entered. Default end date is current date, and default start date is 7 days prior to current date.
    if ($all) {
        $EndDate = (get-date).AddDays(1).toString("MM-dd-yyyy")
        $StartDate = ((Get-Date).AddHours(-2150)).tostring("MM-dd-yyyy")
    }    
    elseif ($Days){
        if ($Days -gt '364') {Write-Host "Error: Maximum days is one year" -ForegroundColor Red; Break}
        $StartDate = ((Get-Date).AddDays(-$Days)).tostring("MM-dd-yyyy")
        $EndDate = (get-date).AddDays(1).toString("MM-dd-yyyy")
    }
    else {
        if (!($EndDate)) {
            $EndDate = (get-date).AddDays(1).toString("MM-dd-yyyy")
        }
        if (!($StartDate)) {
            $StartDate = ((Get-Date).AddHours(-168)).tostring("MM-dd-yyyy")
        }
    }
    
    $date = (get-date).toString("MM-dd-yy hh-mm-ss tt")

    if ($CompanyName) {
        $Company = $CompanyName
    }
    Else {
        $Company = ((Get-MsolCompanyInformation).displayname).trim()
    }

    if (!($Path)) {
        $OutputLocation = "C:\PSOutPut\Search-365AuditLogs\$Company"
        If (!(test-path $OutputLocation)) {
            New-Item -ItemType Directory -Force -Path $OutputLocation | Out-Null
        }
    } 
    Else {
        $OutputLocation = $Path
    }

    if ($Username) {
        if ($Username -notlike "*,*") { 
            $OutPath      = "$OutputLocation\Audit Log " + $username.Split("@")[0] + " " + $Date + ".xlsx"
            $UALJSONPath  = "$OutputLocation\Audit Log " + $username.Split("@")[0] + " " + $Date + " - RawData-UAL.json"
            $MLJSONPath   = "$OutputLocation\Audit Log " + $username.Split("@")[0] + " " + $Date + " - RawData-ML.json"
        }
        if ($Username -like "*,*") {
            $OutPath      = "$OutputLocation\Audit Log " + " - Users - " + " " + $Date + ".xlsx"
            $UALJSONPath  = "$OutputLocation\Audit Log " + " - Users - " + " " + $Date + " - RawData-UAL.json"
            $MLJSONPath   = "$OutputLocation\Audit Log " + " - Users - " + " " + $Date + " - RawData-ML.json"
        }
    }

    $regex = "((^\s*((([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]))\s*$)|(^\s*((([0-9A-Fa-f]{1,4}:){7}([0-9A-Fa-f]{1,4}|:))|(([0-9A-Fa-f]{1,4}:){6}(:[0-9A-Fa-f]{1,4}|((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3})|:))|(([0-9A-Fa-f]{1,4}:){5}(((:[0-9A-Fa-f]{1,4}){1,2})|:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3})|:))|(([0-9A-Fa-f]{1,4}:){4}(((:[0-9A-Fa-f]{1,4}){1,3})|((:[0-9A-Fa-f]{1,4})?:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){3}(((:[0-9A-Fa-f]{1,4}){1,4})|((:[0-9A-Fa-f]{1,4}){0,2}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){2}(((:[0-9A-Fa-f]{1,4}){1,5})|((:[0-9A-Fa-f]{1,4}){0,3}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){1}(((:[0-9A-Fa-f]{1,4}){1,6})|((:[0-9A-Fa-f]{1,4}){0,4}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(:(((:[0-9A-Fa-f]{1,4}){1,7})|((:[0-9A-Fa-f]{1,4}){0,5}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:)))(%.+)?\s*$))"
    if ($ipaddress) {
        if ($ipaddress -notlike "*,*" -and $ipaddress -notmatch $regex) {
            Write-Host "Error: Invalid IP address" -ForegroundColor Red
            Break
        }
    }
    $sessionID = Get-random -maximum 5000

    #########################
    #Functions
    #########################

    #Validate Mailbox
    function Validate-Mailbox {
        #Check if username was entered
        if (!($username)) {
            Write-Host "Error: Please enter a username" -ForegroundColor Red
            break
        }

        #Verify mailbox is valid
        if ($UserName -and $Username -notlike "*,*") {
            If (!(Get-Mailbox -Identity $username -ErrorAction silentlycontinue)) {
                if (!($DisplayOutput)) {
                    Write-Host "Error - $username is not a valid mailbox" -ForegroundColor Red
                    Write-host "Only performing search on Unified Audit Logs" -ForegroundColor Green
                }

                $NoML            = $True
                $SkipMailboxLogs = $True
                $ALDisabled      = $True

            }
        }
        if ($Username -like "*,*") {
            $users = $username.Split(',')
            foreach ($user in $users){
                If (!(Get-Mailbox -Identity $user -ErrorAction silentlycontinue)) {
                    $SkipMailboxLogs = $True
                }
            }
        }

        if ($UserName -and $Username -notlike "*,*") {
            If (!(Get-MsolUser -UserPrincipalName $username -ErrorAction silentlycontinue)) {
                if (!($DisplayOutput)) {
                    Write-Host Error - $username is not a valid UserPrincipalName -ForegroundColor Red
                    break
                }
            }
        }
    }

    
    $UTCOffset = (Get-TimeZone).BaseUtcOffset.hours
    if ((Get-Date).IsDaylightSavingTime()) {$UTCOffset = $UTCOffset + 1}
    $UTCOffsetAdd = [math]::abs($UTCOffset)

    #########################

    if ($UserName -and (!($InboxRules)) -and (!($MailboxPermissions))) {
        Validate-Mailbox

        if ($IPAddress -and (!($SkipUAL))) {
            do {
                $auditdata3 += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId $sessionid -ResultSize 5000 -StartDate $startdate -EndDate $enddate -UserIds "$username" -IPAddresses $ipaddress
            }while (($auditdata3 | Measure-Object).count % 5000 -eq 0 -and ($auditdata3 | Measure-Object).count -ne 0 -and $auditdata3)
            $auditdata2 = $auditdata3 | Select-Object * -ExcludeProperty resultindex,resultcount -Unique | Sort-Object -Property creationdate -Descending
        }
        If ((!($IPAddress)) -and (!($SkipUAL))) {
            do {
                $auditdata3 += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId $sessionid -ResultSize 5000 -StartDate $startdate -EndDate $enddate -UserIds "$username"
            }while (($auditdata3 | Measure-Object).count % 5000 -eq 0 -and ($auditdata3 | Measure-Object).count -ne 0 -and $auditdata3)
            $auditdata2 = $auditdata3 | Select-Object * -ExcludeProperty resultindex,resultcount -Unique | Sort-Object -Property creationdate -Descending
        }
        If ($auditdata2) {           
            $auditdata1 = $auditdata2.auditdata

            $auditdata2 | ConvertTo-Json -Depth 100 | Out-File -FilePath $UALJSONPath
            Get-ItemHash -Default -FilePath $UALJSONPath

            $auditdata = $auditdata1 | 
                ForEach-Object -Process {
                try {
                    $_ | ConvertFrom-Json -ErrorAction SilentlyContinue
                } 
                catch {
                    Write-Verbose "can't convert file '$_' to JSON"
                }
            }

            $AuditDataRules = $AuditData | Where-Object {$_.Operation -like "New-InboxRule" -or $_.Operation -like "Set-InboxRule" -Or $_.Operation -like "UpdateInboxRules"}
            if ($AuditDataRules) {
                $RuleInfo = foreach ($AuditEntry in $AuditDataRules) {

                    $Udate = [datetime]$AuditEntry.CreationTime
                    $TimeData = $udate.tostring("MM/dd/yyyy hh:mm:ss tt")                                  

                    $UserID = $AuditEntry.UserID.ToLower()

                    $ClientIP1 = $null
                    if ($AuditEntry.ClientIP) {$ClientIP1 = $AuditEntry.ClientIP}
                    if ($AuditEntry.ClientIPAddress -and (!($AuditEntry.ClientIP))) {$ClientIP1 = $AuditEntry.ClientIPAddress}
                    if ($ClientIP1) {
                        $ClientIP = $ClientIP1.Replace('[','').Split(']')[0]
                        if ($Clientip -notmatch $regex) {$Clientip = $Clientip.split(':')[0]}
                        $IPLookup = Get-IPLookUp -IPAddress $ClientIP -OutputObject
                    }

                    [PSCustomObject]@{
                        'Name'  = 'UserName'
                        'Value' = $AuditEntry.UserID
                    }
                    [PSCustomObject]@{
                        'Name'  = 'DateUTC'
                        'Value' = $TimeData
                    }
                    [PSCustomObject]@{
                        'Name'  = 'ClientIP'
                        'Value' = $ClientIP
                    }
                    [PSCustomObject]@{
                        'Name'  = 'City'
                        'Value' = $IPLookup.City
                    }
                    [PSCustomObject]@{
                        'Name'  = 'State'
                        'Value' = $IPLookup.State
                    }
                    [PSCustomObject]@{
                        'Name'  = 'Country'
                        'Value' = $IPLookup.Country
                    }
                    [PSCustomObject]@{
                        'Name'  = 'ISP'
                        'Value' = $IPLookup.ISP
                    }
                    [PSCustomObject]@{
                        'Name'  = ' '
                        'Value' = ' '
                    }
                    $AuditEntry.parameters

                    [PSCustomObject]@{
                        'Name'  = ' '
                        'Value' = ' '
                    }
                    [PSCustomObject]@{
                        'Name'  = ' '
                        'Value' = ' '
                    }
                }
            }            

            $UALData = foreach ($i in $Auditdata) {

                $Udate = [datetime]$i.CreationTime                
                $TimeData = $udate.tostring("MM/dd/yyyy hh:mm:ss tt")
                                

                $UserID = $i.UserID.ToLower()

                $ClientIP1 = $null
                if ($i.ClientIP) {$ClientIP1 = $i.ClientIP}
                if ($i.ClientIPAddress -and (!($i.ClientIP))) {$ClientIP1 = $i.ClientIPAddress}
                if ($ClientIP1) {
                    $ClientIP = $ClientIP1.Replace('[','').Split(']')[0]
                    if ($Clientip -notmatch $regex) {$Clientip = $Clientip.split(':')[0]}
                    $IPLookup = Get-IPLookUp -IPAddress $ClientIP -OutputObject
                }

                if ($i.workload -like "OneDrive" -or $i.workload -like "*Sharepoint*") {
                    $Sharepoint = $i.ObjectId
                }
                Else { $Sharepoint = " " }

                [PSCustomObject]@{
                    'DateUTC'            =    $TimeData
                    'UserID'             =    $UserID.ToLower()
                    'ClientIP'           =    $ClientIP
                    'City'               =    $IPLookup.City
                    'State'              =    $IPLookup.State
                    'Country'            =    $IPLookup.Country
                    'ISP'                =    $IPLookup.ISP
                    'Operation'          =    $i.operation
                    'Workload'           =    $i.Workload
                    'ResultStatus'       =    $i.resultstatus
                    'LogonError'         =    $i.LogonError
                    'UserAgent'          =    ($i.extendedproperties | Where-Object {$_.name -like "UserAgent"}).Value
                    'ObjectID'           =    $Sharepoint
                }
            }

            if ($UALData) {
                $UALData | Export-ExcelDefault -Path $OutPath -WorkSheetName "Unified Audit Logs"
                
                if ($ObjectIDs) {
                    $SharePointOneDrive = $UALData | Where-Object {
                        $_.Workload -like "*OneDrive*" -or $_.Workload -like "*Sharepoint*" -and $_.objectid -like $ObjectIDs
                    }
                }
                Else {
                    $SharePointOneDrive = $UALData | Where-Object {
                        $_.Workload -like "*OneDrive*" -or $_.Workload -like "*Sharepoint*"
                    }
                }
                if ($SharePointOneDrive) {
                    $SharePointOneDrive | Export-ExcelDefault -Path $OutPath -WorkSheetName "OneDrive-Sharepoint-UAL"
                }                

                if ($Ruleinfo) {
                    $Ruleinfo |Export-ExcelDefault -Path $OutPath -WorkSheetName "Inbox Rules"
                }

                $ips = $UALData | Select-Object ClientIP,city,state,country,isp -Unique | 
                    Where-Object {$_.ClientIP -and $_.clientip -notlike "::1" -and $_.clientip -notlike '<null>'}
        
                $IPInfo = foreach ($ip in $ips) {
                    [PSCustomObject]@{      
                        'IP Address'                      = $ip.clientip
                        'City'                            = $ip.city
                        'state'                           = $ip.state
                        'country'                         = $ip.country
                        'ISP'                             = $ip.isp
                        'First Activity (UTC)'            = ($UALData | Where-Object {$_.ClientIP -like $ip.ClientIP} | Sort-Object -Property "Date*" | Select-Object * -first 1)."DateUTC"
                        'Last Activity (UTC)'             = ($UALData | Where-Object {$_.ClientIP -like $ip.ClientIP} | Sort-Object -Property "Date*" | Select-Object * -last 1)."DateUTC"
                        'Users Accessed'                  = ($UALData | Where-Object {$_.ClientIP -like $ip.ClientIP} | Select-Object UserID -Unique).userid -join "; "
                        'Total Connections'               = ($UALData| Where-Object {$_.ClientIP -like $ip.ClientIP}  | Measure-Object).count
                    }
                }
                if ($IPInfo) {
                    $IPInfo | Export-ExcelDefault -Path $OutPath -WorkSheetName "IP Addresses - UAL"                    
                }
            }
            if ($DisplayOutput) {              
                $UALData
            }
        }
        Else {
            Write-Host "No detection of unified audit log data." -ForegroundColor Green
        }        

        if (!($SkipMailboxLogs)) {
            #Check mailbox logs
            if ($Username -notlike "*,*") {
                If ( !((get-mailbox -Identity $username).auditenabled) ) {
                    Write-Host "Mailbox auditing is not enabled for user - enabling..." -ForegroundColor Green
                    Set-Mailbox -Identity $username -AuditEnabled $true
                    $ALDisabled = $True
                    Start-Sleep 10
                    If ((get-mailbox -Identity $username).auditenabled) {            
                        Write-Host "Enabled mailbox auditing" -ForegroundColor Green
                    }
                    Else {
                        Write-Host "Failed to enable Mailbox Auditing."
                    }
                }
            }
            If (!($ALDisabled)) {
                if ($Username -notlike "*,*") {
                    $MBLauditdata = Search-MailboxAuditLog -Identity (Get-Mailbox -Filter "UserPrincipalName -eq '$username'").UserPrincipalName -StartDate $StartDate -EndDate $EndDate -ResultSize 250000 -ShowDetails
                }
                if ($Username -like "*,*") {
                    $users = $username.Split(',')
                    $allmblauditdata0 = foreach ($user in $users){
                        Search-MailboxAuditLog -Identity (Get-Mailbox -Filter "UserPrincipalName -eq '$user'").UserPrincipalName -StartDate $StartDate -EndDate $EndDate -ResultSize 250000 -ShowDetails
                    }
                    $MBLauditdata = $allmblauditdata0  | Sort-Object -Property lastaccessed -Descending
                }

                if ($IPAddress) {
                    $IPAddresses1 = $IPAddress -split ','
                    $Filtered =  $MBLauditdata | Where-Object {$_.ClientIPAddress -in $IPAddresses1 -or $_.ClientIP -in $IPAddresses1}
                    $MBLauditdata = $Filtered
                }

                If (!($MBLauditdata)) {
                    Write-Host "No detection of mailbox audit log data." -ForegroundColor Green
                }
                Else {
                    $MBLauditdata | ConvertTo-Json -Depth 100 | Out-File -FilePath $MLJSONPath
                    Get-ItemHash -Default -FilePath $MLJSONPath

                    $MLData = foreach ($i in $MBLauditdata) {

                        $IPLookup = $null
                        $Udate   = [datetime]$i.LastAccessed                        
                        $TimeData = $udate.AddHours($UTCOffsetAdd).tostring("MM/dd/yyyy hh:mm:ss tt")
                                                                      

                        if ($i.ClientIPAddress) {
                            $ClientIP = $i.ClientIPAddress.Replace('[','').Split(']')[0]
                            if ($Clientip -notmatch $regex) {$Clientip = $Clientip.split(':')[0]}                            
                            $IPLookup = Get-IPLookUp -IPAddress $ClientIP -OutputObject                            
                        }
                        
                        [PSCustomObject]@{
                            'DateUTC'                   =    $TimeData
                            'UserID'                    =    $i.MailboxOwnerUPN
                            'ActingUserID'              =    $i.LogonUserDisplayName
                            'ClientIP'                  =    $i.ClientIPAddress
                            'City'                      =    $IPLookup.City
                            'State'                     =    $IPLookup.State
                            'Country'                   =    $IPLookup.Country
                            'ISP'                       =    $IPLookup.ISP
                            'Operation'                 =    $i.operation
                            'OperationResult'           =    $i.OperationResult
                            'SourceFolder'              =    $i.FolderPathName
                            'DestFolder'                =    $i.DestFolderPathName
                            'SourceItemSubjectsList'    =    $i.SourceItemSubjectsList
                            'SourceItemAttachmentsList' =    $i.SourceItemAttachmentsList
                            'ItemSubject'               =    $i.ItemSubject
                            'ItemAttachments'           =    $i.ItemAttachments
                            'UserAgent'                 =    $i.ClientInfoString
                            'Client'                    =    $i.ClientProcessName
                            'MessageID'                 =    $i.SourceItemInternetMessageIdsList
                        }                        
                    }
                    
                    if ($MLData) {
                        $MLData | Export-ExcelDefault -Path $OutPath -WorkSheetName "Mailbox Logs"
                    
                        if ($DisplayOutput) {                       
                            $MLData
                        }
                    }
                }
            }
        }

        if ($AzureLogs) {
            If ($NoLookup) {
                if ($IPAddress) { $AZData = Search-365AzureADLogs -StartDate $StartDate -EndDate $enddate -InputLogs $AzureLogs -Username $username -Path $OutputLocation -NoLaunch -DisplayOutput -IPAddress $IPAddress}
                Else            { $AZData = Search-365AzureADLogs -StartDate $StartDate -EndDate $enddate -InputLogs $AzureLogs -Username $username -Path $OutputLocation -NoLaunch -DisplayOutput}
            }
            Else {
                if ($IPAddress) { $AZData = Search-365AzureADLogs -StartDate $StartDate -EndDate $enddate -InputLogs $AzureLogs -Username $username -Lookup -Path $OutputLocation -NoLaunch -DisplayOutput -IPAddress $IPAddress}
                Else            { $AZData = Search-365AzureADLogs -StartDate $StartDate -EndDate $enddate -InputLogs $AzureLogs -Username $username -Lookup -Path $OutputLocation -NoLaunch -DisplayOutput}
            }

            If ($AZData) {
                $AZData | Export-ExcelDefault -Path $OutPath -WorkSheetName "Azure AD Logs"
            }
        }

        if ($MLData -or $UALData) {
            Get-ItemHash -Default -FilePath $OutPath
            if ((!($path) -and (!($NoLaunch)))) {
                try {
                    Start-Process "$OutputLocation" | Out-Null
                } 
                catch {
                    Write-Verbose "No results."
                }
            }
        }
    }

    if ($IPAddress -and (!($UserName))) {

        do {
            $auditdata3 += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId $sessionid -ResultSize 5000 -StartDate $startdate -EndDate $enddate -IPAddresses $ipaddress
        }while (($auditdata3 | Measure-Object).count % 5000 -eq 0 -and ($auditdata3 | Measure-Object).count -ne 0 -and $auditdata3)
        $auditdata2 = $auditdata3 | Select-Object * -ExcludeProperty resultindex,resultcount -Unique | Sort-Object -Property creationdate -Descending
        
        
        If ($auditdata2) {
            $auditdata1 = $auditdata2.auditdata

            $auditdata2 | ConvertTo-Json -Depth 100 | Out-File -FilePath "$OutputLocation\IP Search - $date - RawData-UAL.json"

            Get-ItemHash -Default -FilePath "$OutputLocation\IP Search - $date - RawData-UAL.json"

            $auditdata = $auditdata1 | 
                ForEach-Object -Process {
                try {
                    $_ | ConvertFrom-Json
                } 
                catch {
                    Write-Verbose "can't convert file '$_' to JSON"
                }
            }
            
            $Location = foreach ($i in $Auditdata) {

                $Udate = [datetime]$i.CreationTime
                $TimeData = $udate.tostring("MM/dd/yyyy hh:mm:ss tt")                

                $UserID = $i.UserID.ToLower()

                $ClientIP1 = $null
                if ($i.ClientIP) {$ClientIP1 = $i.ClientIP}
                if ($i.ClientIPAddress -and (!($i.ClientIP))) {$ClientIP1 = $i.ClientIPAddress}
                if ($ClientIP1) {
                    $ClientIP = $ClientIP1.Replace('[','').Split(']')[0]
                    if ($Clientip -notmatch $regex) {$Clientip = $Clientip.split(':')[0]}
                    $IPLookup = Get-IPLookUp -IPAddress $ClientIP -OutputObject
                }

                if ($i.workload -like "OneDrive" -or $i.workload -like "*Sharepoint*") {
                    $Sharepoint = $i.ObjectId
                }
                Else { $Sharepoint = " " }

                [PSCustomObject]@{
                    'DateUTC'            =    $TimeData
                    'UserID'             =    $UserID.ToLower()
                    'ClientIP'           =    $ClientIP
                    'City'               =    $IPLookup.City
                    'State'              =    $IPLookup.State
                    'Country'            =    $IPLookup.Country
                    'ISP'                =    $IPLookup.ISP
                    'Operation'          =    $i.operation
                    'Workload'           =    $i.Workload
                    'ResultStatus'       =    $i.resultstatus
                    'LogonError'         =    $i.LogonError
                    'UserAgent'          =    ($i.extendedproperties | Where-Object {$_.name -like "UserAgent"}).Value
                    'ObjectID'           =    $Sharepoint
                }                
            }

            $AllSuccess = $Location | Where-Object {$_.Operation -notlike "UserLoginFailed"}
            $ips = $AllSuccess | Select-Object ClientIP,city,state,country,isp -Unique
            $IPReport = foreach ($ip in $ips) {
                [PSCustomObject]@{      
                    'IP Address'                      = $ip.clientip
                    'City'                            = $ip.city
                    'state'                           = $ip.state
                    'Country'                         = $ip.country
                    'ISP'                             = $ip.isp
                    'First Activity (UTC)'            = ($AllSuccess | Where-Object {$_.ClientIP -like $ip.ClientIP} | Sort-Object -Property "Date*" | Select-Object * -first 1)."DateUTC"
                    'Last Activity (UTC)'             = ($AllSuccess | Where-Object {$_.ClientIP -like $ip.ClientIP} | Sort-Object -Property "Date*" | Select-Object * -last 1)."DateUTC"
                    'Users Accessed'                  = ($AllSuccess | Where-Object {$_.ClientIP -like $ip.ClientIP} | Select-Object UserID -Unique).userid -join ","
                    'Total Connections'               = ($AllSuccess | Where-Object {$_.ClientIP -like $ip.ClientIP} | Measure-Object).count                    
                }                
            }

            $SuccessUsers = $AllSuccess | Select-Object UserID -Unique
            $UserReport = foreach ($UserAcct in $SuccessUsers) {    
                [PSCustomObject]@{
                    'Username'                              = $UserAcct.UserID
                    'First Activity (UTC)'                  = ($AllSuccess | Where-Object {$_.UserID -like $UserAcct.UserID} | Sort-Object -Property "Date*" | Select-Object * -first 1)."DateUTC"
                    'Last Activity (UTC)'                   = ($AllSuccess | Where-Object {$_.UserID -like $UserAcct.UserID} | Sort-Object -Property "Date*" | Select-Object * -Last 1)."DateUTC"
                    'Total Connections'                     = ($AllSuccess | Where-Object {$_.UserID -like $UserAcct.UserID} | Measure-Object).count
                }
            }

            $OutPath = "$OutputLocation\IP Search - $date.xlsx"
            $Location | Export-ExcelDefault -Path $OutPath -WorkSheetName "IP"

            $SharePointOneDrive = $Location | Where-Object {$_.Workload -like "*OneDrive*" -or $_.Workload -like "*Sharepoint*"}
            if ($SharePointOneDrive) {
                $SharePointOneDrive | Export-ExcelDefault -Path $OutPath -WorkSheetName "OneDrive-Sharepoint-UAL"
            }

            if ($IPReport) {
                $IPReport   | Export-ExcelDefault -Path $OutPath -WorkSheetName "IP Report - Success"
            }
            if ($UserReport) {
                $UserReport | Export-ExcelDefault -Path $OutPath -WorkSheetName "User Report - Success"
            }

            Get-ItemHash -Default -FilePath $OutPath

            if ($DisplayOutput) {                
                $Location                     
            }
            if (!($path) -and (!($NoLaunch)) -and (!($IncludeMailboxLogs))) {
                Start-Process $OutputLocation
            }
        }
        Else {
            Write-Host "No detection of Unified Audit data." -ForegroundColor Green
        }        
    }    

    if ($ObjectIDs -and (!($Username))) {

        if ($IPAddress)  {
            do {
                $auditdata3 += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId $sessionid -ResultSize 5000 -StartDate $startdate -EndDate $enddate -ObjectIds $ObjectIDs -IPAddresses $ipaddress
            }while (($auditdata3 | Measure-Object).count % 5000 -eq 0 -and ($auditdata3 | Measure-Object).count -ne 0 -and $auditdata3)
            $auditdata2 = $auditdata3 | Select-Object * -ExcludeProperty resultindex,resultcount -Unique | Sort-Object -Property creationdate -Descending
        }
        if (!($IPAddress)) {
            do {
                $auditdata3 += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId $sessionid -ResultSize 5000 -StartDate $startdate -EndDate $enddate -ObjectIds $ObjectIDs
            }while (($auditdata3 | Measure-Object).count % 5000 -eq 0 -and ($auditdata3 | Measure-Object).count -ne 0 -and $auditdata3)
            $auditdata2 = $auditdata3 | Select-Object * -ExcludeProperty resultindex,resultcount -Unique | Sort-Object -Property creationdate -Descending
        }
        
        If ($auditdata2) {
            $auditdata1 = $auditdata2.auditdata
            $auditdata2 | ConvertTo-Json -Depth 100 | Out-File -FilePath "$OutputLocation\ObjectID Search - $date - RawData-UAL.json"

            Get-ItemHash -Default -FilePath "$OutputLocation\ObjectID Search - $date - RawData-UAL.json"

            $auditdata = $auditdata1 | 
                ForEach-Object -Process {
                try {
                    $_ | ConvertFrom-Json
                } 
                catch {
                    Write-Verbose "can't convert file '$_' to JSON"
                }
            }

            $Location = foreach ($i in $Auditdata) {

                $Udate = [datetime]$i.CreationTime
                $TimeData = $udate.tostring("MM/dd/yyyy hh:mm:ss tt")

                $UserID = $i.UserID.ToLower()

                $ClientIP1 = $null
                if ($i.ClientIP) {$ClientIP1 = $i.ClientIP}
                if ($i.ClientIPAddress -and (!($i.ClientIP))) {$ClientIP1 = $i.ClientIPAddress}
                if ($ClientIP1) {
                    $ClientIP = $ClientIP1.Replace('[','').Split(']')[0]
                    if ($Clientip -notmatch $regex) {$Clientip = $Clientip.split(':')[0]}
                    $IPLookup = Get-IPLookUp -IPAddress $ClientIP -OutputObject
                }

                if ($i.workload -like "OneDrive" -or $i.workload -like "*Sharepoint*") {
                    $Sharepoint = $i.ObjectId
                }
                Else { $Sharepoint = " " }

                [PSCustomObject]@{
                    'DateUTC'            =    $TimeData
                    'UserID'             =    $UserID.ToLower()
                    'ClientIP'           =    $ClientIP
                    'City'               =    $IPLookup.City
                    'State'              =    $IPLookup.State
                    'Country'            =    $IPLookup.Country
                    'ISP'                =    $IPLookup.ISP
                    'Operation'          =    $i.operation
                    'Workload'           =    $i.Workload
                    'ResultStatus'       =    $i.resultstatus
                    'LogonError'         =    $i.LogonError
                    'UserAgent'          =    ($i.extendedproperties | Where-Object {$_.name -like "UserAgent"}).Value
                    'ObjectID'           =    $Sharepoint
                }                
            }

            $OutPath = "$OutputLocation\ObjectID Search - $date.xlsx"
            
            $Location | Export-ExcelDefault -Path $OutPath -WorkSheetName "Objects"
            $SharePointOneDrive = $Location | Where-Object {$_.Workload -like "*OneDrive*" -or $_.Workload -like "*Sharepoint*"}
            if ($SharePointOneDrive) {
                $SharePointOneDrive | Export-ExcelDefault -Path $OutPath -WorkSheetName "OneDrive-Sharepoint-UAL"
            }

            Get-ItemHash -Default -FilePath $OutPath
            
            if ($DisplayOutput) {
                $Location
                   
            }
            if (!($path) -and (!($NoLaunch)) ) {
                Start-Process $OutputLocation
            }
        }
        Else {
            Write-Host "No detection of Unified Audit data." -ForegroundColor Green
        }    
    }

    if ($FailedSignIns) {

        do {
            $auditdata3 += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId $sessionid -ResultSize 5000 -StartDate $startdate -EndDate $enddate -Operations UserLoginFailed 
        }while (($auditdata3 | Measure-Object).count % 5000 -eq 0 -and ($auditdata3 | Measure-Object).count -ne 0 -and $auditdata3)
        $auditdata2 = $auditdata3 | Select-Object * -ExcludeProperty resultindex,resultcount -Unique | Sort-Object -Property creationdate -Descending

        If (!($auditdata2)) {
            Write-Host "No recent failed sign ins found." -ForegroundColor Green
        }
        Else {
            $auditdata1 = $auditdata2.auditdata
            $auditdata2 | ConvertTo-Json -Depth 100 | Out-File -FilePath "$OutputLocation\Failed Sign Ins - $date - RawData-UAL.json"

            Get-ItemHash -Default -FilePath "$OutputLocation\Failed Sign Ins - $date - RawData-UAL.json"

            $auditdata = $auditdata1 | 
                ForEach-Object -Process {
                try {
                    $_ | ConvertFrom-Json
                } 
                catch {
                    Write-Verbose "can't convert file '$_' to JSON"
                }
            }

            $Location = foreach ($i in $Auditdata) {

                $Udate = [datetime]$i.CreationTime
                $TimeData = $udate.tostring("MM/dd/yyyy hh:mm:ss tt")

                $UserID = $i.UserID.ToLower()

                $ClientIP1 = $null
                if ($i.ClientIP) {$ClientIP1 = $i.ClientIP}
                if ($i.ClientIPAddress -and (!($i.ClientIP))) {$ClientIP1 = $i.ClientIPAddress}
                if ($ClientIP1) {
                    $ClientIP = $ClientIP1.Replace('[','').Split(']')[0]
                    if ($Clientip -notmatch $regex) {$Clientip = $Clientip.split(':')[0]}
                    $IPLookup = Get-IPLookUp -IPAddress $ClientIP -OutputObject
                }

                [PSCustomObject]@{
                    'DateUTC'            =    $TimeData
                    'UserID'             =    $UserID
                    'ClientIP'           =    $ClientIP
                    'City'               =    $IPLookup.City
                    'State'              =    $IPLookup.State
                    'Country'            =    $IPLookup.Country
                    'ISP'                =    $IPLookup.ISP
                    'Operation'          =    $i.operation
                    'Workload'           =    $i.Workload
                    'ResultStatus'       =    $i.resultstatus
                    'LogonError'         =    $i.LogonError
                    'UserAgent'          =    ($i.extendedproperties | Where-Object {$_.name -like "UserAgent"}).Value
                }
            }

            $OutFailedPath = "$OutputLocation\Failed Sign Ins - $date.xlsx"
            
            $Location | Export-ExcelDefault -Path $OutFailedPath -WorkSheetName "Results"

            Get-ItemHash -Default -FilePath $OutFailedPath

            if ($DisplayOutput) {
                Write-Host 'Audit Log search results:' -ForegroundColor Green
                Write-Host 'Recent sign in locations via a web browser (City, State, Country, IP Address):'  -ForegroundColor Green
                $Location | Format-Table
            }

            if (!($path) -and (!($NoLaunch))) {
                Start-Process $OutputLocation
            }
        }
    }

    if ($FailedSignInStatistics) {
                     
        If (!((Get-AdminAuditLogConfig).UnifiedAuditLogIngestionEnabled)) {
            Write-Host "Unified Audit Logs status shows disabled. This may be incorrect reporting from Microsoft. Please verify status. Attempting search..." -ForegroundColor Green            
        }

        do {
            $auditdata3 += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId $sessionid -ResultSize 5000 -StartDate (get-date).date.AddDays(-8) -EndDate (get-date).date.AddDays(1) -Operations UserLoginFailed 
        }while (($auditdata3 | Measure-Object).count % 5000 -eq 0 -and ($auditdata3 | Measure-Object).count -ne 0 -and $auditdata3)
        $auditdata2 = $auditdata3 | Select-Object * -ExcludeProperty resultindex,resultcount -Unique | Sort-Object -Property creationdate -Descending

        if (!($Auditdata2)) {
            Write-Host "No recent failed sign ins found."
            break
        }

        $auditdata1 = $auditdata2.auditdata
        $auditdata2 | ConvertTo-Json -Depth 100 | Out-File -FilePath "$OutputLocation\Failed Sign In Statistics - $date - RawData-UAL.json"

        Get-ItemHash -Default -FilePath "$OutputLocation\Failed Sign In Statistics - $date - RawData-UAL.json"

        $auditdata = $auditdata1 | 
            ForEach-Object -Process {
            try {
                $_ | ConvertFrom-Json
            } 
            catch {
                Write-Verbose "can't convert file '$_' to JSON"
            }
        }

        $Location = foreach ($i in $Auditdata) {

            $Udate = [datetime]$i.CreationTime
            $TimeData = $udate.tostring("MM/dd/yyyy hh:mm:ss tt")

            $UserID = $i.UserID.ToLower()

            $ClientIP1 = $null
            if ($i.ClientIP) {$ClientIP1 = $i.ClientIP}
            if ($i.ClientIPAddress -and (!($i.ClientIP))) {$ClientIP1 = $i.ClientIPAddress}
            if ($ClientIP1) {
                $ClientIP = $ClientIP1.Replace('[','').Split(']')[0]
                if ($Clientip -notmatch $regex) {$Clientip = $Clientip.split(':')[0]}
                $IPLookup = Get-IPLookUp -IPAddress $ClientIP -OutputObject
            }

            [PSCustomObject]@{
                'DateUTC'            =    $TimeData
                'UserID'             =    $UserID.ToLower()
                'ClientIP'           =    $ClientIP
                'City'               =    $IPLookup.City
                'State'              =    $IPLookup.State
                'Country'            =    $IPLookup.Country
                'ISP'                =    $IPLookup.ISP
                'Operation'          =    $i.operation
                'Workload'           =    $i.Workload
                'ResultStatus'       =    $i.resultstatus
                'LogonError'         =    $i.LogonError
                'UserAgent'          =    ($i.extendedproperties | Where-Object {$_.name -like "UserAgent"}).Value
            }
        }     

        $ReportDate = (Get-Date).toString("MM/dd/yyyy - hh:mm tt")
        $TotalAttempts = ($Location).Count
        $MailboxCount = (Get-mailbox -ResultSize unlimited).count
        $FailedPerDay = $TotalAttempts / "7"  | ForEach-Object {$_.ToString("#.###")}
        $FailedPerMailbox = $FailedPerDay / $mailboxCount | ForEach-Object {$_.ToString("#.###")}

        $FailedStatistics = [PSCustomObject]@{
            'Date report was gathered:'               = $ReportDate
            'Failed Sign In Statistics (Past 7 days):' = " "
            ' ' = " "
            'Total Mailbox Count'                     = $MailboxCount
            'Total Sign In Attempts'                  = $TotalAttempts
            'Daily Average Per Mailbox'               = $FailedPerMailbox
            'Daily Average Attempts'                  = $FailedPerDay
        }
        if (!($location.Count)) { Write-Host Error - Less than two results found. Unable to display statistics. -ForegroundColor Green}
        if ($DisplayOutput) {
            Write-Host "7 Day Statistics:" -ForegroundColor Green
            $FailedStatistics.psobject.Properties | Select-Object Name, Value 
        }
        
        $FailedSignInStatisticsPath = "$OutputLocation\Failed Sign In Statistics - $date.xlsx"
        
        $FailedStatistics.psobject.Properties | Select-Object Name, Value | 
            Export-ExcelDefault -Path $FailedSignInStatisticsPath -WorkSheetName "Statistics"         
        $Location | Export-ExcelDefault -Path $FailedSignInStatisticsPath -WorkSheetName "Results"
        
        Get-ItemHash -Default -FilePath "$OutputLocation\Failed Sign In Statistics - $date.xlsx"

        if (!($path) -and (!($NoLaunch))) {
            Start-Process $OutputLocation
        }        
    }

    if ($Inputlog) {
        $auditdata2 = Import-Csv -Path $Inputlog
        $auditdata1 = $auditdata2.auditdata
        $auditdata2 | ConvertTo-Json -Depth 100 | Out-File -FilePath "$OutputLocation\Audit Log - InputCSV - $date - RawData-UAL.json"

        Get-ItemHash -Default -FilePath "$OutputLocation\Audit Log - InputCSV - $date - RawData-UAL.json"

        $auditdata = $auditdata1 | 
            ForEach-Object -Process {
            try {
                $_ | ConvertFrom-Json -ErrorAction SilentlyContinue
            } 
            catch {
                Write-Verbose "can't convert file '$_' to JSON"
            }
        }

        $Location = foreach ($i in $Auditdata) {

            $Udate = [datetime]$i.CreationTime
            $TimeData = $udate.tostring("MM/dd/yyyy hh:mm:ss tt")

            $UserID = $i.UserID.ToLower()

            $ClientIP1 = $null
            if ($i.ClientIP) {$ClientIP1 = $i.ClientIP}
            if ($i.ClientIPAddress -and (!($i.ClientIP))) {$ClientIP1 = $i.ClientIPAddress}
            if ($ClientIP1) {
                $ClientIP = $ClientIP1.Replace('[','').Split(']')[0]
                if ($Clientip -notmatch $regex) {$Clientip = $Clientip.split(':')[0]}
                $IPLookup = Get-IPLookUp -IPAddress $ClientIP -OutputObject
            }

            if ($i.workload -like "OneDrive" -or $i.workload -like "*Sharepoint*") {
                $Sharepoint = $i.ObjectId
            }
            Else { $Sharepoint = " " }


            [PSCustomObject]@{
                 'DateUTC'           =    $TimeData
                'UserID'             =    $UserID.ToLower()
                'ClientIP'           =    $ClientIP
                'City'               =    $IPLookup.City
                'State'              =    $IPLookup.State
                'Country'            =    $IPLookup.Country
                'ISP'                =    $IPLookup.ISP
                'Operation'          =    $i.operation
                'Workload'           =    $i.Workload
                'ResultStatus'       =    $i.resultstatus
                'LogonError'         =    $i.LogonError
                'UserAgent'          =    ($i.extendedproperties | Where-Object {$_.name -like "UserAgent"}).Value
                'ObjectID'           =    $Sharepoint
            }
        }

        If (!($location)) {
            Write-Host "No detection of recent sign in via a web browser" -ForegroundColor Green
        }
        Else {
            if (!(Test-Path "$OutputLocation\Audit Log - InputCSV - $date.xlsx")) {$ual = $true} Else {$ual = $false}
            if ($ual) {
                $Inputoutputpath = "$OutputLocation\Audit Log - InputCSV - $date.xlsx"
                $Location | Export-ExcelDefault -Path $Inputoutputpath -WorkSheetName "Unified Audit Logs"                

                $SharePointOneDrive = $Location | Where-Object {
                    $_.Workload -like "*OneDrive*" -or $_.Workload -like "*Sharepoint*"
                } 
                if ($SharePointOneDrive) {
                    $SharePointOneDrive | Export-ExcelDefault -Path $Inputoutputpath -WorkSheetName "OneDrive-Sharepoint-UAL"
                }
            }
            if ($DisplayOutput) {
                $Location
            }
        }

        Get-ItemHash -Default -FilePath "$OutputLocation\Audit Log - InputCSV - $date.xlsx"

        if (!($path) -and (!($NoLaunch))) {
            Start-Process $OutputLocation
        }
    }

    if ($InboxRules) {
        if ($Username) {
            do {
                $auditdata3 += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId $sessionid -ResultSize 5000 -StartDate $startdate -EndDate $enddate -operations 'New-InboxRule,Set-InboxRule,UpdateInboxRules' -UserIds "$username"
            }while (($auditdata3 | Measure-Object).count % 5000 -eq 0 -and ($auditdata3 | Measure-Object).count -ne 0 -and $auditdata3)
            $auditdata2 = $auditdata3 | Select-Object * -ExcludeProperty resultindex,resultcount -Unique | Sort-Object -Property creationdate -Descending
        }
        Else {
            do {
                $auditdata3 += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId $sessionid -ResultSize 5000 -StartDate $startdate -EndDate $enddate -operations 'New-InboxRule,Set-InboxRule,UpdateInboxRules'
            }while (($auditdata3 | Measure-Object).count % 5000 -eq 0 -and ($auditdata3 | Measure-Object).count -ne 0 -and $auditdata3)
            $auditdata2 = $auditdata3 | Select-Object * -ExcludeProperty resultindex,resultcount -Unique | Sort-Object -Property creationdate -Descending
        }

        If (!($auditdata2)) {
            Write-Host "No inbox rules found" -ForegroundColor Green
        }
        Else {
            $auditdata1 = $auditdata2.auditdata
            $auditdata2 | ConvertTo-Json -Depth 100 | Out-File -FilePath "$OutputLocation\Audit Log - Inbox Rules - $date - RawData-UAL.json"

            Get-ItemHash -Default -FilePath "$OutputLocation\Audit Log - Inbox Rules - $date - RawData-UAL.json"

            $auditdata = $auditdata1 | 
                ForEach-Object -Process {
                try {
                    $_ | ConvertFrom-Json
                } 
                catch {
                    Write-Verbose "can't convert file '$_' to JSON"
                }
            }

            $AuditDataRules = $AuditData | Where-Object {$_.UserId}
            $RuleInfo = foreach ($AuditEntry in $AuditDataRules) {
                $Udate = [datetime]$AuditEntry.CreationTime
                $TimeData = $udate.tostring("MM/dd/yyyy hh:mm:ss tt")

                $UserID = $AuditEntry.UserID.ToLower()

                $ClientIP1 = $null
                if ($AuditEntry.ClientIP) {$ClientIP1 = $AuditEntry.ClientIP}
                if ($AuditEntry.ClientIPAddress -and (!($AuditEntry.ClientIP))) {$ClientIP1 = $AuditEntry.ClientIPAddress}
                if ($ClientIP1) {
                    $ClientIP = $ClientIP1.Replace('[','').Split(']')[0]
                    if ($Clientip -notmatch $regex) {$Clientip = $Clientip.split(':')[0]}
                    $IPLookup = Get-IPLookUp -IPAddress $ClientIP -OutputObject
                }

                [PSCustomObject]@{
                    'Name'  = 'UserName'
                    'Value' = $AuditEntry.UserID
                }
                [PSCustomObject]@{
                    'Name'  = $time
                    'Value' = $TimeData
                }
                [PSCustomObject]@{
                    'Name'  = 'ClientIP'
                    'Value' = $ClientIP
                }
                [PSCustomObject]@{
                    'Name'  = 'City'
                    'Value' = $IPLookup.City
                }
                [PSCustomObject]@{
                    'Name'  = 'State'
                    'Value' = $IPLookup.State
                }
                [PSCustomObject]@{
                    'Name'  = 'Country'
                    'Value' = $IPLookup.Country
                }
                [PSCustomObject]@{
                    'Name'  = 'ISP'
                    'Value' = $IPLookup.ISP
                }                
                [PSCustomObject]@{
                    'Name'  = 'Operation'
                    'Value' = $AuditEntry.Operation
                }
                [PSCustomObject]@{
                    'Name'  = ' '
                    'Value' = ' '
                }
                $AuditEntry.parameters

                [PSCustomObject]@{
                    'Name'  = ' '
                    'Value' = ' '
                }
                [PSCustomObject]@{
                    'Name'  = ' '
                    'Value' = ' '
                }
            }

            $Location = foreach ($i in $Auditdata) {

                $Udate = [datetime]$i.CreationTime
                $TimeData = $udate.tostring("MM/dd/yyyy hh:mm:ss tt")

                $UserID = $i.UserID.ToLower()

                $ClientIP1 = $null
                if ($i.ClientIP) {$ClientIP1 = $i.ClientIP}
                if ($i.ClientIPAddress -and (!($i.ClientIP))) {$ClientIP1 = $i.ClientIPAddress}
                if ($ClientIP1) {
                    $ClientIP = $ClientIP1.Replace('[','').Split(']')[0]
                    if ($Clientip -notmatch $regex) {$Clientip = $Clientip.split(':')[0]}
                    $IPLookup = Get-IPLookUp -IPAddress $ClientIP -OutputObject
                }
                
                [PSCustomObject]@{
                    'DateUTC'            =    $TimeData
                    'UserID'             =    $UserID.ToLower()
                    'ClientIP'           =    $ClientIP
                    'City'               =    $IPLookup.City
                    'State'              =    $IPLookup.State
                    'Country'            =    $IPLookup.Country
                    'ISP'                =    $IPLookup.ISP
                    'Operation'          =    $i.operation
                    'Workload'           =    $i.Workload
                    'ResultStatus'       =    $i.resultstatus
                    'LogonError'         =    $i.LogonError
                    'UserAgent'          =    ($i.extendedproperties | Where-Object {$_.name -like "UserAgent"}).Value
                }
            }
            $OutPath = "$OutputLocation\Audit Log - Inbox Rules - $date.xlsx"
            $Location | Export-ExcelDefault -Path $OutPath -WorkSheetName "Results"
            $RuleInfo | Export-ExcelDefault -Path $OutPath -WorkSheetName "Rules"

            Get-ItemHash -Default -FilePath $OutPath    

            if ($DisplayOutput) {
                $Location
            }
            if (!($path) -and (!($NoLaunch))) {
                Start-Process $OutputLocation
            }
        }
    }

    if ($MailboxPermissions) {
        if ($Username) {
            do {
                $auditdata3 += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId $sessionid -ResultSize 5000 -StartDate $startdate -EndDate $enddate -Operations "Add-MailboxPermission,Remove-MailboxPermission" -UserIds "$username" 
            }while (($auditdata3 | Measure-Object).count % 5000 -eq 0 -and ($auditdata3 | Measure-Object).count -ne 0 -and $auditdata3)
        }
        Else {
            do {
                $auditdata3 += Search-unifiedAuditLog -SessionCommand ReturnLargeSet -SessionId $sessionid -ResultSize 5000 -StartDate $startdate -EndDate $enddate -Operations "Add-MailboxPermission,Remove-MailboxPermission"
            }while (($auditdata3 | Measure-Object).count % 5000 -eq 0 -and ($auditdata3 | Measure-Object).count -ne 0 -and $auditdata3)
        }
        if ($auditdata3) {

            $auditdata2 = $auditdata3 | Select-Object * -ExcludeProperty resultindex,resultcount -Unique | Sort-Object -Property creationdate -Descending
            $auditdata1 = $auditdata2.auditdata

            $auditdata2 | ConvertTo-Json -Depth 100 | Out-File -FilePath "$OutputLocation\MailboxPermissions Logs - $date - RawData-UAL.json"
            
            Get-ItemHash -Default -FilePath "$OutputLocation\MailboxPermissions Logs - $date - RawData-UAL.json"

            $auditdata = $auditdata1 | 
                ForEach-Object -Process {
                try {
                    $_ | ConvertFrom-Json -ErrorAction SilentlyContinue
                } 
                catch {
                    Write-Verbose "can't convert file '$_' to JSON"
                }
            }

            $Location = foreach ($i in $Auditdata) {

                $UserAccessing = $null; $UserGettingAccessed = $null; $AccessRights = $null; $UserAccessing1 = $null

                $Udate = [datetime]$i.CreationTime
                $TimeData = $udate.tostring("MM/dd/yyyy hh:mm:ss tt")

                $UserID = $i.UserID.ToLower()

                $ClientIP1 = $null
                if ($i.ClientIP) {$ClientIP1 = $i.ClientIP}
                if ($i.ClientIPAddress -and (!($i.ClientIP))) {$ClientIP1 = $i.ClientIPAddress}
                if ($ClientIP1) {
                    $ClientIP = $ClientIP1.Replace('[','').Split(']')[0]
                    if ($Clientip -notmatch $regex) {$Clientip = $Clientip.split(':')[0]}
                    $IPLookup = Get-IPLookUp -IPAddress $ClientIP -OutputObject
                }

                if ($i.Parameters) {
                    $UserGettingAccessed1  = ($i.Parameters | Where-Object {$_.name -like "Identity"}).Value
                    if ($UserGettingAccessed1) {
                        $UserGettingAccessed = (get-mailbox -Identity $UserGettingAccessed1 -ErrorAction silentlycontinue).UserPrincipalName
                    }
                    if ((!($UserGettingAccessed)) -and (($i.Parameters | Where-Object {$_.name -like "Identity"}).value -like "*/*")) {
                        $UserGettingAccessed2 = (($i.Parameters | Where-Object {$_.name -like "Identity"}).value -split "/")[-1]
                        $UserGettingAccessed = (get-mailbox -Identity $UserGettingAccessed2 -ErrorAction silentlycontinue).UserPrincipalName
                        if (!($UserGettingAccessed)) {$UserGettingAccessed = $UserGettingAccessed2 }
                        
                    }
                    if (!($UserGettingAccessed)) {$UserGettingAccessed = '' }

                    $AccessRights         = ($i.Parameters | Where-Object {$_.name -like "AccessRights"}).Value
                    if (!($AccessRights )) {$AccessRights = ''}

                    $UserAccessing1       = ($i.Parameters | Where-Object {$_.name -like "User"}).Value
                    if ($UserAccessing1) {
                        if ($UserAccessing1 -like "*Discovery Management*") {$UserAccessing = 'Discovery Management'}
                        Else {
                            $UserAccessing = (get-mailbox -Identity $UserAccessing1 -ErrorAction silentlycontinue).UserPrincipalName  
                        }
                    }
                    if ((!($UserAccessing)) -and (($i.Parameters | Where-Object {$_.name -like "User"}).value -like "*/*")) {
                        $UserAccessing2 = (($i.Parameters | Where-Object {$_.name -like "User"}).value -split "/")[-1]
                        $UserAccessing  = (get-mailbox -Identity $UserAccessing2 -ErrorAction silentlycontinue).UserPrincipalName
                        if (!($UserAccessing)) { $UserAccessing = $UserAccessing2}
                    }
                    if (!($UserAccessing)) { $UserAccessing = ''}  
                }

                [PSCustomObject]@{
                    'DateUTC'            =    $TimeData
                    'UserID'             =    $UserID.ToLower()
                    'ClientIP'           =    $ClientIP
                    'City'               =    $IPLookup.City
                    'State'              =    $IPLookup.State
                    'Country'            =    $IPLookup.Country
                    'ISP'                =    $IPLookup.ISP
                    'Operation'          =    $i.operation
                    'Workload'           =    $i.Workload
                    'ResultStatus'       =    $i.resultstatus                    
                    'UserAccessing'      =    $UserAccessing
                    'UserGettingAccessed'=    $UserGettingAccessed
                    'AccessRights'       =    $AccessRights
                    'UserAgent'          =    ($i.extendedproperties | Where-Object {$_.name -like "UserAgent"}).Value
                }              
            }
            $OutFile = "$OutputLocation\MailboxPermissions Logs - $date.xlsx"
            $Location   | Export-ExcelDefault -Path $OutFile -WorkSheetName 'Logs' -GetHash

            if ($DisplayOutput) {
                $Location
            }

            if (!($path) -and (!($NoLaunch))) {
                Start-Process $OutputLocation
            }
        }
    }
}