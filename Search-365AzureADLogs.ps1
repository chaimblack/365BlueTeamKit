<#
.SYNOPSIS
Parses data downloaded from an Azure AD login report json file.

.DESCRIPTION
Parses data downloaded from an Azure AD login report json file. Future plans on downloading directly from Microsoft once searching is not in AzureAD Preview.

.PARAMETER InputLogs
Specifies the json file to input.

.PARAMETER Lookup
Performs geolookup using "Get-IPlookup" script (API required). If not specified, uses Microsoft data (no ISP). 

.PARAMETER Path
If specified, defines output path

.PARAMETER NoLaunch
If defined, does not launch folder after completion

.PARAMETER DisplayOutput
Displays output.

.PARAMETER NoCopyJSON
By default, this script saves the raw data in json format to the output location and captures a hash. If selecting this switch, does not save this data.

.PARAMETER StartDate
Defines start date. If not defined, will search entire log file. Requires EndDate to also be specified.

.PARAMETER EndDate
Defines end date.  If not defined, will search entire log file. Requires StartDate to also be specified.

.PARAMETER Days
Defines how many days ago the start date should be. If not defined, will search entire log file.

.EXAMPLE
Search-365AzureADLogs -InputLogs 'C:\data.json'
Parses data downloaded from an Azure AD login report json file.

.NOTES
Last updated 9/13/2021 Chaim Black. Select UTC in "Show Dates" when downloading data.
#>

function Search-365AzureADLogs {
    [CmdletBinding()]
    Param(        
        [Parameter(Mandatory=$true)]
        [string]$InputLogs,
        [Parameter()]
        [string]$Path,
        [Parameter()]
        [switch]$NoLaunch,
        [Parameter()]
        [switch]$Lookup,
        [Parameter()]
        [switch]$DisplayOutput,
        [Parameter()]
        [string]$IPAddress,
        [Parameter()]
        [string]$Username,
        [Parameter()]
        [switch]$NoCopyJSON,
        [Parameter()]
        [DateTime]$StartDate,
        [Parameter()]
        [DateTime]$EndDate,
        [Parameter()]
        [int]$Days
    )

    If (!(Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
        Write-Host "Missing prerequisites. Installing now..."
        Install-Module ImportExcel -Force -AllowClobber

        If ($?) {Write-Host "Reporting module installed successfully." ; Import-Module ImportExcel}
        else {Write-Host "Reporting module failed to install." -ForegroundColor Red ; return }
    }
    Else {Import-Module ImportExcel}

    if (!($path)) {
        if ((Get-MsolCompanyInformation -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).DisplayName) {
            $CompanyName = (Get-MsolCompanyInformation -ErrorAction SilentlyContinue -WarningAction SilentlyContinue).DisplayName.trim()
            $SaveLocation = "C:\PSOutput\Search-365AzureADLogs\$CompanyName\"
        }
        Else {
            $SaveLocation = "C:\PSOutput\Search-365AzureADLogs\"
        }
        If (!(test-path $SaveLocation)) {
            New-Item -ItemType Directory -Force -Path $SaveLocation | Out-Null
        }
    }
    Else {
        $SaveLocation = $Path
    }
    $date = (get-date).toString("MM-dd-yy hh-mm-ss tt")

    if ($Days -or $StartDate -or $EndDate) {
        if ($StartDate -and (!($EndDate))) { Write-Host "Error: Requires EndDate when specifying StartDate"}
        if ($EndDate -and (!($StartDate))) { Write-Host "Error: Requires StartDate when specifying EndDate"}        

        if ($Days){
            $StartDate = ((Get-Date).AddDays(-$Days)).tostring("MM-dd-yyyy")
            $EndDate = (get-date).AddDays(1).toString("MM-dd-yyyy")
        }
    }

    if (!($StartDate)) {[DateTime]$StartDate = [DateTime]"1/1/1900"}
    if (!($EndDate))   {[DateTime]$EndDate   = [DateTime]"1/1/2300"}

    if ($InputLogs -notlike "*.json") {
        Write-host "Failed to locate json file." -ForegroundColor Red
        Break
    }

    if ($InputLogs -like "*,*") {
        $AllLogs = $InputLogs -split ','         
        foreach ($Log in $AllLogs) {
            $InputRawJSON       = get-content $Log
            $inputData          = $InputRawJSON | ConvertFrom-Json
            $UnfilteredData     = $UnfilteredData + $inputData
        }
    }
    Else {
        $InputRawJSON   = get-content $InputLogs
        $UnfilteredData = $InputRawJSON | ConvertFrom-Json   
    }

    $FullLogs  = foreach ($Entry in $UnfilteredData) {
        $CARule   = $null; $CARuleFailure  = $null; $AuthMethod = $null; $AppPassword = $null; $LegacyCAReport = $Null
        
        $Udate    = [datetime]$Entry.createdDateTime
        $TimeData = $udate.ToUniversalTime().tostring("MM/dd/yyyy hh:mm:ss tt")
        
        if ($Udate -lt [datetime]$StartDate )                          {Continue}
        if ($Udate -gt [datetime]$EndDate.AddDays(1))                  {Continue}
        if ($UserName) {
            $Username1 = $Username -split ','
            if ($Entry.userPrincipalName -notin $Username1)            {Continue}

        }
        if ($IPAddress) {
            $IPAddresses1 = $IPAddress -split ','
            if ($Entry.ipAddress -notin $IPAddresses1)                 {Continue}
        }        
        
        if ($entry.clientAppUsed -like "Mobile Apps and Desktop clients" -or $entry.clientAppUsed -like "Browser") {$AuthMethod = "ModernAuth"}
        Else {$AuthMethod = "LegacyAuth"}
        
        if ($entry.authenticationDetails.authenticationStepResultDetail -like "MFA requirement skipped due to app password") {$AppPassword = 'True'}
        
        $LegacyCAReport = (($Entry | Where-Object {$_.appliedConditionalAccessPolicies}).appliedConditionalAccessPolicies | Where-Object {$_.displayname -like "*Legacy Auth*"}).result

        if ($Entry.appliedConditionalAccessPolicies) {
            $AllCA = foreach ($Policy in $Entry.appliedConditionalAccessPolicies) {
                $Name    = $Policy.displayName
                $Control = $Policy.enforcedGrantControls
                $Result  = $Policy.result

                if ($Policy.result -like 'failure' ) { $Failure = 'True'}
                Else {$Failure = 'False'}

                $CARule1  = "Policy: " + $Name + "; Control: " + $Control + "; Result: " + $Result

                [PSCustomObject]@{
                    'Rule'    = $CARule1
                    'Failure' = $Failure
                }
            }
            $CARule        = $AllCA.rule -join '; '
            $CARuleFailure = ($AllCA | Where-Object {$_.Failure -like 'True'}).rule -join '; '
        }
        Else {
            $CARule         = $null
            $CARuleFailure  = $null
        }        
        
        if ($Lookup) {
            $IPLookup = Get-IPlookup -OutputObject -IPAddress $Entry.ipAddress

            $City                               = $IPLookup.City
            $State                              = $IPLookup.State
            $Country                            = $IPLookup.Country
            $ISP                                = $IPLookup.ISP   
        }
        Else {
            $City                               = $entry.location.City
            $State                              = $Entry.location.state
            $Country                            = $Entry.location.countryOrRegion
            $ISP                                = ''
        }

        [PSCustomObject]@{
            'Date(UTC)'                          = $TimeData
            'UserPrincipalName'                  = $Entry.userPrincipalName
            'IPAddress'                          = $Entry.ipAddress
            'City'                               = $City
            'State'                              = $State
            'Country'                            = $Country
            'ISP'                                = $ISP
            'AppDisplayName'                     = $Entry.appDisplayName
            'ClientAPP'                          = $Entry.clientAppUsed
            'ResultStatus'                       = $Entry.authenticationDetails.succeeded -join '; '
            'AuthMethod'                         = $AuthMethod
            'AppPassword'                        = $AppPassword
            'LegacyAuthCAReport'                 = $LegacyCAReport
            'UserAgent'                          = $Entry.userAgent
            'AuthenticationRequirement'          = $Entry.authenticationRequirement
            'FailureReason'                      = $Entry.status.failureReason  
            'LoginAdditionalDetails'             = $Entry.status.additionalDetails
            'OperatingSystem'                    = $Entry.deviceDetail.operatingSystem
            'Browser'                            = $Entry.deviceDetail.browser                
            'AuthenticationStepResultDetail'     = $Entry.authenticationDetails.authenticationStepResultDetail
            'ID'                                 = $Entry.ID
            'UserDisplayName'                    = $Entry.userDisplayName
            'UserId'                             = $Entry.userId
            'AppId'                              = $Entry.appId                
            'CorrelationId'                      = $Entry.correlationId
            'ConditionalAccessStatus'            = $Entry.conditionalAccessStatus
            'OriginalRequestId'                  = $Entry.originalRequestId
            'IsInteractive'                      = $Entry.isInteractive
            'TokenIssuerName'                    = $Entry.tokenIssuerName
            'TokenIssuerType'                    = $Entry.tokenIssuerType
            'ProcessingTimeInMilliseconds'       = $Entry.processingTimeInMilliseconds
            'RiskDetail'                         = $Entry.riskDetail
            'RiskLevelAggregated'                = $Entry.riskLevelAggregated
            'RiskLevelDuringSignIn'              = $Entry.riskLevelDuringSignIn
            'RiskState'                          = $Entry.riskState
            'RiskEventTypes'                     = $Entry.riskEventTypes -join '; '
            'RiskEventTypes_v2'                  = $Entry.riskEventTypes_v2  -join '; '
            'ResourceDisplayName'                = $Entry.resourceDisplayName
            'ResourceId'                         = $Entry.resourceId
            'AuthenticationMethodsUsed'          = $Entry.authenticationMethodsUsed  -join '; '
            'AlternateSignInName'                = $Entry.alternateSignInName
            'ServicePrincipalName'               = $Entry.servicePrincipalName
            'ServicePrincipalId'                 = $Entry.servicePrincipalId
            'MFADetail_authMethod'               = $Entry.mfaDetail.authMethod
            'MFADetail_authDetail'               = $Entry.mfaDetail.authDetail
            'Status_ErrorCode'                   = $Entry.status.ErrorCode
            'DeviceDetail_deviceId'              = $Entry.deviceDetail.deviceId
            'DeviceDetail_displayName'           = $Entry.deviceDetail.displayName
            'DeviceDetail_isCompliant'           = $Entry.deviceDetail.isCompliant
            'DeviceDetail_isManaged'             = $Entry.deviceDetail.isManaged
            'DeviceDetail_trustType'             = $Entry.deviceDetail.trustType
            'Location_altitude'                  = $entry.altitude
            'Location_latitude'                  = $entry.latitude
            'Location_longitude'                 = $entry.longitude
            'AppliedConditionalAccessPolicies'   = $CARule
            'AppliedCAFailures'                  = $CARuleFailure
            'AuthenticationProcessingDetails_key'                    = $Entry.authenticationProcessingDetails.Key -join '; '
            'AuthenticationProcessingDetails_value'                  = $Entry.authenticationProcessingDetails.value -join '; '
            'NetworkLocationDetails'                                 = $Entry.networkLocationDetails -join '; '
            'AuthenticationDetails_AuthenticationStepDateTime'       = $Entry.authenticationDetails.authenticationStepDateTime -join '; '
            'AuthenticationDetails_AuthenticationMethod'             = $Entry.authenticationDetails.authenticationMethod -join '; '
            'AuthenticationDetails_AuthenticationMethodDetail'       = $Entry.authenticationDetails.authenticationMethodDetail  -join '; '
            'AuthenticationDetails_AuthenticationStepRequirement'    = $Entry.authenticationDetails.authenticationStepRequirement -join '; '
            'AuthenticationRequirementPolicies'                      = $Entry.authenticationRequirementPolicies  -join '; '
            'JSON'                                                   = $Entry
        }
    }

    if ($FullLogs) {

        $FilteredData = $FullLogs | Select-Object * -Unique | Sort-Object -Property 'Date(UTC)' -Descending
     
        if (!($NoCopyJSON)) {
            $OutputRawJSONSorted = $FilteredData.json  | ConvertTo-Json -Depth 100
            $OutJsonFile = "$SaveLocation\Azure Logs - $date - RawDataAAD.json"
            $OutputRawJSONSorted | Set-Content -Path $OutJsonFile
            Get-ItemHash -Default -FilePath $OutJsonFile
        }

        $SaveFile = "$SaveLocation\Azure Logs - $date.xlsx"
        $FullLogs | Select-Object * -ExcludeProperty json | Export-ExcelDefault -path $SaveFile -WorksheetName "Azure AD Logs" -GetHash

        if (!($path) -and (!($NoLaunch))) {
            Start-Process $SaveLocation
        }

        if ($DisplayOutput) {$FullLogs}
    }
}