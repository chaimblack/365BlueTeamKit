<#
.SYNOPSIS
Gets DNS record information for an Office 365 tenant.

.DESCRIPTION
Gets DNS record information for an Office 365 tenant. This script looks at SPF, DKIM and DMARC records. It can be used with or without a connection to Office 365. 

.PARAMETER DomainName
If defined, will not require connection to Office 365 and this defines the domain['s] the script will look at.

.PARAMETER Path
Defines the output path. Default saves to C:\PSOutPut\Get-365DNSRecords

.PARAMETER NoLaunch
Prevents Windows Explorer from launching after search.

.PARAMETER OutObject
Displays results in output as object.

.PARAMETER DisplayErrorReport
If found, displays errors in report format.

.EXAMPLE
Get-365DNSRecords

.NOTES
Created by Chaim Black. Last updated 11/10/2021.
Geared to running with an active connection to Office 365, this script gets DNS information for a tenant and will produce an error report based off of common issues.
This does not require a connection to Office 365 if a domain is specified, but it will lack some information.

Required modules:
    Import-Excel
#>
function Get-365DNSRecords {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [string]$DomainName,
        [Parameter()]
        [string]$Path,
        [Parameter()]
        [switch]$NoLaunch,
        [Parameter()]
        [switch]$OutObject,
        [Parameter()]
        [switch]$DisplayErrorReport
    )

    #Check for Excel Reporting Module; Install if not found
    If (!(Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
         Write-Host "Missing prerequisites. Installing now..."
        Install-Module ImportExcel -Force -AllowClobber
        If ($?) {Write-Host "Reporting module installed successfully." ; Import-Module ImportExcel}
        else {Write-Host "Reporting module failed to install." -ForegroundColor Red ; return }
    }
    Else{Import-Module ImportExcel}

    $date = (Get-Date).ToString('MM-dd-yyyy-hh-mm-ss tt')
    if (!($Path)) {
        if (Get-MsolCompanyInformation -ErrorAction SilentlyContinue) {
            $Company = ((Get-MsolCompanyInformation).displayname).trim()
            $OutputLocation = "C:\PSOutPut\Get-365DNSRecords\$Company"
        }
        Else { $OutputLocation = "C:\PSOutPut\Get-365DNSRecords\" }
        If (!(test-path $OutputLocation)) {
            New-Item -ItemType Directory -Force -Path $OutputLocation | Out-Null
        }
    } 
    Else {
        $OutputLocation = $Path
    }

    $OutFile = "$OutputLocation\DNS Report - " + $Date + ".xlsx"

    if ($DomainName) {
        if ($DomainName -like '*') {
            $Domains = $DomainName -split ','
        }
        Else {$Domains = $DomainName}
    }
    Else {
        if (!(Get-Command 'Get-DKIMSigningConfig' -ErrorAction SilentlyContinue)) {
            Write-Host 'Not connected to Exchange Online and MSOnline. Please connect.' -ForegroundColor Red
            Break
        }
        $Domains = (Get-MsolDomain | Where-Object {$_.status -Like "Verified"}).name
    }

    # SPF
    $SPFResults = foreach ($Domain in $Domains) {
        $DNS = $null

        $DNS = Resolve-DnsName -Name $Domain -ErrorAction SilentlyContinue -Type TXT -Server 8.8.8.8 | Where-Object {$_.Strings -Match "spf1"}
        if ($DNS.strings) { $Record = [string]$DNS.strings}
        Else {$Record = 'None found' }

        if (($DNS | Measure-Object).count -gt '1') {$SPFNotes = 'Multiple DNS records found. Please remove duplicates.' }
        elseif ($DNS.strings -match '~') {$SPFNotes = "Review and verify record is correct and consider switching the '~all' to a hard-fail of '-all'"}
        elseif(!($DNS.strings)) {$SPFNotes = 'No SPF record found. Verify if domain sends outgoing mail and apply record if needed.'}
        Else {$SPFNotes = ''}

        if ($DNS.strings) {$Status = 'Enabled'}
        Else {$Status = 'Disabled'}

        [PSCustomObject]@{ 
            'DomainName'  = $Domain
            'Status'      = $Status
            'Record'      = $Record
            'Notes'       = $SPFNotes
        }
    }

    #DKIM
    $DKIMResults = foreach ($Domain in $Domains) {
        $DKIMOUtput = $null; $DKIMDomain1 = $null; $DKIMDomain1V = $null; $DKIMDomain2 = $null; $DKIMDomain2V = $null
        $Selector1Verify = $null; $Selector2Verify = $null

        $DKIMDomain1  = 'SELECTOR1._DOMAINKEY.' + $Domain
        $DKIMDomain1V = Resolve-DnsName -Type TXT -name $DKIMDomain1 -Server 8.8.8.8 -ErrorAction silentlycontinue

        $DKIMDomain2  = 'SELECTOR2._DOMAINKEY.' + $Domain
        $DKIMDomain2V = Resolve-DnsName -Type TXT -name $DKIMDomain2 -Server 8.8.8.8 -ErrorAction silentlycontinue

        if ((!($DomainName)) -and (Get-Command 'Get-DKIMSigningConfig' -ErrorAction SilentlyContinue)) {
            $DKIMOUtput   = Get-DKIMSigningConfig -Identity $Domain -ErrorAction silentlycontinue            
        
            if ($DKIMDomain1V.strings -join '' -and $DKIMDomain1V.strings -join '' -like $DKIMOUtput.Selector1PublicKey -join '') {$Selector1Verify = $True}
            Else {$Selector1Verify = $False}

            if ($DKIMDomain2V.strings -join '' -and $DKIMDomain2V.strings -join '' -like $DKIMOUtput.Selector2PublicKey -join '') {$Selector2Verify = $True}
            Else {$Selector2Verify = $False}

            if ($Selector1Verify -or $Selector2Verify) {$SelectorsVerified = $True }
            Else                                       {$SelectorsVerified = $False}

            if (!($DKIMOUtput)) {
                $DKIMOUtput = [PSCustomObject]@{
                    'Domain' = $Domain
                    'Enabled' = $false
               }
            }

            if (!($DKIMOUtput.enabled)) {
                $Status    = "Disabled"
                $DKIMNotes = "DKIM not enabled. Verify if domain sends outgoing mail and apply record if needed."
            }
            elseif ($DKIMOUtput.enabled -and $SelectorsVerified) {
                $Status    = "Enabled"
                $DKIMNotes = "DKIM enabled and DNS records verified."
            }
            elseif ($DKIMOUtput.enabled -and (!($SelectorsVerified))){
                $Status    = "Enabled. DNS verification failed."
                $DKIMNotes = "DKIM enabled but DNS records are not setup properly. Verify if domain sends outgoing mail and apply record if needed."
            }
            Else   {$Status = "Disabled"}
        
            [PSCustomObject]@{
                'DomainName'        = $Domain
                'status'            = $Status
                'Verified'          = $SelectorsVerified
                'Selector1Verified' = $Selector1Verify
                'Selector2Verified' = $Selector2Verify
                'Selector1-365'     = $DKIMOUtput.Selector1PublicKey -join ''
                'Selector1-DNS'     = $DKIMDomain1V.strings -join ''
                'Selector2-365'     = $DKIMOUtput.Selector2PublicKey -join ''
                'Selector2-DNS'     = $DKIMDomain2V.strings -join ''
                'Notes'             = $DKIMNotes
            }
        }
        Else {

            [PSCustomObject]@{
                'DomainName'        = $Domain
                'Selector1-DNS'     = $DKIMDomain1V.strings -join ''
                'Selector2-DNS'     = $DKIMDomain2V.strings -join ''
                'Notes'             = 'Unable to verify DKIM status. Not connected to Office 365.'
            }

        }
    }

    #DMARC
    $DMARCResults = foreach ($Domain in $Domains) {
        $DMARCDomain1 = $null; $DMARCDomain1V = $null

        $DMARCDomain1  = '_dmarc.' + $Domain
        $DMARCDomain1V = Resolve-DnsName -Type TXT -name $DMARCDomain1 -Server 8.8.8.8 -ErrorAction silentlycontinue

        if (($DMARCDomain1V).strings) { 
            $DMarcEnabled = 'Enabled'
            $DMarcNotes   = ""
            $Record       = $DMARCDomain1V.Strings -join ''
        }
        Else { 
            $DMarcEnabled = 'Disabled'
            $DMarcNotes   = "DMARC not configured. Verify if domain sends outgoing mail and apply record if needed."
            $Record       = ''
        }

        [PSCustomObject]@{
            'DomainName'   = $Domain
            'Status'       = $DMarcEnabled
            'Record'       = $Record
            'Notes'        = $DMarcNotes
        }
    }

    
    $SPFReport = foreach ($SPFResult in $SPFResults) {
        if (!($SPFResult.Notes)) {Continue}
        [PSCustomObject]@{
            'Name'  = 'DomainName'
            'Value' = $SPFResult.DomainName                
        }
        [PSCustomObject]@{
            'Name'  = 'RecordType'
            'Value' = 'SPF'                
        }
        [PSCustomObject]@{
            'Name'  = 'Status'
            'Value' = $SPFResult.Status
        }
        [PSCustomObject]@{
            'Name'  = 'Record'
            'Value' = $SPFResult.Record
        }
        [PSCustomObject]@{
            'Name'  = 'Notes'
            'Value' = $SPFResult.Notes
        }
        [PSCustomObject]@{
            'Name'  = ''
            'Value' = ''
        }
    }

    $DKIMReport = foreach ($DKIMResult in $DKIMResults) {
        if (!($DKIMResult.Notes)) {Continue}
        [PSCustomObject]@{
            'Name'  = 'DomainName'
            'Value' = $DKIMResult.DomainName                
        }
        [PSCustomObject]@{
            'Name'  = 'RecordType'
            'Value' = 'DKIM'                
        }
        [PSCustomObject]@{
            'Name'  = 'Status'
            'Value' = $DKIMResult.Status
        }
        [PSCustomObject]@{
            'Name'  = 'Record1'
            'Value' = $DKIMResult.'Selector1-DNS'
        }
        [PSCustomObject]@{
            'Name'  = 'Record2'
            'Value' = $DKIMResult.'Selector2-DNS'
        }
        [PSCustomObject]@{
            'Name'  = 'Notes'
            'Value' = $DKIMResult.Notes
        }
        [PSCustomObject]@{
            'Name'  = ''
            'Value' = ''
        }
    }

    $DmarcReport = foreach ($DMARCResult in $DMARCResults) {
        if (!($DMARCResult.Notes)) {Continue}
        [PSCustomObject]@{
            'Name'  = 'DomainName'
            'Value' = $DMARCResult.DomainName                
        }
        [PSCustomObject]@{
            'Name'  = 'RecordType'
            'Value' = 'DMARC'                
        }
        [PSCustomObject]@{
            'Name'  = 'Status'
            'Value' = $DMARCResult.Status
        }
        [PSCustomObject]@{
            'Name'  = 'Record'
            'Value' = $DMARCResult.Record
        }
        [PSCustomObject]@{
            'Name'  = 'Notes'
            'Value' = $DMARCResult.Notes
        }
        [PSCustomObject]@{
            'Name'  = ''
            'Value' = ''
        }
    }

    $ErrorReport = $SPFReport + $DKIMReport + $DmarcReport    

    #Export
    $SPFResults |
        Export-ExcelDefault -Path $OutFile -WorkSheetName 'SPF'

    $DKIMResults |
        Export-ExcelDefault -Path $OutFile -WorkSheetName 'DKIM'

    $DMARCResults |
        Export-ExcelDefault -Path $OutFile -WorkSheetName 'DMARC'

    $ErrorReport |
        Export-ExcelDefault -Path $OutFile -WorkSheetName 'ErrorReport'

    if (!($NoLaunch)) {Start-Process $OutputLocation}

    if ($OutObject) {
        [PSCustomObject]@{
            'SPF'         = $SPFResults
            'DKIM'        = $DKIMResults
            'DMARC'       = $DMARCResults
            'ErrorReport' = $ErrorReport
        }
    }

    if ($DisplayErrorReport) {
        $ErrorReport
    }
}