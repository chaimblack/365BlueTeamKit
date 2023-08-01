<#
.SYNOPSIS
Gets geolocation information on IP Address.

.DESCRIPTION
Gets geolocation information on IP Address using an API from https://ipgeolocation.io

.PARAMETER IPAddress
Defines IP Address

.PARAMETER OutputObject
Outputs result as an object. Disabled by default.

.PARAMETER CachePath
Defines cache location. If not specified, set to C:\PSOutPut\Get-IPlookup

.EXAMPLE
Get-IPLookUp -IPAddress 131.107.0.89
Gets geolocation information on the 131.107.0.89 IP Address.

.EXAMPLE
Get-IPLookUp -IPAddress 131.107.0.89 -OutputObject
Gets geolocation information on the 131.107.0.89 IP Address. Outputs result in an object

.NOTES
Last updated: 9/3/2021 by Chaim Black.

API key required to use this script. With using this script for audit logs requiring ISP information, this script
uses the API from https://ipgeolocation.io. The API key should be placed on line 53 in this script within the parenthesis.
#>

function Get-IPLookUp {
    [CmdletBinding()]
    Param(
        [Parameter()]
        [string]$IPAddress,
        [Parameter()]
        [switch]$OutputObject,
        [Parameter()]
        [string]$CachePath,
        [Parameter()]
        [switch]$APIInformation
    )

    ##############################
    #Prerequisites
    ##############################

    #API key required to run IP lookup in audit logs.
    $APIKey = ""
    if (!($APIKey)) {Write-Host "Error: No API key defined. Please add to script on line 54." -ForegroundColor Red; Break}
    
    #Regex for searching in ipv4/ipv6 format:
    $regex = "((^\s*((([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]))\s*$)|(^\s*((([0-9A-Fa-f]{1,4}:){7}([0-9A-Fa-f]{1,4}|:))|(([0-9A-Fa-f]{1,4}:){6}(:[0-9A-Fa-f]{1,4}|((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3})|:))|(([0-9A-Fa-f]{1,4}:){5}(((:[0-9A-Fa-f]{1,4}){1,2})|:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3})|:))|(([0-9A-Fa-f]{1,4}:){4}(((:[0-9A-Fa-f]{1,4}){1,3})|((:[0-9A-Fa-f]{1,4})?:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){3}(((:[0-9A-Fa-f]{1,4}){1,4})|((:[0-9A-Fa-f]{1,4}){0,2}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){2}(((:[0-9A-Fa-f]{1,4}){1,5})|((:[0-9A-Fa-f]{1,4}){0,3}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(([0-9A-Fa-f]{1,4}:){1}(((:[0-9A-Fa-f]{1,4}){1,6})|((:[0-9A-Fa-f]{1,4}){0,4}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:))|(:(((:[0-9A-Fa-f]{1,4}){1,7})|((:[0-9A-Fa-f]{1,4}){0,5}:((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}))|:)))(%.+)?\s*$))"

    #Regex ipv4
    $ipv4Regex = "^(?:(?:0?0?\d|0?[1-9]\d|1\d\d|2[0-5][0-5]|2[0-4]\d)\.){3}(?:0?0?\d|0?[1-9]\d|1\d\d|2[0-5][0-5]|2[0-4]\d)$"

    #Regex for private IP Addresses
    $PrivRegex = "(^127\.)|(^10\.)|(^172\.1[6-9]\.)|(^172\.2[0-9]\.)|(^172\.3[0-1]\.)|(^192\.168\.)"

    $localipv6 = $IPAddress[0,1,2,3] -join""

    if ($IPAddress -match $ipv4Regex) {$IPType = 'ipv4'} Else {$IPType = 'ipv6'}

    #Set Cache Location
    if (!($CachePath)) {
        $CachePath = "C:\PSOutPut\Get-IPlookup\"    
    }
    
    If (!(Test-Path $CachePath)) {
        New-Item -ItemType Directory -Force -Path $CachePath | Out-Null
    }

    #Setup Cache
    $LookupTime = (Get-Date).toString("MM/dd/yyyy")

    If (test-path "$CachePath\IPCache.csv") {
        Test-FileAccess -TestWrite -FilePath "$CachePath\IPCache.csv"
        try {$CacheFile = Import-Csv -Path "$CachePath\IPCache.csv"}
        Catch {Start-Sleep -Seconds 1}
        if (!($CacheFile)) {
            try {$CacheFile = Import-Csv -Path "$CachePath\IPCache.csv"}
            Catch {Start-Sleep -Seconds 4}
        }
        #Remove old cache file
        if ($CacheFile[0].LookupTime -lt (Get-Date).AddDays(-5).toString("MM/dd/yyyy")){
            Remove-Item -Path "$CachePath\IPCache.csv" -Force
        }
        #Restrict cache file to 150 items. Added to reduce lookup time due to long times it takes to search cache. Example approx lookup times: 150 results = .09 seconds; 500 results = .29 seconds; 100 = .59 seconds
        if (($CacheFile | Measure-Object).count -gt "150"){
            Remove-Item -Path "$CachePath\IPCache.csv" -Force
        }
    }
    If (!(test-path "$CachePath\IPCache.csv")) {
        $newcache = [PSCustomObject]@{
            'IP'               = 'TBD'
            'City'             = 'TBD'
            'State'            = 'TBD'
            'Country'          = 'TBD'
            'ISP'              = 'TBD'
            'IPType'           = 'TBD'            
            'CSV'              = 'TBD'
            'LookupTime'       = $LookupTime
        }    

        $newcache | Export-Csv -Path "$CachePath\IPCache.csv"
    }

    ##############################
    #Script
    ##############################
 
    if (
        $IPAddress -match $PrivRegex -or `
        $IPAddress -like '::1' -or `
        $IPAddress -like '::' -or `
        $IPAddress -like "127.0.0.1" -or `
        $IPAddress -like "0.0.0.0" -or `
        $localipv6 -like "fe80"
    ) {
        $PrivIPadd = "Private"

        $Array = [PSCustomObject]@{
            'IP'               = $IPAddress
            'City'             = ''
            'State'            = ''
            'Country'          = ''
            'ISP'              = ''
            'IPType'           = $PrivIPadd            
            'CSV'              = $PrivIPadd
            'LookupTime'       = $LookupTime
        }
                
        if ($OutputObject) {
            $Array
        }
        Else {
            $PrivIPadd
        } 
    }

    if ($IPAddress -match $regex -and (!($PrivIPadd))) {
        Test-FileAccess -TestWrite -FilePath "$CachePath\IPCache.csv"
        Try {$IPCacheAll = Import-Csv -Path "$CachePath\IPCache.csv"}
        Catch {Start-Sleep -Seconds 1}
        if (!($IPCacheAll)) {
            Try {$IPCacheAll = Import-Csv -Path "$CachePath\IPCache.csv"}
            Catch {Start-Sleep -Seconds 4}
        }
        if ($IPAddress -in $IPCacheAll.ip) {
            $Output = ($IPCacheAll | Where-Object {$_.ip -like $IPAddress})[0]
            if ($OutputObject){
                $Output
            }
            Else {
                $Output.csv
            }
        }
        Else {
            $loc = "https://api.ipgeolocation.io/ipgeo?apiKey=" + "$APIKey"  + '&ip=' + $IPAddress
            try {
                $ip = Invoke-RestMethod -Method Get -Uri "$Loc"-ErrorAction SilentlyContinue -ErrorVariable IPError -WarningAction SilentlyContinue| Select-Object *
                if ($IPError -like "*The remote name could not be resolved*") {$NoInternet = $True}
            }
            Catch {
                Write-Verbose "Test Connection"
            }
            if ($IPError) {                
                $Array = [PSCustomObject]@{
                    'IP'               = $IPAddress
                    'City'             = "API-Error"
                    'State'            = "API-Error"
                    'Country'          = "API-Error"
                    'ISP'              = "API-Error"
                    'IPType'           = "API-Error"
                    'LookupTime'       = $LookupTime
                    'CSV'              = ($IPError).message
                    'Error'            = $ip.error
                }                

                if ($OutputObject){
                    $Array
                    #Write-Information $ip.error
                }
                Else {
                    $Array.csv
                    #Write-Information $ip.error
                }
            }
            elseif ($NoInternet) {
                $Array = [PSCustomObject]@{
                    'IP'               = $IPAddress
                    'City'             = "API-Error"
                    'State'            = "API-Error"
                    'Zip'              = "API-Error"
                    'Country'          = "API-Error"
                    'ISP'              = "API-Error"
                    'IPType'           = "API-Error"                                        
                    'LookupTime'       = $LookupTime
                    'CSV'              = "Error: No Internet"
                    'Error'            = "No Internet"
                } 

                if ($OutputObject){
                    $Array
                }
                Else {
                    $Array.csv
                }              
            }
            Else {
                
                $CSVArray = [PSCustomObject]@{                    
                    'City'             = if ($ip.City) { $ip.City } Else {'Unknown City'}
                    'State'            = if ($ip.state_prov) { $ip.state_prov } Else {'Unknown State'}
                    'Country'          = if ($ip.country_name) { $ip.country_name } Else {'Unknown Country'}
                    'ISP'              = if ($ip.isp) { $ip.isp } Else {'Unknown ISP'}
                }

                $Array = [PSCustomObject]@{
                    'IP'               = $ip.ip
                    'City'             = if ($ip.City) { $ip.City } Else {'Unknown City'}
                    'State'            = if ($ip.state_prov) { $ip.state_prov } Else {'Unknown State'}
                    'Country'          = if ($ip.country_name) { $ip.country_name } Else {'Unknown Country'}
                    'ISP'              = if ($ip.isp) { $ip.isp } Else {'Unknown ISP'}
                    'IPType'           = $IPType                                      
                    'LookupTime'       = $LookupTime
                    'CSV'              = $CSVArray.City + ", " + $CSVArray.state + ", " + $CSVArray.country + ", " + $CSVArray.isp
                }

                if ($OutputObject){
                    $Array
                }
                Else {
                    $CSVArray.City + ", " + $CSVArray.state + ", " + $CSVArray.country + ", " + $CSVArray.isp
                }

                if (!(Test-FileAccess -TestWrite -FilePath "$CachePath\IPCache.csv")) {
                    $imp = Import-Csv -Path "$CachePath\IPCache.csv"
                    $export = @()
                    $export += $imp
                    $export += $Array
                    try { $export | Select-Object * -Unique | Export-Csv -Path "$CachePath\IPCache.csv" }
                    Catch { Write-Verbose "Exporting results" }
                }
            }
        }
    }
}