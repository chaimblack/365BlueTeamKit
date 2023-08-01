<#
.SYNOPSIS
Gets hash of file and exports hash as CSV.
 
.DESCRIPTION
Gets hash of file and exports hash as CSV.
 
.PARAMETER FilePath
Defines file path

.PARAMETER OutPath
Defines out path.

.PARAMETER DisplayOutput
If defined, will display output.

.PARAMETER NoLaunch
If defined, will not launch folder upon completion

.PARAMETER Algorithm
Defines Algorithm. If not defined, selects sha256

.PARAMETER SetSourceReadOnly
Sets source file to read-only

.PARAMETER SaveOriginalLocation
Saves output to source file location

.PARAMETER Default
Sets the following values to true: SetSourceReadOnly, SaveOriginalLocation, Nolaunch

.EXAMPLE
Get-ItemHash -FilePath 'C:\PSOutPut\Get-ItemHash\test.txt' -NoLaunch
Gets hash of file and does not launch

.Notes
Last updated by Chaim Black on 6/27/2021

This script is used when capturing audit logs to set the file as read-only and to
get the hash of the file and save it to a csv file.
#>

function Get-ItemHash {

    [CmdletBinding()]
    Param(
        [Parameter()]
        [string]$FilePath,
        [Parameter()]
        [string]$OutPath,
        [Parameter()]
        [switch]$DisplayOutput,
        [Parameter()]
        [switch]$SaveOriginalLocation,
        [Parameter()]
        [switch]$SetSourceReadOnly,
        [Parameter()]
        [switch]$NoLaunch,
        [ValidateSet("SHA1", "SHA256", "SHA384", "SHA512", "MACTripleDES", "MD5", "RIPEMD160")]
        [System.String]
        $Algorithm="SHA256",
        [Parameter()]
        [switch]$Default
    )

    if ($Default) {
        $NoLaunch             = $True
        $SetSourceReadOnly    = $True
        $SaveOriginalLocation = $True
    }

    if (!($FilePath)) { $FilePath = Read-Host "Enter Path"}

    $Verify = Get-ItemProperty -Path $FilePath -ErrorAction SilentlyContinue
    if (!($Verify)) {
        Write-Host "Error: Failed to locate source file." -ForegroundColor Red
    }

    if ($SetSourceReadOnly) {
        Set-ItemProperty -Path $FilePath -Name IsReadOnly -Value $True
    }

    if ($OutPath) {
        $OutputPath = $OutPath
    }
    Elseif ($SaveOriginalLocation) {$OutputPath = $Verify.Directory.FullName}
    else {    
        $OutputPath = "C:\PSOutput\Get-ItemHash"
        If (!(test-path $OutputPath)) {
            New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null
        }
    }

    $date = (get-date).toString("MM-dd-yy hh-mm-ss tt")

    $Hash = Get-FileHash -Path $FilePath -Algorithm $Algorithm

    $OutFile = $OutputPath + '\' + ($Hash.Path).Split('\')[-1] + ' - ' +$Date + '.hash.csv'

    if ($Hash) {
        $Hash | Export-csv -Path $OutFile
        Set-ItemProperty -Path $OutFile -Name IsReadOnly -Value $True

        if ($DisplayOutput) { $Hash }
        if (!($NoLaunch)) { Start-Process $OutputPath}
    }
}