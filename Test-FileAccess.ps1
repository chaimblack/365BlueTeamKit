<#
.SYNOPSIS
Tests file for read or write access.
 
.DESCRIPTION
Tests file for read or write access.
 
.PARAMETER FilePath
Defines file path

.PARAMETER DisplayOutput
Display output for success or failure. Default is False.

.PARAMETER DisplaySuccess
Display output for success. Default is False.

.PARAMETER HideError
Hide output if error. Default is false.

.PARAMETER TestRead
Tests read access to file.

.PARAMETER TestWrite
Tests write access to file.

.EXAMPLE
Test-FileAccess -FilePath 'C:\folder\testfile.txt' -TestWrite
Tests file to verify write access

.Notes
Last updated by Chaim Black on 7/27/2021
#>
function Test-FileAccess {
    [CmdletBinding()]
    Param(        
        [Parameter()]
        [string]$FilePath,
        [Parameter()]
        [switch]$DisplaySuccess,
        [Parameter()]
        [switch]$HideError,
        [Parameter()]
        [switch]$DisplayOutput,
        [Parameter()]
        [switch]$TestWrite,
        [Parameter()]
        [switch]$TestRead
    )

    if (!($FilePath)) {
        if ($DisplayOutput -or (!($HideError))) {
            [PSCustomObject]@{
                'Name'    = 'Test-FileAccess'
                'Error'   = $true
                'Path'    = $FilePath                
                'Message' = 'File path not defined.'
            }
        }
        Else {
            Write-Verbose 'File path not defined.'
        }
    }
    elseif ((!(Test-Path -Path $FilePath))) { 
        if ($DisplayOutput -or (!($HideError))) {
            [PSCustomObject]@{
                'Name'    = 'Test-FileAccess'
                'Error'   = $true
                'Path'    = $FilePath                
                'Message' = 'Failed to locate file for file lock verification.'
            }
        }
        Else {
            Write-Verbose 'Failed to locate file for file lock verification.'
        }
    }
    Else {

        if ($TestWrite) {
            $IsReadOnly = (Get-ItemProperty -Path $FilePath).IsReadOnly
            if ($IsReadOnly) {
                if ($DisplayOutput -or (!($HideError))) {
                    [PSCustomObject]@{
                        'Name'    = 'Test-FileAccess'
                        'Error'   = $true
                        'Path'    = $FilePath                        
                        'Message' = 'File lock verification failed. File is set to read only.'
                    }
                }
                Else {
                    Write-Verbose "File lock verification failed. File is set to read only."
                }
            }
            Else {
                $file = New-Object -TypeName System.IO.FileInfo -ArgumentList $FilePath
                $ErrorActionPreference = "SilentlyContinue"
                [System.IO.FileStream] $fs = $file.OpenWrite()
                if (!$?) {
                    $Failure = $true
                }
                else {
                    $fs.Dispose()
                    $Failure = $false
                    if ($DisplayOutput -or $DisplaySuccess) {
                        [PSCustomObject]@{
                            'Name'    = 'Test-FileAccess'
                            'Error'   = $False
                            'Path'    = $FilePath                            
                            'Message' = 'File is not locked and allows for write access.'
                        }
                    }
                    Else {
                        write-verbose 'File is not locked and allows for write access.'
                    }
                }

                if ($Failure) {
                    Start-Sleep -Seconds 3
                    $file = New-Object -TypeName System.IO.FileInfo -ArgumentList $FilePath
                    $ErrorActionPreference = "SilentlyContinue"
                    [System.IO.FileStream] $fs = $file.OpenWrite()
                    if (!$?) {
                        $Failure1 = $true
                    }
                    else {
                        $fs.Dispose()
                        $Failure1 = $false
                        if ($DisplayOutput -or $DisplaySuccess) {
                            [PSCustomObject]@{
                                'Name'    = 'Test-FileAccess'
                                'Error'   = $False
                                'Path'    = $FilePath
                                'Message' = 'File is not locked and allows for write access.'
                            }
                        }
                        Else {
                            write-verbose 'File is not locked and allows for write access.'
                        }
                    }

                    if ($Failure1) {
                        Start-Sleep -Seconds 4
                        $file = New-Object -TypeName System.IO.FileInfo -ArgumentList $FilePath
                        $ErrorActionPreference = "SilentlyContinue"
                        [System.IO.FileStream] $fs = $file.OpenWrite()
                        if (!$?) {
                            $Failure1 = $true
                            if ($DisplayOutput -or (!($HideError))) {
                                [PSCustomObject]@{
                                    'Name'    = 'Test-FileAccess'
                                    'Error'   = $True
                                    'Path'    = $FilePath                                    
                                    'Message' = 'File lock verification failed. File is unable to open with write access.'
                                }
                            }
                            Else {
                                Write-Verbose "File lock verification failed. File is unable to open with write access"
                            }
                        }
                        else {
                            $fs.Dispose()
                            $Failure1 = $false
                            if ($DisplayOutput -or $DisplaySuccess) {
                                [PSCustomObject]@{
                                    'Name'    = 'Test-FileAccess'
                                    'Error'   = $False
                                    'Path'    = $FilePath                                    
                                    'Message' = 'File is not locked and allows for write access.'
                                }
                            }
                            Else {
                                Write-Verbose 'File is not locked and allows for write access.'
                            }
                        }
                    }
                }
            }
        }

        if ($TestRead) {
            $file = New-Object -TypeName System.IO.FileInfo -ArgumentList $FilePath
            $ErrorActionPreference = "SilentlyContinue"
            [System.IO.FileStream] $fs = $file.OpenRead()
            if (!$?) {
                $Failure = $true
            }
            else {
                $fs.Dispose()
                $Failure = $false
                if ($DisplayOutput -or $DisplaySuccess) {
                    [PSCustomObject]@{
                        'Name'    = 'Test-FileAccess'
                        'Error'   = $False
                        'Path'    = $FilePath                        
                        'Message' = 'File is not locked and allows for read access.'
                    }
                }
                else {
                    Write-Verbose 'File is not locked and allows for read access.'
                }
            }

            if ($Failure) {
                Start-Sleep -Seconds 3
                $file = New-Object -TypeName System.IO.FileInfo -ArgumentList $FilePath
                $ErrorActionPreference = "SilentlyContinue"
                [System.IO.FileStream] $fs = $file.OpenRead()
                if (!$?) {
                    $Failure1 = $true
                }
                else {
                    $fs.Dispose()
                    $Failure1 = $false
                    if ($DisplayOutput -or $DisplaySuccess) {
                        [PSCustomObject]@{
                            'Name'    = 'Test-FileAccess'
                            'Error'   = $False
                            'Path'    = $FilePath                            
                            'Message' = 'File is not locked and allows for read access.'
                        }
                    }
                    else {
                        Write-Verbose 'File is not locked and allows for read access.'
                    }
                }

                if ($Failure1) {
                    Start-Sleep -Seconds 4
                    $file = New-Object -TypeName System.IO.FileInfo -ArgumentList $FilePath
                    $ErrorActionPreference = "SilentlyContinue"
                    [System.IO.FileStream] $fs = $file.OpenRead()
                    if (!$?) {
                        $Failure1 = $true
                        if ($DisplayOutput -or (!($HideError))) {
                            [PSCustomObject]@{
                                'Name'    = 'Test-FileAccess'
                                'Error'   = $True
                                'Path'    = $FilePath                                
                                'Message' = 'File lock verification failed. File is unable to open with read access.'
                            }
                        }                        
                        Else {
                            Write-Verbose "File lock verification failed. File is unable to open with read access."
                        }
                    }
                    else {
                        $fs.Dispose()
                        $Failure1 = $false
                        if ($DisplayOutput -or $DisplaySuccess) {
                            [PSCustomObject]@{
                                'Name'    = 'Test-FileAccess'
                                'Error'   = $False
                                'Path'    = $FilePath                                
                                'Message' = 'File is not locked and allows for read access.'
                            }
                        }
                        else {
                            Write-Verbose 'File is not locked and allows for read access.'
                        }
                    }
                }
            }
        }
    }
}