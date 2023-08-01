<#
.SYNOPSIS
Uses Export-Excel to output with desired settings
 
.DESCRIPTION
Uses Export-Excel to output with desired settings

.PARAMETER Path
Defines path of file to save to.

.PARAMETER WorkSheetName
Defines worksheet name to use.

.PARAMETER Data
Defines data which is used in script. Accepts value from pipeline so it is not required to specify.

.PARAMETER GetHash
If defined, gets hash of outputted Excel file after completing and sets to read only.

.EXAMPLE
Get-Process | Export-ExcelDefault -Path C:\Process.xlsx -WorkSheetName Data
Gets processes, exports as Excel to C:\Process.xlsx with WorkSheetName of Data.

.Notes
Last updated by Chaim Black on 7/30/2021
#>
function Export-ExcelDefault {
    [CmdletBinding()]
    Param(        
        [Parameter(Mandatory=$True)]
        [string]$Path,
        [Parameter()]
        [string]$WorkSheetName,
        [Parameter(Mandatory=$True,ValuefromPipeline=$True)]
        [array]$Data,
        [Parameter()]
        [switch]$GetHash
    )
    Begin{        
        if (Test-Path -Path $Path) {
            Test-FileAccess -TestWrite -FilePath $Path
        }
        $output = @()
    }

    Process {
        $output += $data
    }

    End {
        if ($WorkSheetName) {
            $Output | Export-Excel -AutoSize -AutoFilter -NoNumberConversion '*' -FreezeTopRow -Path $Path -WorksheetName $WorkSheetName        
        }
        Else {
            $output | Export-Excel -AutoSize -AutoFilter -NoNumberConversion '*' -FreezeTopRow -Path $Path        
        }
        if ($GetHash) {
            Get-ItemHash -Default -FilePath $Path
        }
    }    
}