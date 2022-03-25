<#

.SYNOPSIS 
Takes input string and appends to a .log file, in a specified or dynamically chosen location.

.DESCRIPTION 
This function takes a string input and appends it to a .log file. If $LogFileName exists in the session that calls the function, then that variable is used as the location for the log file. Otherwise, the function will automatically choose the location. It does so based on the name of the script that called this function. Based on that, it chooses "".\logs\<script name>_yyyy-MM-ddTHH-MM-SS.log". Each line that is output to the log file is formatted with this format: "[M/dd/yyyy HH:mm:SS AM] - <input string>".

.PARAMETER String
Required. This is the input string; the text that you'd like to write to the log.

.PARAMETER CallerInvocation
Optional. Use this only if you are calling write-log from a script or function other than your main script AND you would like this function to automatically determine the best suitable log file name on it's own.

.PARAMETER LogFileFullName
Optional. Use this if you'd like to specify the name of the log file specifically. This is helpful if you want to specify different log files for different parts of your script.

.EXAMPLE
Write-Log "The script is starting"

.EXAMPLE
Write-Log "Im writing this log from a shared function script" $MyInvocation

.EXAMPLE
Write-Log "$($MBX.DisplayName) has $ItemCount Items in their $FolderName folder"
Write-Log "This is from the OnPrem MBX part of my script. I want it to go to a different log file than everyting else." -LogFileFullName "c:\temp\onpremusers_$date.log"
Write-Log "This log entry is inside of a function that is used in a script" -CallerInvocation $MyInvocation

#>
Function Write-Log {
    [CmdletBinding(SupportsShouldProcess = $True)]
	Param (
    [parameter(Position=0,Mandatory=$True)]
    [string]$String,

    [parameter(Position=1)]
    $CallerInvocation = $MyInvocation,

    [parameter(Position=2)]
    [string]$LogFileFullName
    
    )
    If($Verbose -eq $False) {
        $VerbosePreference = $Global:VerbosePreference
    }
    If($WhatIf -eq $False) {
        $WhatIfPreference = $Global:WhatIfPreference
    }

    Write-Verbose ("Here are the variables that are not null in $($MyInvocation.InvocationName): " + ($PSBoundParameters | Out-String))

    If($LogFileFullName) {
        $LogFileName = $LogFileFullName
    }
    #Set LogFileName if it's not listed in the script
    ElseIf(!$LogfileName)
        {
        Write-Verbose "`$LogFileName is null"
        If($CallerInvocation.ScriptName)
            {
            [string]$ScriptPath = split-path $CallerInvocation.ScriptName -Parent
            [string]$ScriptName = split-path $CallerInvocation.ScriptName -leaf
            [string]$LogFilePath = $ScriptPath + "\logs\"
            If(!(Test-Path $LogFilePath)){New-Item -ItemType Directory -Path $LogFilePath | Out-Null}
            [string]$LogFileDate = (get-date -f s).replace(":","-")
            [string]$Global:LogFileName = $LogFilePath + $ScriptName.split('.')[0] + "_$LogFileDate.log"
            Write-Verbose "Setting `$LogFileName to $($LogFilePath + $ScriptName.split('.')[0] + "_$LogFileDate.log")"
            }
            Else{$Global:LogFileName = Read-Host "`$LogFileName doesn't exist. Please input full log file name to continue..."}
        }
	# Get the current date

    
    [string]$date = Get-Date -Format G
    
	# Write everything to our log file
    If($LogfileName) {
        Write-Verbose "Writing string to $LogFileName"
        ( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFileName -Append
    }
    Else{Write-Error "`$LogFileName is not found, so no log can be written to..."}

}