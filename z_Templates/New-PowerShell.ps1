<#
  .SYNOPSIS
  Once sentence what this script does.

  .DESCRIPTION
  More verbose explanation, use cases and scenarios when this can be used.

  .PARAMETER InputPath
  Specifies the path to the CSV-based input file.

  .PARAMETER OutputPath
  Add one for each parameter this script will accept.
  What is it used for, what data format is expected?

  .INPUTS
  Does the script accept pipe input?

  .OUTPUTS
  Does the script generate pipe output?

  .EXAMPLE
  PS> .\New-PowerShell.ps1

  .EXAMPLE
  PS> .\New-PowerShell.ps1 -ExampleParameterWithDefault C:\Data\January.csv

  .EXAMPLE
  PS> .\New-PowerShell.ps1 -ExampleParameterWithDefault C:\Data\January.csv -BeVerbose -VerboseLogFileName c:\temp\Mylog.txt

  .NOTES
  MIT License

  Copyright (c) 2021 TeamsAdminSamples

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  SOFTWARE.

#>
[CmdletBinding()] #few examples how a script accepts input parameters
param
(
    [Parameter(Mandatory=$false)]
    [string]$ExampleParameterWithDefault='.\MyInputFile.csv',
    [Parameter(Mandatory=$false)]
    [switch]$BeVerbose,
    [Parameter(Mandatory=$false)]
    [string] $VerboseLogFileName = 'VerboseOutput.Log'
)


function Write-VerboseLog{
    <#
        .SYNOPSIS
        Writes messages to Verbose output and Verbose log file
        .DESCRIPTION
        This fuction will direct verbose output to the console if the -Verbose 
        paramerter is specified.  It will also write output to the Verboselog.
        .EXAMPLE
        Write-VerboseLog -Message 'This is a single verbose message'
        .EXAMPLE
        # Write-VerboseLog -Message ('This is a more complex version where we want to include the value of a variable: {0}' -f ($MyVariable | Out-String) )
    #>
    param(
      [String]$Message  
    )
  
    $VerboseMessage = ('{0} Line:{1} {2}' -f (Get-Date), $MyInvocation.ScriptLineNumber, $Message)
    Write-Verbose -Message $VerboseMessage
    Add-Content -Path $VerboseLogFileName -Value $VerboseMessage
  }

  function Write-ScriptError{
    <#
        .SYNOPSIS
        Writes error messages to log file and terminates script if needed.
        .DESCRIPTION
        This function will write error message with detail of script location to the log. 
        Script will be terminated if the -Terminating $true parameter is passed.
        .EXAMPLE
        Write-ScriptError 'This is a terminating error. Script will not continue' -Terminating $true
        .EXAMPLE
        Write-ScriptError 'This is a not-terminating error. Script will continue in execution'

    #>
    param(
      [Parameter(Mandatory=$true,HelpMessage='Provide message to include in the log')]
      [string] $Message,
      
      [Parameter(Mandatory=$false,HelpMessage='Error object to report')]
      [Object]$ErrorObject,
  
      [bool] $Terminating = $false
    )
    
    Write-VerboseLog -Message $Message
    Write-VerboseLog -Message ('Error occurred: {0}' -f $ErrorObject.Exception.Message)
    Write-VerboseLog -Message ('STACKTRACE:{0}' -f $ErrorObject.ScriptStackTrace)
    if ($Terminating){Write-Error -Message $Message -ErrorAction Stop}
  }


