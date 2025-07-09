<#
.SYNOPSIS
    Bridge script to execute PowerShell commands from the web application.
.DESCRIPTION
    This script serves as a bridge between the web application and PowerShell.
    It receives commands from the web application and executes them, returning the results.
#>

# Parameters for the script
param (
    [Parameter(Mandatory=$true)]
    [string]$Command
)

# Set error action preference
$ErrorActionPreference = "Stop"

try {
    # Create a temporary file to store the command
    $tempScriptPath = [System.IO.Path]::GetTempFileName() + ".ps1"
    Set-Content -Path $tempScriptPath -Value $Command
    
    # Execute the command and capture the output
    $output = & powershell.exe -ExecutionPolicy Bypass -File $tempScriptPath
    
    # Return the output as JSON
    $result = @{
        "success" = $true
        "output" = $output
    } | ConvertTo-Json
    
    Write-Output $result
} catch {
    # Return the error as JSON
    $result = @{
        "success" = $false
        "error" = $_.Exception.Message
    } | ConvertTo-Json
    
    Write-Output $result
} finally {
    # Clean up the temporary file
    if (Test-Path $tempScriptPath) {
        Remove-Item -Path $tempScriptPath -Force
    }
}