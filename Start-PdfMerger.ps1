<#
.SYNOPSIS
    Launcher script for the Image to PDF Converter application.
.DESCRIPTION
    This script starts a simple HTTP server to serve the web application and handles
    the PowerShell integration for converting images to PDF.
#>

# Set error action preference
$ErrorActionPreference = "Stop"

# Define the port for the HTTP server
$port = 8080

# Get the directory of this script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Function to check if a port is in use
function Test-PortInUse {
    param (
        [int]$Port
    )
    
    $connections = netstat -ano | Select-String -Pattern "\s+TCP\s+.*:$Port\s+.*LISTENING"
    return $connections.Count -gt 0
}

# Find an available port if the default is in use
while (Test-PortInUse -Port $port) {
    $port++
}

# Create a simple HTTP server
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")

try {
    # Start the listener
    $listener.Start()
    
    Write-Host "Image to PDF Converter is running at http://localhost:$port/"
    Write-Host "Press Ctrl+C to stop the server"
    
    # Open the browser
    Start-Process "http://localhost:$port/"
    
    # Handle requests
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        # Get the requested URL
        $requestUrl = $request.Url.LocalPath
        
        # Handle API requests
        if ($requestUrl -eq "/api/powershell" -and $request.HttpMethod -eq "POST") {
            # Read the request body
            $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
            $requestBody = $reader.ReadToEnd()
            $reader.Close()
            
            # Parse the JSON request
            $requestData = $requestBody | ConvertFrom-Json
            
            # Execute the PowerShell script
            try {
                # Create a temporary file to store the script
                $tempScriptPath = [System.IO.Path]::GetTempFileName() + ".ps1"
                Set-Content -Path $tempScriptPath -Value $requestData.script
                
                # Execute the script and capture the output
                $scriptOutput = & powershell.exe -ExecutionPolicy Bypass -File $tempScriptPath 2>&1
                
                # Log the output for debugging
                Write-Host "PowerShell script output: $scriptOutput"
                
                # Prepare the response
                $responseData = @{
                    "success" = $true
                    "output" = $scriptOutput
                } | ConvertTo-Json -Depth 10
            } catch {
                # Prepare the error response
                $responseData = @{
                    "success" = $false
                    "error" = $_.Exception.Message
                } | ConvertTo-Json
            } finally {
                # Clean up the temporary file
                if (Test-Path $tempScriptPath) {
                    Remove-Item -Path $tempScriptPath -Force
                }
            }
            
            # Send the response
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseData)
            $response.ContentLength64 = $buffer.Length
            $response.ContentType = "application/json"
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
            $response.OutputStream.Close()
        }
        # Serve static files
        else {
            # Map the URL to a file path
            $filePath = if ($requestUrl -eq "/") {
                Join-Path -Path $scriptDir -ChildPath "index.html"
            } else {
                Join-Path -Path $scriptDir -ChildPath $requestUrl.TrimStart("/")
            }
            
            # Check if the file exists
            if (Test-Path -Path $filePath -PathType Leaf) {
                # Determine the content type
                $contentType = switch ([System.IO.Path]::GetExtension($filePath)) {
                    ".html" { "text/html" }
                    ".css" { "text/css" }
                    ".js" { "application/javascript" }
                    ".svg" { "image/svg+xml" }
                    ".png" { "image/png" }
                    ".jpg" { "image/jpeg" }
                    ".gif" { "image/gif" }
                    default { "application/octet-stream" }
                }
                
                # Read the file content
                $fileContent = [System.IO.File]::ReadAllBytes($filePath)
                
                # Send the response
                $response.ContentLength64 = $fileContent.Length
                $response.ContentType = $contentType
                $response.OutputStream.Write($fileContent, 0, $fileContent.Length)
                $response.OutputStream.Close()
            } else {
                # File not found
                $response.StatusCode = 404
                $response.Close()
            }
        }
    }
} catch {
    Write-Host "Error: $_"
} finally {
    # Stop the listener
    if ($listener -ne $null) {
        $listener.Stop()
    }
}