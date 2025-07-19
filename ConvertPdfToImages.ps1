<#
.SYNOPSIS
    Converts PDF documents to image files.
.DESCRIPTION
    This PowerShell script takes a PDF document and converts it to image files using GhostScript.
    It supports extracting all pages or specific page ranges and different image formats and quality settings.
.PARAMETER PdfFilePath
    The path to the PDF document to convert.
.PARAMETER OutputFolder
    The folder where the image files will be saved.
.PARAMETER Format
    The image format to use (png or jpg).
.PARAMETER Quality
    The image quality (high, medium, low).
.PARAMETER PageRange
    Optional. Specific pages to extract (e.g., "1-3,5,7-9"). If not specified, all pages will be extracted.
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$PdfFilePath,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputFolder,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("png", "jpg")]
    [string]$Format = "png",
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("high", "medium", "low")]
    [string]$Quality = "medium",
    
    [Parameter(Mandatory=$false)]
    [string]$PageRange = ""
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Check if the PDF file exists
if (-not (Test-Path -Path $PdfFilePath)) {
    Write-Error "PDF file does not exist: $PdfFilePath"
    exit 1
}

# Create output directory if it doesn't exist
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

try {
    # Try to find GhostScript in common locations
    $ghostscriptPath = $null
    $possiblePaths = @(
        "C:\Program Files\gs\gs*\bin\gswin64c.exe",
        "C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe"
    )
    
    foreach ($path in $possiblePaths) {
        $foundPaths = Get-Item -Path $path -ErrorAction SilentlyContinue
        if ($foundPaths) {
            $ghostscriptPath = $foundPaths[0].FullName
            break
        }
    }
    
    if (-not $ghostscriptPath) {
        Write-Error "GhostScript not found. Please install GhostScript to convert PDF to images."
        exit 1
    }
    
    Write-Output "Using GhostScript at: $ghostscriptPath"
    
    # Set quality parameter based on selected quality
    $qualityParam = switch ($Quality) {
        "high" { "-r300" }
        "medium" { "-r150" }
        "low" { "-r72" }
        default { "-r150" }
    }
    
    # Process page range if specified
    if ($PageRange) {
        Write-Output "Processing specific pages: $PageRange"
        
        # Parse page range (e.g., "1-3,5,7-9")
        $pageRanges = $PageRange -split ','
        $pageList = @()
        
        foreach ($range in $pageRanges) {
            if ($range -match '-') {
                $start, $end = $range -split '-'
                $start = [int]$start
                $end = [int]$end
                
                for ($i = $start; $i -le $end; $i++) {
                    $pageList += $i
                }
            } else {
                $pageList += [int]$range
            }
        }
        
        $pageList = $pageList | Sort-Object -Unique
        
        # Extract specific pages
        foreach ($pageNum in $pageList) {
            $outputFile = Join-Path -Path $OutputFolder -ChildPath "page_${pageNum}.${Format}"
            
            Write-Output "Converting page $pageNum to $outputFile"
            
            # Use GhostScript to convert PDF page to image
            $arguments = @(
                "-dNOPAUSE",
                "-dBATCH",
                "-dSAFER",
                $qualityParam,
                "-dFirstPage=${pageNum}",
                "-dLastPage=${pageNum}",
                "-sDEVICE=$(if ($Format -eq 'jpg') { 'jpeg' } else { 'png16m' })",
                "-sOutputFile=${outputFile}",
                $PdfFilePath
            )
            
            Start-Process -FilePath $ghostscriptPath -ArgumentList $arguments -Wait -NoNewWindow
        }
    } else {
        Write-Output "Converting all pages to images"
        
        # Extract all pages
        $outputPattern = Join-Path -Path $OutputFolder -ChildPath "page_%03d.${Format}"
        
        # Use GhostScript to convert all PDF pages to images
        $arguments = @(
            "-dNOPAUSE",
            "-dBATCH",
            "-dSAFER",
            $qualityParam,
            "-sDEVICE=$(if ($Format -eq 'jpg') { 'jpeg' } else { 'png16m' })",
            "-sOutputFile=${outputPattern}",
            $PdfFilePath
        )
        
        Start-Process -FilePath $ghostscriptPath -ArgumentList $arguments -Wait -NoNewWindow
    }
    
    # Count the number of images created
    $imageCount = (Get-ChildItem -Path $OutputFolder -Filter "page_*.${Format}" | Measure-Object).Count
    
    Write-Output "Successfully converted PDF to $imageCount images in $OutputFolder"
    exit 0
} catch {
    Write-Error "Error converting PDF to images: $_"
    exit 1
}