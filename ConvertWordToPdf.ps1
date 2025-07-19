<#
.SYNOPSIS
    Converts Word documents to PDF format.
.DESCRIPTION
    This PowerShell script takes a Word document and converts it to PDF format using the Word COM object.
    It supports both .doc and .docx file formats.
.PARAMETER WordFilePath
    The path to the Word document to convert.
.PARAMETER OutputPath
    The path where the PDF file will be saved.
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$WordFilePath,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath
)

# Set error action preference
$ErrorActionPreference = "Stop"

# Check if the Word file exists
if (-not (Test-Path -Path $WordFilePath)) {
    Write-Error "Word file does not exist: $WordFilePath"
    exit 1
}

try {
    # Load Word COM object
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    
    # Open the document
    Write-Output "Opening Word document: $WordFilePath"
    $doc = $word.Documents.Open($WordFilePath)
    
    # Save as PDF
    Write-Output "Converting to PDF: $OutputPath"
    $wdFormatPDF = 17  # PDF format constant
    $doc.SaveAs([ref]$OutputPath, [ref]$wdFormatPDF)
    
    # Close document and Word application
    $doc.Close()
    $word.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Write-Output "PDF created successfully: $OutputPath"
    exit 0
} catch {
    Write-Error "Error creating PDF: $_"
    exit 1
}