<#
.SYNOPSIS
    Converts image files to a PDF document.
.DESCRIPTION
    This PowerShell script takes a folder of image files and converts them into a single PDF document.
    It supports various page sizes and orientations.
.PARAMETER ImageFolder
    The folder containing the image files to convert.
.PARAMETER OutputPath
    The path where the PDF file will be saved.
.PARAMETER PageSize
    The page size for the PDF (A4, Letter, Legal).
.PARAMETER Orientation
    The page orientation (Portrait or Landscape).
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$ImageFolder,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("A4", "Letter", "Legal")]
    [string]$PageSize = "A4",
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Portrait", "Landscape")]
    [string]$Orientation = "Portrait"
)

# Load necessary assemblies
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

# Check if the image folder exists
if (-not (Test-Path -Path $ImageFolder)) {
    Write-Error "Image folder does not exist: $ImageFolder"
    exit 1
}

# Get all image files in the folder, sorted by name
$imageFiles = Get-ChildItem -Path $ImageFolder -Filter *.png | Sort-Object Name

if ($imageFiles.Count -eq 0) {
    Write-Error "No image files found in the folder: $ImageFolder"
    exit 1
}

try {
    # Create a PDF document
    $doc = New-Object -TypeName System.Drawing.Printing.PrintDocument
    $doc.DocumentName = $OutputPath
    
    # Set page settings
    $pageSizeObj = switch ($PageSize) {
        "A4" { New-Object System.Drawing.Printing.PaperSize("A4", 827, 1169) }
        "Letter" { New-Object System.Drawing.Printing.PaperSize("Letter", 850, 1100) }
        "Legal" { New-Object System.Drawing.Printing.PaperSize("Legal", 850, 1400) }
        default { New-Object System.Drawing.Printing.PaperSize("A4", 827, 1169) }
    }
    
    $doc.DefaultPageSettings.PaperSize = $pageSizeObj
    $doc.DefaultPageSettings.Landscape = ($Orientation -eq "Landscape")
    
    # Create a PDF printer
    $printerSettings = New-Object System.Drawing.Printing.PrinterSettings
    $printerSettings.PrinterName = "Microsoft Print to PDF"
    $printerSettings.PrintToFile = $true
    $printerSettings.PrintFileName = $OutputPath
    $doc.PrinterSettings = $printerSettings
    
    # Current image index for the PrintPage event
    $script:currentImageIndex = 0
    
    # Handle the PrintPage event
    $printPageHandler = {
        param($sender, $e)
        
        if ($script:currentImageIndex -lt $imageFiles.Count) {
            $imagePath = $imageFiles[$script:currentImageIndex].FullName
            $image = [System.Drawing.Image]::FromFile($imagePath)
            
            # Calculate dimensions to fit the page while maintaining aspect ratio
            $pageWidth = $e.PageSettings.PaperSize.Width - $e.PageSettings.Margins.Left - $e.PageSettings.Margins.Right
            $pageHeight = $e.PageSettings.PaperSize.Height - $e.PageSettings.Margins.Top - $e.PageSettings.Margins.Bottom
            
            $ratio = [Math]::Min($pageWidth / $image.Width, $pageHeight / $image.Height)
            $newWidth = $image.Width * $ratio
            $newHeight = $image.Height * $ratio
            
            # Center the image on the page
            $x = ($pageWidth - $newWidth) / 2 + $e.PageSettings.Margins.Left
            $y = ($pageHeight - $newHeight) / 2 + $e.PageSettings.Margins.Top
            
            # Draw the image
            $e.Graphics.DrawImage($image, $x, $y, $newWidth, $newHeight)
            
            # Dispose the image
            $image.Dispose()
            
            # Increment the image index
            $script:currentImageIndex++
            
            # Indicate if there are more pages
            $e.HasMorePages = ($script:currentImageIndex -lt $imageFiles.Count)
        } else {
            $e.HasMorePages = $false
        }
    }
    
    # Add the event handler
    $doc.add_PrintPage($printPageHandler)
    
    # Print the document
    $doc.Print()
    
    # Clean up
    $doc.Dispose()
    
    Write-Output "PDF created successfully: $OutputPath"
    exit 0
} catch {
    Write-Error "Error creating PDF: $_"
    exit 1
}