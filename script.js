// Global variables to store selected images
let selectedImages = [];

// DOM elements
const dropArea = document.getElementById('dropArea');
const fileInput = document.getElementById('fileInput');
const imageList = document.getElementById('imageList');
const noImagesText = document.querySelector('.no-images-text');
const convertBtn = document.getElementById('convertBtn');
const clearBtn = document.getElementById('clearBtn');
const statusMessage = document.getElementById('statusMessage');
const pdfNameInput = document.getElementById('pdfName');

// Event listeners for drag and drop functionality
dropArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropArea.classList.add('dragover');
});

dropArea.addEventListener('dragleave', () => {
    dropArea.classList.remove('dragover');
});

dropArea.addEventListener('drop', (e) => {
    e.preventDefault();
    dropArea.classList.remove('dragover');
    
    if (e.dataTransfer.files.length > 0) {
        handleFiles(e.dataTransfer.files);
    }
});

dropArea.addEventListener('click', () => {
    fileInput.click();
});

fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) {
        handleFiles(fileInput.files);
    }
});

// Handle the selected files
function handleFiles(files) {
    const imageFiles = Array.from(files).filter(file => file.type.startsWith('image/'));
    
    if (imageFiles.length === 0) {
        showStatus('Please select valid image files (JPG, PNG, GIF, etc.)', 'error');
        return;
    }
    
    // Add new images to the selectedImages array
    imageFiles.forEach(file => {
        // Check if the file is already in the array
        const isDuplicate = selectedImages.some(img => 
            img.name === file.name && 
            img.size === file.size && 
            img.lastModified === file.lastModified
        );
        
        if (!isDuplicate) {
            selectedImages.push(file);
        }
    });
    
    // Update the UI
    updateImagePreview();
    updateButtonStates();
    
    // Show success message
    showStatus(`${imageFiles.length} image(s) added successfully`, 'success');
}

// Update the image preview section
function updateImagePreview() {
    // Clear the current preview
    imageList.innerHTML = '';
    
    // Show/hide the "no images" text
    if (selectedImages.length === 0) {
        noImagesText.style.display = 'block';
        return;
    } else {
        noImagesText.style.display = 'none';
    }
    
    // Create preview elements for each image
    selectedImages.forEach((file, index) => {
        const imageItem = document.createElement('div');
        imageItem.className = 'image-item';
        imageItem.setAttribute('data-index', index);
        imageItem.draggable = true;
        
        // Create image preview
        const img = document.createElement('img');
        img.className = 'image-preview';
        img.src = URL.createObjectURL(file);
        img.alt = file.name;
        
        // Create image actions (remove, move)
        const actions = document.createElement('div');
        actions.className = 'image-actions';
        
        // File name display (truncated if too long)
        const fileName = document.createElement('span');
        fileName.className = 'file-name';
        fileName.textContent = file.name.length > 15 ? 
            file.name.substring(0, 12) + '...' : 
            file.name;
        fileName.title = file.name; // Show full name on hover
        
        // Remove button
        const removeBtn = document.createElement('button');
        removeBtn.className = 'remove-btn';
        removeBtn.innerHTML = '❌';
        removeBtn.title = 'Remove image';
        removeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            removeImage(index);
        });
        
        // Add elements to the DOM
        actions.appendChild(fileName);
        actions.appendChild(removeBtn);
        
        imageItem.appendChild(img);
        imageItem.appendChild(actions);
        imageList.appendChild(imageItem);
        
        // Add drag and drop reordering functionality
        setupDragAndDrop(imageItem);
    });
}

// Setup drag and drop for reordering images
function setupDragAndDrop(element) {
    element.addEventListener('dragstart', (e) => {
        e.dataTransfer.setData('text/plain', e.target.getAttribute('data-index'));
        element.classList.add('dragging');
    });
    
    element.addEventListener('dragend', () => {
        element.classList.remove('dragging');
    });
    
    element.addEventListener('dragover', (e) => {
        e.preventDefault();
        element.classList.add('drag-over');
    });
    
    element.addEventListener('dragleave', () => {
        element.classList.remove('drag-over');
    });
    
    element.addEventListener('drop', (e) => {
        e.preventDefault();
        element.classList.remove('drag-over');
        
        const sourceIndex = parseInt(e.dataTransfer.getData('text/plain'));
        const targetIndex = parseInt(element.getAttribute('data-index'));
        
        if (sourceIndex !== targetIndex) {
            // Reorder the images array
            const temp = selectedImages[sourceIndex];
            
            if (sourceIndex < targetIndex) {
                // Moving forward
                for (let i = sourceIndex; i < targetIndex; i++) {
                    selectedImages[i] = selectedImages[i + 1];
                }
            } else {
                // Moving backward
                for (let i = sourceIndex; i > targetIndex; i--) {
                    selectedImages[i] = selectedImages[i - 1];
                }
            }
            
            selectedImages[targetIndex] = temp;
            
            // Update the UI
            updateImagePreview();
        }
    });
}

// Remove an image from the selection
function removeImage(index) {
    selectedImages.splice(index, 1);
    updateImagePreview();
    updateButtonStates();
}

// Update the state of action buttons
function updateButtonStates() {
    convertBtn.disabled = selectedImages.length === 0;
    clearBtn.disabled = selectedImages.length === 0;
}

// Show status messages
function showStatus(message, type) {
    statusMessage.textContent = message;
    statusMessage.className = 'status-message';
    statusMessage.classList.add(type);
    
    // Auto-hide after 5 seconds
    setTimeout(() => {
        statusMessage.style.display = 'none';
    }, 5000);
}

// Clear all selected images
clearBtn.addEventListener('click', () => {
    selectedImages = [];
    updateImagePreview();
    updateButtonStates();
    showStatus('All images cleared', 'info');
});

// Convert images to PDF
convertBtn.addEventListener('click', async () => {
    if (selectedImages.length === 0) {
        showStatus('Please select at least one image', 'error');
        return;
    }
    
    // Show processing status
    showStatus('Processing images and creating PDF...', 'info');
    
    try {
        // Get PDF options
        const pdfName = pdfNameInput.value.trim() || 'converted_document';
        const pageSize = document.getElementById('pageSize').value;
        const orientation = document.getElementById('orientation').value;
        
        // Save images to temporary folder
        const tempFolder = await createTempFolder();
        await saveImagesToTemp(tempFolder);
        
        // Run PowerShell script to convert images to PDF
        const result = await convertToPdf(tempFolder, pdfName, pageSize, orientation);
        console.log('Convert to PDF result:', result);
        
        if (result.success) {
            // Create download link with the full path if available, otherwise just the name
            const pdfPath = result.path || pdfName;
            console.log('Using PDF path for download:', pdfPath);
            createDownloadLink(pdfPath);
            showStatus(`PDF created successfully. Click the download button to save ${pdfName}.pdf`, 'success');
        } else {
            showStatus('Failed to create PDF. Please try again.', 'error');
        }
    } catch (error) {
        console.error('Error creating PDF:', error);
        showStatus('An error occurred while creating the PDF', 'error');
    }
});

// Create a download link for the PDF
function createDownloadLink(pdfPath) {
    console.log('Creating download link for:', pdfPath);
    
    // Remove any existing download button
    const existingBtn = document.getElementById('downloadBtn');
    if (existingBtn) {
        existingBtn.remove();
        console.log('Removed existing download button');
    }
    
    // Create a new download button
    const downloadBtn = document.createElement('button');
    downloadBtn.id = 'downloadBtn';
    downloadBtn.className = 'primary-btn';
    downloadBtn.textContent = 'Download PDF';
    downloadBtn.addEventListener('click', async () => {
        try {
            // Request the PDF file from the server
            // Check if pdfPath already has .pdf extension
            const fullPdfPath = pdfPath.toLowerCase().endsWith('.pdf') ? pdfPath : `${pdfPath}.pdf`;
            console.log('Requesting PDF file:', fullPdfPath);
            
            const script = `
                $pdfPath = "${fullPdfPath}"
                Write-Host "Looking for PDF at: $pdfPath"
                if (Test-Path $pdfPath) {
                    $bytes = [System.IO.File]::ReadAllBytes($pdfPath)
                    $base64 = [Convert]::ToBase64String($bytes)
                    Write-Output $base64
                } else {
                    Write-Error "PDF file not found: $pdfPath"
                }
            `;
            
            const base64Data = await runPowerShellScript(script);
            
            // Create a download link
            const link = document.createElement('a');
            link.href = `data:application/pdf;base64,${base64Data}`;
            link.download = `${pdfName}.pdf`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            showStatus('PDF downloaded successfully', 'success');
        } catch (error) {
            console.error('Error downloading PDF:', error);
            showStatus('Failed to download PDF', 'error');
        }
    });
    
    // Add the download button to the action buttons container
    const actionButtons = document.querySelector('.action-buttons');
    actionButtons.appendChild(downloadBtn);
}

// Create a temporary folder for image processing
async function createTempFolder() {
    const folderName = 'temp_' + Date.now();
    const folderPath = `${folderName}`;
    
    // Create the folder using PowerShell
    const script = `
        $folderPath = "${folderPath}"
        if (-not (Test-Path $folderPath)) {
            New-Item -ItemType Directory -Path $folderPath | Out-Null
        }
        Write-Output $folderPath
    `;
    
    try {
        const result = await runPowerShellScript(script);
        return result.trim();
    } catch (error) {
        console.error('Error creating temp folder:', error);
        throw error;
    }
}

// Save selected images to the temporary folder
async function saveImagesToTemp(tempFolder) {
    for (let i = 0; i < selectedImages.length; i++) {
        const file = selectedImages[i];
        const reader = new FileReader();
        
        // Convert the file to base64
        const base64Data = await new Promise((resolve, reject) => {
            reader.onload = () => resolve(reader.result.split(',')[1]);
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
        
        // Use PowerShell to save the file
        const fileName = `image_${i.toString().padStart(3, '0')}.png`;
        const script = `
            $base64Data = "${base64Data}"
            $bytes = [Convert]::FromBase64String($base64Data)
            $filePath = Join-Path -Path "${tempFolder}" -ChildPath "${fileName}"
            [System.IO.File]::WriteAllBytes($filePath, $bytes)
            Write-Output "Saved ${fileName}"
        `;
        
        await runPowerShellScript(script);
    }
    
    return true;
}

// Convert the images to PDF using PowerShell
async function convertToPdf(tempFolder, pdfName, pageSize, orientation) {
    // PowerShell script to convert images to PDF
    const script = `
        # Load necessary assemblies
        Add-Type -AssemblyName System.Drawing
        Add-Type -AssemblyName System.Windows.Forms
        
        # Define paths
        $tempFolder = "${tempFolder}"
        $outputPdfPath = "${pdfName}.pdf"
        $pageSize = "${pageSize}"
        $orientation = "${orientation}"
        
        # Get all image files in the temp folder, sorted by name
        $imageFiles = Get-ChildItem -Path $tempFolder -Filter *.png | Sort-Object Name
        
        if ($imageFiles.Count -eq 0) {
            Write-Error "No images found in the temporary folder."
            exit 1
        }
        
        try {
            # Create a PDF document
            $doc = New-Object -TypeName System.Drawing.Printing.PrintDocument
            $doc.DocumentName = $outputPdfPath
            
            # Set page settings
            $pageSizeObj = switch ($pageSize) {
                "A4" { New-Object System.Drawing.Printing.PaperSize("A4", 827, 1169) }
                "Letter" { New-Object System.Drawing.Printing.PaperSize("Letter", 850, 1100) }
                "Legal" { New-Object System.Drawing.Printing.PaperSize("Legal", 850, 1400) }
                default { New-Object System.Drawing.Printing.PaperSize("A4", 827, 1169) }
            }
            
            $doc.DefaultPageSettings.PaperSize = $pageSizeObj
            $doc.DefaultPageSettings.Landscape = ($orientation -eq "landscape")
            
            # Create a PDF printer
            $printerSettings = New-Object System.Drawing.Printing.PrinterSettings
            $printerSettings.PrinterName = "Microsoft Print to PDF"
            $printerSettings.PrintToFile = $true
            $printerSettings.PrintFileName = $outputPdfPath
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
            
            # Clean up the temporary folder
            Remove-Item -Path $tempFolder -Recurse -Force
            
            # Get the full path of the PDF
            $fullPath = (Get-Item -Path $outputPdfPath).FullName
            
            # Return success with the file path
            $result = @{
                "success" = $true
                "message" = "PDF created successfully"
                "path" = $fullPath
            } | ConvertTo-Json
            
            Write-Output $result
        } catch {
            # Return error
            $result = @{
                "success" = $false
                "message" = "Error creating PDF: $_"
            } | ConvertTo-Json
            
            Write-Output $result
        }
    `;
    
    try {
        const resultStr = await runPowerShellScript(script);
        console.log('Raw PowerShell result:', resultStr);
        
        try {
            // Try to parse the JSON result
            const jsonResult = JSON.parse(resultStr);
            console.log('Parsed JSON result:', jsonResult);
            return jsonResult;
        } catch (parseError) {
            console.error('JSON parse error:', parseError);
            // If parsing fails, check if it contains success message
            if (resultStr.includes("PDF created successfully")) {
                // Extract the path if available
                const pathMatch = resultStr.match(/"path"\s*:\s*"([^"]+)"/i);
                const path = pathMatch ? pathMatch[1] : null;
                
                return { 
                    success: true, 
                    message: 'PDF created successfully',
                    path: path
                };
            } else {
                return { success: false, message: resultStr };
            }
        }
    } catch (error) {
        console.error('Error in PowerShell script:', error);
        return { success: false, message: error.toString() };
    }
}

// Function to run PowerShell scripts
async function runPowerShellScript(script) {
    return new Promise((resolve, reject) => {
        console.log('Executing PowerShell script:', script);
        
        // Create a request to the PowerShell API endpoint
        fetch('/api/powershell', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ script: script })
        })
        .then(response => {
            console.log('PowerShell response status:', response.status);
            return response.json();
        })
        .then(data => {
            console.log('PowerShell response data:', data);
            if (data.success) {
                resolve(data.output);
            } else {
                reject(new Error(data.error || 'PowerShell execution failed'));
            }
        })
        .catch(error => {
            console.error('Error executing PowerShell script:', error);
            reject(error);
        });
    });
}

// Initialize the UI
updateButtonStates();