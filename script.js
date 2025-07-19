// Global variables to store selected files
let selectedImages = [];
let selectedWordDoc = null;
let selectedPdf = null;

// DOM elements
// Tab elements
const tabButtons = document.querySelectorAll('.tab-btn');
const tabContents = document.querySelectorAll('.tab-content');

// Image to PDF elements
const dropArea = document.getElementById('dropArea');
const fileInput = document.getElementById('fileInput');
const imageList = document.getElementById('imageList');
const noImagesText = document.querySelector('.no-images-text');
const convertBtn = document.getElementById('convertBtn');
const clearBtn = document.getElementById('clearBtn');
const statusMessage = document.getElementById('statusMessage');
const pdfNameInput = document.getElementById('pdfName');

// Word to PDF elements
const wordDropArea = document.getElementById('wordDropArea');
const wordFileInput = document.getElementById('wordFileInput');
const docInfo = document.getElementById('docInfo');
const docName = document.getElementById('docName');
const noDocText = document.querySelector('.no-doc-text');
const wordConvertBtn = document.getElementById('wordConvertBtn');
const wordClearBtn = document.getElementById('wordClearBtn');
const wordPdfNameInput = document.getElementById('wordPdfName');
const wordDownloadBtn = document.getElementById('wordDownloadBtn');
const wordViewPdfBtn = document.getElementById('wordViewPdfBtn');
const removeDocBtn = document.getElementById('removeDocBtn');

// PDF to Image elements
const pdfDropArea = document.getElementById('pdfDropArea');
const pdfFileInput = document.getElementById('pdfFileInput');
const pdfInfo = document.getElementById('pdfInfo');
const pdfFileName = document.getElementById('pdfFileName');
const noPdfText = document.querySelector('.no-pdf-text');
const pdfConvertBtn = document.getElementById('pdfConvertBtn');
const pdfClearBtn = document.getElementById('pdfClearBtn');
const extractPages = document.getElementById('extractPages');
const pageRangeOption = document.getElementById('pageRangeOption');
const pageRange = document.getElementById('pageRange');
const extractedImages = document.getElementById('extractedImages');
const imageGrid = document.getElementById('imageGrid');
const downloadAllBtn = document.getElementById('downloadAllBtn');
const removePdfBtn = document.getElementById('removePdfBtn');

// Event listeners for drag and drop functionality
// Image to PDF drag and drop
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
            
            // Show the download and view buttons
            const downloadBtn = document.getElementById('downloadBtn');
            const viewPdfBtn = document.getElementById('viewPdfBtn');
            
            if (downloadBtn && viewPdfBtn) {
                // Display both buttons
                downloadBtn.style.display = 'inline-block';
                viewPdfBtn.style.display = 'inline-block';
                
                // Remove any existing event listeners by cloning the buttons
                const newDownloadBtn = downloadBtn.cloneNode(true);
                const newViewPdfBtn = viewPdfBtn.cloneNode(true);
                
                downloadBtn.parentNode.replaceChild(newDownloadBtn, downloadBtn);
                viewPdfBtn.parentNode.replaceChild(newViewPdfBtn, viewPdfBtn);
                
                // Set up download button with the new reference
                newDownloadBtn.addEventListener('click', async (e) => {
                    e.preventDefault(); // Prevent default button behavior
                    console.log('Download button clicked');
                    try {
                        await downloadPdf(pdfPath, pdfName);
                    } catch (err) {
                        console.error('Error in download handler:', err);
                        showStatus('Download failed. Please try again.', 'error');
                    }
                });
                
                // Set up view in browser button with the new reference
                newViewPdfBtn.addEventListener('click', async (e) => {
                    e.preventDefault(); // Prevent default button behavior
                    console.log('View PDF button clicked');
                    try {
                        await viewPdfInBrowser(pdfPath, pdfName);
                    } catch (err) {
                        console.error('Error in view handler:', err);
                        showStatus('Failed to view PDF. Please try downloading instead.', 'error');
                    }
                });
                
                showStatus(`PDF created successfully. You can download or view the PDF in browser.`, 'success');
            } else {
                showStatus('Action buttons not found', 'error');
            }
        } else {
            showStatus('Failed to create PDF. Please try again.', 'error');
        }
    } catch (error) {
        console.error('Error creating PDF:', error);
        showStatus('An error occurred while creating the PDF', 'error');
    }
});

// Function to download the PDF
async function downloadPdf(pdfPath, pdfName) {
    try {
        // Request the PDF file from the server
        // Check if pdfPath already has .pdf extension
        const fullPdfPath = pdfPath.toLowerCase().endsWith('.pdf') ? pdfPath : `${pdfPath}.pdf`;
        console.log('Requesting PDF file:', fullPdfPath);
        
        // Get the filename from the path for the download
        const fileName = fullPdfPath.split('\\').pop();
        console.log('File name for download:', fileName);
        
        const script = `
            $pdfPath = "${fullPdfPath}"
            Write-Host "Looking for PDF at: $pdfPath"
            if (Test-Path $pdfPath) {
                Write-Host "PDF file found, reading bytes..."
                $bytes = [System.IO.File]::ReadAllBytes($pdfPath)
                $base64 = [Convert]::ToBase64String($bytes)
                Write-Output $base64
            } else {
                Write-Error "PDF file not found: $pdfPath"
                exit 1
            }
        `;
        
        const base64Data = await runPowerShellScript(script);
        console.log('Received base64 data of length:', base64Data ? base64Data.length : 0);
        
        if (!base64Data || base64Data.length === 0) {
            throw new Error('No PDF data received from server');
        }
        
        // Create a download link with proper attributes
        const link = document.createElement('a');
        link.href = `data:application/pdf;base64,${base64Data}`;
        
        // Use the extracted filename or fallback to pdfName
        const downloadName = fileName || `${pdfName}.pdf`;
        link.download = downloadName;
        console.log('Setting download filename to:', downloadName);
        
        // Make sure the link has the download attribute
        link.setAttribute('download', downloadName);
        
        // Force the browser to treat this as a download
        link.type = 'application/pdf';
        link.target = '_blank';
        
        // Append link to body
        document.body.appendChild(link);
        
        // Use a small timeout to ensure the browser has time to process
        setTimeout(() => {
            console.log('Triggering download click');
            link.click();
            
            // Remove the link after a short delay
            setTimeout(() => {
                document.body.removeChild(link);
                showStatus(`PDF downloaded successfully as ${downloadName}`, 'success');
            }, 200);
        }, 200);
    } catch (error) {
        console.error('Error downloading PDF:', error);
        showStatus('Failed to download PDF', 'error');
    }
}

// Function to view PDF in browser
async function viewPdfInBrowser(pdfPath, pdfName) {
    try {
        // Request the PDF file from the server
        // Check if pdfPath already has .pdf extension
        const fullPdfPath = pdfPath.toLowerCase().endsWith('.pdf') ? pdfPath : `${pdfPath}.pdf`;
        console.log('Requesting PDF file for browser view:', fullPdfPath);
        
        const script = `
            $pdfPath = "${fullPdfPath}"
            Write-Host "Looking for PDF at: $pdfPath"
            if (Test-Path $pdfPath) {
                Write-Host "PDF file found, reading bytes..."
                $bytes = [System.IO.File]::ReadAllBytes($pdfPath)
                $base64 = [Convert]::ToBase64String($bytes)
                Write-Output $base64
            } else {
                Write-Error "PDF file not found: $pdfPath"
            }
        `;
        
        const base64Data = await runPowerShellScript(script);
        console.log('Received base64 data of length:', base64Data.length);
        
        if (!base64Data || base64Data.length === 0) {
            throw new Error('No PDF data received from server');
        }
        
        // Create a data URL for the PDF
        const pdfDataUrl = `data:application/pdf;base64,${base64Data}`;
        
        // Create a direct download link first to ensure the browser has the PDF data
        const link = document.createElement('a');
        link.href = pdfDataUrl;
        link.target = '_blank';
        document.body.appendChild(link);
        
        // Open the PDF in a new tab with proper viewer
        const newTab = window.open('', '_blank');
        if (newTab) {
            newTab.document.write(`
                <html>
                <head>
                    <title>${pdfName}.pdf</title>
                    <style>
                        body, html {
                            margin: 0;
                            padding: 0;
                            height: 100%;
                            overflow: hidden;
                        }
                        object {
                            width: 100%;
                            height: 100%;
                        }
                    </style>
                </head>
                <body>
                    <object data="${pdfDataUrl}" type="application/pdf" width="100%" height="100%">
                        <p>It appears your browser doesn't support embedded PDFs. 
                        <a href="${pdfDataUrl}" download="${pdfName}.pdf">Click here to download the PDF</a>.</p>
                    </object>
                </body>
                </html>
            `);
            newTab.document.close();
            
            // Remove the temporary link
            document.body.removeChild(link);
            
            showStatus(`PDF opened in a new browser tab`, 'success');
        } else {
            // If popup is blocked, offer direct download
            showStatus(`Popup blocked. Please allow popups to view PDF in browser.`, 'error');
            link.download = `${pdfName}.pdf`;
            link.click();
            document.body.removeChild(link);
        }
    } catch (error) {
        console.error('Error viewing PDF in browser:', error);
        showStatus('Failed to open PDF in browser', 'error');
    }
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
            // Clean up the result string to ensure it's valid JSON
            let cleanedStr = resultStr.trim();
            // If there are multiple lines, try to find the JSON part
            if (cleanedStr.includes('\n')) {
                const jsonLines = cleanedStr.split('\n').filter(line => 
                    line.trim().startsWith('{') && line.trim().endsWith('}'));
                if (jsonLines.length > 0) {
                    cleanedStr = jsonLines[0];
                }
            }
            
            const jsonResult = JSON.parse(cleanedStr);
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

// Tab switching functionality
tabButtons.forEach(button => {
    button.addEventListener('click', () => {
        // Remove active class from all buttons and contents
        tabButtons.forEach(btn => btn.classList.remove('active'));
        tabContents.forEach(content => content.classList.remove('active'));
        
        // Add active class to clicked button
        button.classList.add('active');
        
        // Show corresponding content
        const tabId = button.getAttribute('data-tab');
        document.getElementById(`${tabId}-content`).classList.add('active');
    });
});

// Page range toggle for PDF to Image
extractPages.addEventListener('change', () => {
    if (extractPages.value === 'range') {
        pageRangeOption.style.display = 'flex';
    } else {
        pageRangeOption.style.display = 'none';
    }
});

// Word to PDF drag and drop functionality
wordDropArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    wordDropArea.classList.add('dragover');
});

wordDropArea.addEventListener('dragleave', () => {
    wordDropArea.classList.remove('dragover');
});

wordDropArea.addEventListener('drop', (e) => {
    e.preventDefault();
    wordDropArea.classList.remove('dragover');
    
    if (e.dataTransfer.files.length > 0) {
        handleWordFile(e.dataTransfer.files[0]);
    }
});

wordDropArea.addEventListener('click', () => {
    wordFileInput.click();
});

wordFileInput.addEventListener('change', () => {
    if (wordFileInput.files.length > 0) {
        handleWordFile(wordFileInput.files[0]);
    }
});

// Handle Word file selection
function handleWordFile(file) {
    // Check if it's a Word document
    if (!file.name.toLowerCase().endsWith('.doc') && !file.name.toLowerCase().endsWith('.docx')) {
        showStatus('Please select a valid Word document (.doc or .docx)', 'error');
        return;
    }
    
    // Store the selected Word document
    selectedWordDoc = file;
    
    // Update the UI
    updateWordDocPreview();
    updateWordButtonStates();
    
    // Show success message
    showStatus(`Word document "${file.name}" selected successfully`, 'success');
}

// Update the Word document preview
function updateWordDocPreview() {
    if (selectedWordDoc) {
        noDocText.style.display = 'none';
        docInfo.style.display = 'flex';
        docName.textContent = selectedWordDoc.name;
    } else {
        noDocText.style.display = 'block';
        docInfo.style.display = 'none';
    }
}

// Update Word to PDF button states
function updateWordButtonStates() {
    wordConvertBtn.disabled = !selectedWordDoc;
    wordClearBtn.disabled = !selectedWordDoc;
}

// Remove Word document
removeDocBtn.addEventListener('click', () => {
    selectedWordDoc = null;
    updateWordDocPreview();
    updateWordButtonStates();
    wordDownloadBtn.style.display = 'none';
    wordViewPdfBtn.style.display = 'none';
    showStatus('Word document removed', 'info');
});

// Clear Word document
wordClearBtn.addEventListener('click', () => {
    selectedWordDoc = null;
    updateWordDocPreview();
    updateWordButtonStates();
    wordDownloadBtn.style.display = 'none';
    wordViewPdfBtn.style.display = 'none';
    showStatus('Word document cleared', 'info');
});

// Convert Word to PDF
wordConvertBtn.addEventListener('click', async () => {
    if (!selectedWordDoc) {
        showStatus('Please select a Word document', 'error');
        return;
    }
    
    // Show processing status
    showStatus('Converting Word document to PDF...', 'info');
    
    try {
        // Get PDF name
        const pdfName = wordPdfNameInput.value.trim() || 'converted_document';
        
        // Save Word document to temporary folder
        const tempFolder = await createTempFolder();
        const wordFilePath = await saveWordToTemp(tempFolder, selectedWordDoc);
        
        // Run PowerShell script to convert Word to PDF
        const result = await convertWordToPdf(wordFilePath, pdfName);
        console.log('Convert Word to PDF result:', result);
        
        if (result.success) {
            // Show the download and view buttons
            wordDownloadBtn.style.display = 'inline-block';
            wordViewPdfBtn.style.display = 'inline-block';
            
            // Remove any existing event listeners by cloning the buttons
            const newDownloadBtn = wordDownloadBtn.cloneNode(true);
            const newViewPdfBtn = wordViewPdfBtn.cloneNode(true);
            
            wordDownloadBtn.parentNode.replaceChild(newDownloadBtn, wordDownloadBtn);
            wordViewPdfBtn.parentNode.replaceChild(newViewPdfBtn, wordViewPdfBtn);
            
            // Update references
            wordDownloadBtn = newDownloadBtn;
            wordViewPdfBtn = newViewPdfBtn;
            
            // Set up download button
            wordDownloadBtn.addEventListener('click', async () => {
                try {
                    await downloadPdf(result.path, pdfName);
                } catch (err) {
                    console.error('Error in download handler:', err);
                    showStatus('Download failed. Please try again.', 'error');
                }
            });
            
            // Set up view in browser button
            wordViewPdfBtn.addEventListener('click', async () => {
                try {
                    await viewPdfInBrowser(result.path, pdfName);
                } catch (err) {
                    console.error('Error in view handler:', err);
                    showStatus('Failed to view PDF. Please try downloading instead.', 'error');
                }
            });
            
            showStatus(`Word document converted to PDF successfully`, 'success');
        } else {
            showStatus('Failed to convert Word to PDF. Please try again.', 'error');
        }
    } catch (error) {
        console.error('Error converting Word to PDF:', error);
        showStatus('An error occurred while converting the Word document', 'error');
    }
});

// Save Word document to temporary folder
async function saveWordToTemp(tempFolder, wordFile) {
    const reader = new FileReader();
    
    // Convert the file to base64
    const base64Data = await new Promise((resolve, reject) => {
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = reject;
        reader.readAsDataURL(wordFile);
    });
    
    // Use PowerShell to save the file
    const fileName = wordFile.name;
    const filePath = `${tempFolder}\\${fileName}`;
    const script = `
        $base64Data = "${base64Data}"
        $bytes = [Convert]::FromBase64String($base64Data)
        $filePath = "${filePath}"
        [System.IO.File]::WriteAllBytes($filePath, $bytes)
        Write-Output $filePath
    `;
    
    try {
        const result = await runPowerShellScript(script);
        return result.trim();
    } catch (error) {
        console.error('Error saving Word file:', error);
        throw error;
    }
}

// Convert Word to PDF using PowerShell
async function convertWordToPdf(wordFilePath, pdfName) {
    // PowerShell script to convert Word to PDF using Word COM object
    const script = `
        try {
            # Load Word COM object
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            
            # Open the document
            $wordFilePath = "${wordFilePath}"
            $doc = $word.Documents.Open($wordFilePath)
            
            # Set PDF output path
            $pdfPath = "${pdfName}.pdf"
            
            # Save as PDF
            $wdFormatPDF = 17  # PDF format constant
            $doc.SaveAs([ref]$pdfPath, [ref]$wdFormatPDF)
            
            # Close document and Word application
            $doc.Close()
            $word.Quit()
            
            # Release COM objects
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            
            # Get the full path of the PDF
            $fullPath = (Get-Item -Path $pdfPath).FullName
            
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
            // Clean up the result string to ensure it's valid JSON
            let cleanedStr = resultStr.trim();
            // If there are multiple lines, try to find the JSON part
            if (cleanedStr.includes('\n')) {
                const jsonLines = cleanedStr.split('\n').filter(line => 
                    line.trim().startsWith('{') && line.trim().endsWith('}'));
                if (jsonLines.length > 0) {
                    cleanedStr = jsonLines[0];
                }
            }
            
            const jsonResult = JSON.parse(cleanedStr);
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

// PDF to Image drag and drop functionality
pdfDropArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    pdfDropArea.classList.add('dragover');
});

pdfDropArea.addEventListener('dragleave', () => {
    pdfDropArea.classList.remove('dragover');
});

pdfDropArea.addEventListener('drop', (e) => {
    e.preventDefault();
    pdfDropArea.classList.remove('dragover');
    
    if (e.dataTransfer.files.length > 0) {
        handlePdfFile(e.dataTransfer.files[0]);
    }
});

pdfDropArea.addEventListener('click', () => {
    pdfFileInput.click();
});

pdfFileInput.addEventListener('change', () => {
    if (pdfFileInput.files.length > 0) {
        handlePdfFile(pdfFileInput.files[0]);
    }
});

// Handle PDF file selection
function handlePdfFile(file) {
    // Check if it's a PDF document
    if (!file.name.toLowerCase().endsWith('.pdf')) {
        showStatus('Please select a valid PDF document', 'error');
        return;
    }
    
    // Store the selected PDF document
    selectedPdf = file;
    
    // Update the UI
    updatePdfPreview();
    updatePdfButtonStates();
    
    // Show success message
    showStatus(`PDF document "${file.name}" selected successfully`, 'success');
}

// Update the PDF preview
function updatePdfPreview() {
    if (selectedPdf) {
        noPdfText.style.display = 'none';
        pdfInfo.style.display = 'flex';
        pdfFileName.textContent = selectedPdf.name;
    } else {
        noPdfText.style.display = 'block';
        pdfInfo.style.display = 'none';
    }
}

// Update PDF to Image button states
function updatePdfButtonStates() {
    pdfConvertBtn.disabled = !selectedPdf;
    pdfClearBtn.disabled = !selectedPdf;
}

// Remove PDF document
removePdfBtn.addEventListener('click', () => {
    selectedPdf = null;
    updatePdfPreview();
    updatePdfButtonStates();
    extractedImages.style.display = 'none';
    showStatus('PDF document removed', 'info');
});

// Clear PDF document
pdfClearBtn.addEventListener('click', () => {
    selectedPdf = null;
    updatePdfPreview();
    updatePdfButtonStates();
    extractedImages.style.display = 'none';
    showStatus('PDF document cleared', 'info');
});

// Convert PDF to Images
pdfConvertBtn.addEventListener('click', async () => {
    if (!selectedPdf) {
        showStatus('Please select a PDF document', 'error');
        return;
    }
    
    // Show processing status
    showStatus('Converting PDF to images...', 'info');
    
    try {
        // Get options
        const imageFormat = document.getElementById('imageFormat').value;
        const imageQuality = document.getElementById('imageQuality').value;
        const extractOption = extractPages.value;
        const pageRangeValue = pageRange.value.trim();
        
        // Validate page range if selected
        if (extractOption === 'range' && !pageRangeValue) {
            showStatus('Please enter a valid page range', 'error');
            return;
        }
        
        // Save PDF to temporary folder
        const tempFolder = await createTempFolder();
        const pdfFilePath = await savePdfToTemp(tempFolder, selectedPdf);
        
        // Run PowerShell script to convert PDF to images
        const result = await convertPdfToImages(pdfFilePath, tempFolder, imageFormat, imageQuality, extractOption, pageRangeValue);
        console.log('Convert PDF to Images result:', result);
        
        if (result.success) {
            // Display the extracted images
            displayExtractedImages(result.images);
            showStatus(`PDF converted to ${result.images.length} images successfully`, 'success');
        } else {
            showStatus('Failed to convert PDF to images. Please try again.', 'error');
        }
    } catch (error) {
        console.error('Error converting PDF to images:', error);
        showStatus('An error occurred while converting the PDF', 'error');
    }
});

// Save PDF to temporary folder
async function savePdfToTemp(tempFolder, pdfFile) {
    const reader = new FileReader();
    
    // Convert the file to base64
    const base64Data = await new Promise((resolve, reject) => {
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = reject;
        reader.readAsDataURL(pdfFile);
    });
    
    // Use PowerShell to save the file
    const fileName = pdfFile.name;
    const filePath = `${tempFolder}\\${fileName}`;
    const script = `
        $base64Data = "${base64Data}"
        $bytes = [Convert]::FromBase64String($base64Data)
        $filePath = "${filePath}"
        [System.IO.File]::WriteAllBytes($filePath, $bytes)
        Write-Output $filePath
    `;
    
    try {
        const result = await runPowerShellScript(script);
        return result.trim();
    } catch (error) {
        console.error('Error saving PDF file:', error);
        throw error;
    }
}

// Convert PDF to Images using PowerShell
async function convertPdfToImages(pdfFilePath, outputFolder, format, quality, extractOption, pageRange) {
    // PowerShell script to convert PDF to images
    const script = `
        try {
            # Add necessary assemblies
            Add-Type -AssemblyName System.Drawing
            Add-Type -AssemblyName System.IO
            
            # Load PDF processing library (using GhostScript via .NET)
            $ghostscriptPath = "C:\\Program Files\\gs\\gs9.54.0\\bin\\gswin64c.exe"
            
            # Check if GhostScript is installed
            if (-not (Test-Path $ghostscriptPath)) {
                # Try to find GhostScript in common locations
                $ghostscriptPath = (Get-ChildItem -Path "C:\\Program Files\\gs" -Filter "gswin64c.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1).FullName
                
                if (-not $ghostscriptPath) {
                    throw "GhostScript not found. Please install GhostScript to convert PDF to images."
                }
            }
            
            # Define paths
            $pdfPath = "${pdfFilePath}"
            $outputFolder = "${outputFolder}"
            $format = "${format}".ToLower()
            $quality = "${quality}"
            
            # Set quality parameter based on selected quality
            $qualityParam = switch ($quality) {
                "high" { "-r300" }
                "medium" { "-r150" }
                "low" { "-r72" }
                default { "-r150" }
            }
            
            # Determine which pages to extract
            $extractOption = "${extractOption}"
            $pageRangeValue = "${pageRange}"
            
            $pageParam = ""
            if ($extractOption -eq "range" -and $pageRangeValue) {
                # Parse page range (e.g., "1-3,5,7-9")
                $pageParam = "-dFirstPage=1 -dLastPage=1"
                
                # We'll handle page ranges by extracting one page at a time
                $pageRanges = $pageRangeValue -split ','
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
            }
            
            # Create output directory for images if it doesn't exist
            $imagesFolder = Join-Path -Path $outputFolder -ChildPath "images"
            if (-not (Test-Path $imagesFolder)) {
                New-Item -ItemType Directory -Path $imagesFolder | Out-Null
            }
            
            # List to store image paths
            $imageFiles = @()
            
            if ($extractOption -eq "range" -and $pageList.Count -gt 0) {
                # Extract specific pages
                foreach ($pageNum in $pageList) {
                    $outputFile = Join-Path -Path $imagesFolder -ChildPath "page_${pageNum}.${format}"
                    
                    # Use GhostScript to convert PDF page to image
                    $arguments = @(
                        "-dNOPAUSE",
                        "-dBATCH",
                        "-dSAFER",
                        $qualityParam,
                        "-dFirstPage=${pageNum}",
                        "-dLastPage=${pageNum}",
                        "-sDEVICE=png16m",
                        "-sOutputFile=${outputFile}",
                        $pdfPath
                    )
                    
                    Start-Process -FilePath $ghostscriptPath -ArgumentList $arguments -Wait -NoNewWindow
                    
                    if (Test-Path $outputFile) {
                        $imageFiles += @{
                            "path" = $outputFile
                            "page" = $pageNum
                            "base64" = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($outputFile))
                        }
                    }
                }
            } else {
                # Extract all pages
                $outputPattern = Join-Path -Path $imagesFolder -ChildPath "page_%03d.${format}"
                
                # Use GhostScript to convert all PDF pages to images
                $arguments = @(
                    "-dNOPAUSE",
                    "-dBATCH",
                    "-dSAFER",
                    $qualityParam,
                    "-sDEVICE=png16m",
                    "-sOutputFile=${outputPattern}",
                    $pdfPath
                )
                
                Start-Process -FilePath $ghostscriptPath -ArgumentList $arguments -Wait -NoNewWindow
                
                # Get all generated image files
                $generatedImages = Get-ChildItem -Path $imagesFolder -Filter "page_*.${format}" | Sort-Object Name
                
                foreach ($img in $generatedImages) {
                    # Extract page number from filename (page_001.png -> 1)
                    $pageNum = [int]($img.BaseName -replace 'page_', '' -replace '^0+', '')
                    if ($pageNum -eq 0) { $pageNum = 1 } # Handle case where all zeros were removed
                    
                    $imageFiles += @{
                        "path" = $img.FullName
                        "page" = $pageNum
                        "base64" = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($img.FullName))
                    }
                }
            }
            
            # Return success with the image files
            $result = @{
                "success" = $true
                "message" = "PDF converted to images successfully"
                "images" = $imageFiles
            } | ConvertTo-Json -Depth 5
            
            Write-Output $result
        } catch {
            # Return error
            $result = @{
                "success" = $false
                "message" = "Error converting PDF to images: $_"
            } | ConvertTo-Json
            
            Write-Output $result
        }
    `;
    
    try {
        const resultStr = await runPowerShellScript(script);
        console.log('Raw PowerShell result length:', resultStr.length);
        
        try {
            // Clean up the result string to ensure it's valid JSON
            let cleanedStr = resultStr.trim();
            // If there are multiple lines, try to find the JSON part
            if (cleanedStr.includes('\n')) {
                const jsonStart = cleanedStr.indexOf('{');
                const jsonEnd = cleanedStr.lastIndexOf('}') + 1;
                if (jsonStart >= 0 && jsonEnd > jsonStart) {
                    cleanedStr = cleanedStr.substring(jsonStart, jsonEnd);
                }
            }
            
            const jsonResult = JSON.parse(cleanedStr);
            console.log('Parsed JSON result:', jsonResult);
            return jsonResult;
        } catch (parseError) {
            console.error('JSON parse error:', parseError);
            return { success: false, message: 'Failed to parse conversion result' };
        }
    } catch (error) {
        console.error('Error in PowerShell script:', error);
        return { success: false, message: error.toString() };
    }
}

// Display extracted images in the UI
function displayExtractedImages(images) {
    // Clear previous images
    imageGrid.innerHTML = '';
    
    // Sort images by page number
    images.sort((a, b) => a.page - b.page);
    
    // Create image elements
    images.forEach(image => {
        const imageItem = document.createElement('div');
        imageItem.className = 'image-grid-item';
        
        // Create image element
        const img = document.createElement('img');
        img.src = `data:image/${image.path.split('.').pop()};base64,${image.base64}`;
        img.alt = `Page ${image.page}`;
        
        // Create image actions
        const actions = document.createElement('div');
        actions.className = 'image-actions';
        
        // Page number
        const pageNum = document.createElement('span');
        pageNum.textContent = `Page ${image.page}`;
        
        // Download button
        const downloadBtn = document.createElement('button');
        downloadBtn.textContent = 'Download';
        downloadBtn.addEventListener('click', () => {
            downloadImage(image);
        });
        
        // Add elements to the DOM
        actions.appendChild(pageNum);
        actions.appendChild(downloadBtn);
        
        imageItem.appendChild(img);
        imageItem.appendChild(actions);
        imageGrid.appendChild(imageItem);
    });
    
    // Show the extracted images section
    extractedImages.style.display = 'block';
    
    // Set up download all button
    downloadAllBtn.onclick = () => {
        downloadAllImages(images);
    };
}

// Download a single image
function downloadImage(image) {
    const link = document.createElement('a');
    link.href = `data:image/${image.path.split('.').pop()};base64,${image.base64}`;
    link.download = `page_${image.page}.${image.path.split('.').pop()}`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Download all images as a ZIP file
async function downloadAllImages(images) {
    showStatus('Preparing images for download...', 'info');
    
    try {
        // Create a temporary folder to store the images
        const tempFolder = await createTempFolder();
        
        // Save all images to the temp folder
        for (const image of images) {
            const script = `
                $base64Data = "${image.base64}"
                $bytes = [Convert]::FromBase64String($base64Data)
                $filePath = Join-Path -Path "${tempFolder}" -ChildPath "page_${image.page}.${image.path.split('.').pop()}"
                [System.IO.File]::WriteAllBytes($filePath, $bytes)
                Write-Output $filePath
            `;
            
            await runPowerShellScript(script);
        }
        
        // Create a ZIP file containing all images
        const zipScript = `
            $tempFolder = "${tempFolder}"
            $zipPath = "${tempFolder}_images.zip"
            
            # Create ZIP file
            Compress-Archive -Path "${tempFolder}\\*" -DestinationPath $zipPath -Force
            
            # Read ZIP file as base64
            $bytes = [System.IO.File]::ReadAllBytes($zipPath)
            $base64 = [Convert]::ToBase64String($bytes)
            
            # Clean up
            Remove-Item -Path $tempFolder -Recurse -Force
            Remove-Item -Path $zipPath -Force
            
            Write-Output $base64
        `;
        
        const base64Zip = await runPowerShellScript(zipScript);
        
        // Create download link for the ZIP file
        const link = document.createElement('a');
        link.href = `data:application/zip;base64,${base64Zip}`;
        link.download = 'extracted_images.zip';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showStatus('All images downloaded as ZIP file', 'success');
    } catch (error) {
        console.error('Error downloading all images:', error);
        showStatus('Failed to download images', 'error');
    }
}

// Initialize the UI
updateButtonStates();