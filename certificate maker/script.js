document.addEventListener('DOMContentLoaded', () => {
    // UI Elements
    const templateUpload = document.getElementById('templateUpload');
    const templateFileName = document.getElementById('templateFileName');
    const excelUpload = document.getElementById('excelUpload');
    const excelFileName = document.getElementById('excelFileName');
    
    const manualNameInput = document.getElementById('manualNameInput');
    const addNameBtn = document.getElementById('addNameBtn');
    
    const fontFamily = document.getElementById('fontFamily');
    const fontSize = document.getElementById('fontSize');
    const fontSizeDisplay = document.getElementById('fontSizeDisplay');
    const fontColor = document.getElementById('fontColor');
    const fontWeight = document.getElementById('fontWeight');
    
    const generateBtn = document.getElementById('generateBtn');
    const namesCountEl = document.getElementById('namesCount');
    
    const canvasContainer = document.getElementById('canvasContainer');
    const previewCanvas = document.getElementById('previewCanvas');
    const canvasPlaceholder = document.getElementById('canvasPlaceholder');
    
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');

    // State
    let templateImage = null;
    let namesList = [];
    let textPos = { x: 50, y: 50 }; // Relative percentage (0-100)
    
    // Create Draggable Text Element
    const draggableText = document.createElement('div');
    draggableText.className = 'draggable-text hidden';
    draggableText.textContent = 'Sample Name';
    canvasContainer.appendChild(draggableText);

    // Context for Canvas
    const ctx = previewCanvas.getContext('2d');

    // Update typography live
    function updateTypography() {
        fontSizeDisplay.textContent = fontSize.value;
        
        // Update draggable text style
        draggableText.style.fontFamily = fontFamily.value;
        // The font size of the draggable text is scaled down visually to match the canvas scaling
        // We handle this in the draw function, but for the UI let's give it a proportional size
        draggableText.style.color = fontColor.value;
        draggableText.style.fontWeight = fontWeight.value;
        
        // Force redraw if template is loaded
        if (templateImage) {
            updateDraggableTextSize();
        }
    }

    [fontFamily, fontSize, fontColor, fontWeight].forEach(el => {
        el.addEventListener('input', updateTypography);
    });

    // Handle Template Upload
    templateUpload.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        templateFileName.textContent = file.name;
        
        const reader = new FileReader();
        reader.onload = (event) => {
            const img = new Image();
            img.onload = () => {
                templateImage = img;
                
                // Set actual canvas size to image resolution
                previewCanvas.width = img.width;
                previewCanvas.height = img.height;
                
                // Show canvas, hide placeholder
                previewCanvas.style.display = 'block';
                canvasPlaceholder.classList.add('hidden');
                draggableText.classList.remove('hidden');
                
                // Initial draw and position
                ctx.drawImage(img, 0, 0);
                
                // Default position center
                textPos = { x: img.width / 2, y: img.height / 2 };
                updateDraggableTextPosition();
                updateDraggableTextSize();
                
                checkReadyState();
            };
            img.src = event.target.result;
        };
        reader.readAsDataURL(file);
    });

    // Update the visual draggable text position to match the canvas element's bounding rect
    function updateDraggableTextPosition() {
        if (!templateImage) return;
        
        const rect = previewCanvas.getBoundingClientRect();
        // Calculate the scale factor between original image and rendered canvas
        const scaleX = rect.width / templateImage.width;
        const scaleY = rect.height / templateImage.height;
        
        // textPos is in original image pixel coordinates
        const containerRect = canvasContainer.getBoundingClientRect();
        const visualX = (textPos.x * scaleX) + (rect.left - containerRect.left);
        const visualY = (textPos.y * scaleY) + (rect.top - containerRect.top);
        
        draggableText.style.left = `${visualX}px`;
        draggableText.style.top = `${visualY}px`;
    }

    function updateDraggableTextSize() {
        if (!templateImage) return;
        
        const rect = previewCanvas.getBoundingClientRect();
        const scale = rect.height / templateImage.height;
        
        // Scale the font size visually
        const scaledFontSize = Math.max(1, parseInt(fontSize.value) * scale);
        draggableText.style.fontSize = `${scaledFontSize}px`;
    }

    // Handle Window Resize
    window.addEventListener('resize', () => {
        updateDraggableTextPosition();
        updateDraggableTextSize();
    });

    // Drag Logic for Draggable Text
    let isDragging = false;
    
    // Mouse start
    draggableText.addEventListener('mousedown', (e) => {
        isDragging = true;
        e.preventDefault(); // Prevent text selection
    });
    
    // Touch start
    draggableText.addEventListener('touchstart', (e) => {
        isDragging = true;
        e.preventDefault(); // Prevent scrolling when touching the text
    }, { passive: false });
    
    // Common move handler
    function handleMove(clientX, clientY) {
        if (!isDragging || !templateImage) return;
        
        const rect = previewCanvas.getBoundingClientRect();
        
        // Clamp to canvas bounds visually
        if (clientX < rect.left) clientX = rect.left;
        if (clientX > rect.right) clientX = rect.right;
        if (clientY < rect.top) clientY = rect.top;
        if (clientY > rect.bottom) clientY = rect.bottom;
        
        // Convert to original image coordinates
        const scaleX = templateImage.width / rect.width;
        const scaleY = templateImage.height / rect.height;
        
        const imgX = (clientX - rect.left) * scaleX;
        const imgY = (clientY - rect.top) * scaleY;
        
        textPos.x = imgX;
        textPos.y = imgY;
        
        updateDraggableTextPosition();
    }
    
    // Mouse move
    document.addEventListener('mousemove', (e) => {
        if (!isDragging) return;
        handleMove(e.clientX, e.clientY);
    });
    
    // Touch move
    document.addEventListener('touchmove', (e) => {
        if (!isDragging) return;
        e.preventDefault(); // Prevent scrolling while dragging
        handleMove(e.touches[0].clientX, e.touches[0].clientY);
    }, { passive: false });
    
    // Mouse end
    document.addEventListener('mouseup', () => {
        isDragging = false;
    });
    
    // Touch end
    document.addEventListener('touchend', () => {
        isDragging = false;
    });

    // Handle Excel Upload
    excelUpload.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        excelFileName.textContent = file.name;
        
        const reader = new FileReader();
        reader.onload = (event) => {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Get first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON (array of arrays)
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            // Extract first column, skip empty, limit to 500
            namesList = [];
            for (let i = 0; i < json.length; i++) {
                const row = json[i];
                if (row && row.length > 0 && row[0] !== undefined && row[0] !== null && String(row[0]).trim() !== '') {
                    // Skip if it looks like a header row (e.g., 'Name', 'Student') on the very first row
                    if (i === 0 && (String(row[0]).toLowerCase() === 'name' || String(row[0]).toLowerCase() === 'student name')) {
                        continue;
                    }
                    namesList.push(String(row[0]).trim());
                }
            }
            
            // Apply maximum limit of 500
            if (namesList.length > 500) {
                namesList = namesList.slice(0, 500);
                alert(`Warning: The Excel sheet contains more than 500 names. Only the first 500 names will be processed as requested.`);
            }
            
            namesCountEl.textContent = namesList.length;
            checkReadyState();
        };
        reader.readAsArrayBuffer(file);
    });

    // Handle Manual Name Add
    function handleManualAdd() {
        const name = manualNameInput.value.trim();
        if (name) {
            namesList.push(name);
            namesCountEl.textContent = namesList.length;
            manualNameInput.value = '';
            checkReadyState();
            // Clear excel file name if they manually add something after so it's not confusing, 
            // though they can use both together. Let's just keep the text.
        }
    }

    addNameBtn.addEventListener('click', handleManualAdd);
    manualNameInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            handleManualAdd();
        }
    });

    // Enable button if both are present
    function checkReadyState() {
        if (templateImage && namesList.length > 0) {
            generateBtn.disabled = false;
        } else {
            generateBtn.disabled = true;
        }
    }

    // Generate Certificates
    generateBtn.addEventListener('click', async () => {
        if (!templateImage || namesList.length === 0) return;
        
        generateBtn.disabled = true;
        progressContainer.classList.remove('hidden');
        
        const zip = new JSZip();
        const folder = zip.folder("certificates");
        
        // Create an offscreen canvas for rendering
        const offscreenCanvas = document.createElement('canvas');
        offscreenCanvas.width = templateImage.width;
        offscreenCanvas.height = templateImage.height;
        const offCtx = offscreenCanvas.getContext('2d');
        
        const total = namesList.length;
        
        // Process sequentially to not freeze browser completely
        for (let i = 0; i < total; i++) {
            const name = namesList[i];
            
            // Update progress UI
            progressText.textContent = `Generating: ${i + 1} / ${total}`;
            progressBar.style.width = `${((i + 1) / total) * 100}%`;
            
            // Clear and draw image
            offCtx.clearRect(0, 0, offscreenCanvas.width, offscreenCanvas.height);
            offCtx.drawImage(templateImage, 0, 0);
            
            // Draw text
            offCtx.font = `${fontWeight.value} ${fontSize.value}px ${fontFamily.value}`;
            offCtx.fillStyle = fontColor.value;
            offCtx.textAlign = 'center';
            offCtx.textBaseline = 'middle';
            
            offCtx.fillText(name, textPos.x, textPos.y);
            
            // Convert to Blob and add to zip
            const blob = await new Promise(resolve => {
                offscreenCanvas.toBlob(resolve, 'image/jpeg', 0.95);
            });
            
            // Generate clean filename
            const safeName = name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
            folder.file(`certificate_${safeName}_${i+1}.jpg`, blob);
            
            // Yield to main thread to allow UI updates
            await new Promise(resolve => setTimeout(resolve, 0));
        }
        
        progressText.textContent = "Zipping files... please wait";
        
        // Generate Zip
        const content = await zip.generateAsync({ type: "blob" });
        
        // Trigger Download
        saveAs(content, "Certificates.zip");
        
        // Reset UI
        setTimeout(() => {
            progressContainer.classList.add('hidden');
            progressBar.style.width = '0%';
            generateBtn.disabled = false;
            progressText.textContent = `Generating: 0 / 0`;
        }, 1000);
    });
    
    // Initialize
    updateTypography();
});
