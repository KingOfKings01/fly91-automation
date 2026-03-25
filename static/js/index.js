function handleFileSelect(input, type) {
    const fileNameDisplay = document.getElementById(`${type}FileName`);
    if (input.files.length > 0) {
        const name = input.files[0].name;
        if (fileNameDisplay) {
            fileNameDisplay.textContent = `Selected: ${name}`;
            fileNameDisplay.style.display = 'block';
        }
        
        // If it's the excel file, we don't upload automatically, wait for submit
        // If it's seal/sign, we still use the uploadMedia function which is already called via onchange
    }
}

document.getElementById('excelForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const formData = new FormData(e.target);
    const btn = document.getElementById('submitBtn');
    const originalText = btn.textContent;
    
    btn.textContent = 'Processing...';
    btn.disabled = true;
    btn.style.opacity = '0.7';

    try {
        const response = await fetch('/upload_excel', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();
        if (data.success) {
            // Check for saved positions to provide "auto-align" on first load
            const sealPos = localStorage.getItem('fly91_seal_pos');
            const signPos = localStorage.getItem('fly91_sign_pos');
            let redirectUrl = `/preview_first?excel=${data.filename}`;
            if (sealPos) redirectUrl += `&seal_pos=${encodeURIComponent(sealPos)}`;
            if (signPos) redirectUrl += `&sign_pos=${encodeURIComponent(signPos)}`;
            
            window.location.href = redirectUrl;
        } else {
            alert(data.error || 'Upload failed');
            btn.textContent = originalText;
            btn.disabled = false;
            btn.style.opacity = '1';
        }
    } catch (err) {
        alert('An error occurred during upload');
        btn.textContent = originalText;
        btn.disabled = false;
        btn.style.opacity = '1';
    }
});

async function uploadMedia(input, type) {
    if (!input.files || !input.files[0]) return;
    
    // Show loading state for the compact-upload area if possible
    const parentArea = input.closest('.compact-upload');
    const originalBg = parentArea.style.background;
    parentArea.style.background = '#f1f5f9';
    
    const formData = new FormData();
    formData.append('file', input.files[0]);
    formData.append('type', type);

    try {
        const response = await fetch('/upload_media', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();
        if (data.success) {
            const previewId = type === 'seal' ? 'sealPreview' : 'signPreview';
            const img = document.getElementById(previewId);
            img.src = data.url + '&t=' + new Date().getTime(); // cache busting
            img.style.display = 'block';
        } else {
            alert(data.error || 'Upload failed');
        }
    } catch (err) {
        alert('An error occurred');
    } finally {
        parentArea.style.background = originalBg;
    }
}

// Optional: Drag and drop visual feedback
const dropZone = document.getElementById('excelDropZone');
if (dropZone) {
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.add('dragging');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragging');
        }, false);
    });
}
