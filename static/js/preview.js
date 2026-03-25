const container = document.getElementById('pdf-container');
const draggables = document.querySelectorAll('.draggable');
const progressModal = document.getElementById('progress-modal');
const progressBar = document.getElementById('progress-bar');
const progressText = document.getElementById('progress-text');
const progressStatus = document.getElementById('progress-status');
const pdfIframe = document.getElementById('pdf-iframe');
const btnConfirm = document.getElementById('btn-confirm');
const btnGenerate = document.getElementById('btn-generate');
const confirmedBadge = document.getElementById('confirmed-badge');

let isDragging = false;
let isResizing = false;
let activeEl = null;
let startX, startY, startLeft, startTop, startWidth, startHeight;
let layoutConfirmed = false;

// Store confirmed positions so they survive overlay hiding
let confirmedPositions = null;

function pxToMm(px) {
    return px * (297 / container.offsetWidth);
}

function getPositions() {
    const seal = document.getElementById('seal-drag');
    const sign = document.getElementById('sign-drag');
    return {
        seal_pos: { x: pxToMm(seal.offsetLeft), y: pxToMm(seal.offsetTop), w: pxToMm(seal.offsetWidth) },
        sign_pos: { x: pxToMm(sign.offsetLeft), y: pxToMm(sign.offsetTop), w: pxToMm(sign.offsetWidth) }
    };
}

function updateStats(el) {
    const rect = el.getBoundingClientRect();
    const containerRect = container.getBoundingClientRect();
    const x = pxToMm(rect.left - containerRect.left);
    const y = pxToMm(rect.top - containerRect.top);
    const w = pxToMm(rect.width);
    const type = el.classList.contains('seal') ? 'seal' : 'sign';
    document.getElementById(`${type}-x`).textContent = Math.round(x);
    document.getElementById(`${type}-y`).textContent = Math.round(y);
    document.getElementById(`${type}-w`).textContent = Math.round(w);
    localStorage.setItem(`fly91_${type}_pos`, JSON.stringify({x, y, w}));
}

// Load remembered positions
function loadPositions() {
    ['seal', 'sign'].forEach(type => {
        const saved = localStorage.getItem(`fly91_${type}_pos`);
        if (saved) {
            const pos = JSON.parse(saved);
            const el = document.querySelector(`.draggable.${type}`);
            el.style.left = `${pos.x}mm`;
            el.style.top = `${pos.y}mm`;
            el.style.width = `${pos.w}mm`;
            if (type === 'seal') el.style.height = `${pos.w}mm`;
        }
    });
}

loadPositions();
draggables.forEach(el => {
    updateStats(el);
    el.addEventListener('mousedown', (e) => {
        if (e.target.classList.contains('resizer')) isResizing = true;
        else isDragging = true;
        activeEl = el;
        startX = e.clientX; startY = e.clientY;
        startLeft = el.offsetLeft; startTop = el.offsetTop;
        startWidth = el.offsetWidth; startHeight = el.offsetHeight;
        e.preventDefault();
        // Reset confirmation if user moves images
        if (layoutConfirmed) {
            layoutConfirmed = false;
            confirmedPositions = null;
            btnGenerate.disabled = true;
            confirmedBadge.style.display = 'none';
            btnConfirm.textContent = 'Step 1: Confirm Layout';
            // Show overlay again so user can reposition
            document.getElementById('overlay').style.visibility = 'visible';
        }
    });
});

window.addEventListener('mousemove', (e) => {
    if (!activeEl) return;
    const dx = e.clientX - startX;
    const dy = e.clientY - startY;
    if (isDragging) {
        activeEl.style.left = `${startLeft + dx}px`;
        activeEl.style.top = `${startTop + dy}px`;
    } else if (isResizing) {
        activeEl.style.width = `${startWidth + dx}px`;
        if (activeEl.classList.contains('seal')) activeEl.style.height = `${startWidth + dx}px`;
        else activeEl.style.height = `${startHeight + dy}px`;
    }
    updateStats(activeEl);
});

window.addEventListener('mouseup', () => {
    isDragging = false;
    isResizing = false;
    activeEl = null;
});

// STEP 1: Confirm layout by generating the first PDF with images baked in
async function confirmLayout() {
    // Capture positions BEFORE hiding anything
    confirmedPositions = getPositions();
    
    btnConfirm.textContent = 'Generating preview...';
    btnConfirm.disabled = true;
    
    const body = {
        excel_filename: excelFilename,
        seal_pos: confirmedPositions.seal_pos,
        sign_pos: confirmedPositions.sign_pos
    };

    try {
        const res = await fetch('/refresh_preview', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        const data = await res.json();
        if (data.success) {
            // Hide overlay with visibility (not display:none) so offsetLeft still works
            document.getElementById('overlay').style.visibility = 'hidden';
            pdfIframe.src = data.pdf_url + "#toolbar=0";
            
            layoutConfirmed = true;
            btnGenerate.disabled = false;
            confirmedBadge.style.display = 'inline-block';
            btnConfirm.textContent = 'Re-confirm Layout';
            btnConfirm.disabled = false;
            
            document.getElementById('step1-num').classList.add('done');
            document.getElementById('step1-num').textContent = '✓';
            document.getElementById('step2-num').classList.add('done');
            document.getElementById('step2-num').textContent = '✓';
        } else {
            alert('Error: ' + (data.error || 'Unknown'));
            btnConfirm.textContent = 'Step 1: Confirm Layout';
            btnConfirm.disabled = false;
        }
    } catch (err) {
        alert('Error generating preview');
        btnConfirm.textContent = 'Step 1: Confirm Layout';
        btnConfirm.disabled = false;
    }
}

// STEP 2: Generate all invoices using the SAVED confirmed positions
async function generateBatch() {
    if (!layoutConfirmed || !confirmedPositions) {
        alert('Please confirm the layout first by clicking "Confirm Layout".');
        return;
    }
    
    // Use the saved confirmed positions, NOT live element positions
    const body = {
        excel_filename: excelFilename,
        seal_pos: confirmedPositions.seal_pos,
        sign_pos: confirmedPositions.sign_pos
    };
    
    try {
        const res = await fetch('/generate_batch', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        const data = await res.json();
        if (data.success) {
            startPolling(data.session_id);
        } else {
            alert('Error: ' + (data.error || 'Unknown'));
        }
    } catch (err) {
        alert('Error starting batch generation.');
    }
}

function startPolling(sessionId) {
    progressModal.classList.add('visible');
    const pollInterval = setInterval(async () => {
        try {
            const res = await fetch(`/batch_progress/${sessionId}`);
            const data = await res.json();
            
            if (data.total > 0) {
                const percent = (data.current / data.total) * 100;
                progressBar.style.width = `${percent}%`;
                progressText.textContent = `Invoice ${data.current} of ${data.total}`;
            }
            
            if (data.status === 'completed') {
                clearInterval(pollInterval);
                progressStatus.textContent = 'Done! Downloading ZIP...';
                progressBar.style.width = '100%';
                window.location.href = data.zip_url;
                setTimeout(() => window.location.href = '/', 3000);
            } else if (data.status === 'error') {
                clearInterval(pollInterval);
                progressStatus.textContent = 'Error occurred during generation.';
            }
        } catch (e) {
            console.error('Poll error:', e);
        }
    }, 1000);
}
