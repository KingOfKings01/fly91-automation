document.getElementById('excelForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const formData = new FormData(e.target);
    const btn = e.target.querySelector('button');
    btn.textContent = 'Processing...';
    btn.disabled = true;

    try {
        const response = await fetch('/upload_excel', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();
        if (data.success) {
            // Redirect to preview of the first row
            window.location.href = `/preview_first?excel=${data.filename}`;
        } else {
            alert(data.error || 'Upload failed');
        }
    } catch (err) {
        alert('An error occurred');
    } finally {
        btn.textContent = 'Upload & Proceed';
        btn.disabled = false;
    }
});

async function uploadMedia(input, type) {
    if (!input.files || !input.files[0]) return;
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
            img.src = data.url;
            img.style.display = 'block';
        } else {
            alert(data.error || 'Upload failed');
        }
    } catch (err) {
        alert('An error occurred');
    }
}
