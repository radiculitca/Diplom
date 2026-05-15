// ===================== DRAG & DROP + FILE PREVIEW =====================
(function () {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('excelFile');
    if (!dropZone || !fileInput) return;

    dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
    dropZone.addEventListener('dragleave', e => { if (!dropZone.contains(e.relatedTarget)) dropZone.classList.remove('drag-over'); });
    dropZone.addEventListener('drop', e => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length) {
            const dt = new DataTransfer();
            for (const f of e.dataTransfer.files) dt.items.add(f);
            fileInput.files = dt.files;
            updateFilePreview(fileInput.files);
            updateUploadBtn();
        }
    });
    dropZone.addEventListener('click', e => {
        if (!e.target.closest('label') && !e.target.closest('input')) fileInput.click();
    });
    fileInput.addEventListener('change', () => { updateFilePreview(fileInput.files); updateUploadBtn(); });
}());

function updateFilePreview(files) {
    const preview = document.getElementById('fileListPreview');
    if (!preview) return;
    if (!files || !files.length) { preview.innerHTML = ''; return; }
    const ext = f => f.name.toLowerCase().endsWith('.csv') ? 'fa-file-csv text-secondary' : 'fa-file-excel text-success';
    preview.innerHTML = Array.from(files).map(f => `
        <div class="d-flex align-items-center gap-2 py-1 border-bottom">
            <i class="fa-solid ${ext(f)} fa-fw"></i>
            <span class="text-truncate small">${f.name}</span>
            <small class="text-muted ms-auto flex-shrink-0">${(f.size / 1024).toFixed(0)}&nbsp;КБ</small>
        </div>`).join('');
}

// ===================== UPLOAD FORM =====================
document.getElementById('uploadForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const formData = new FormData();
    const files = document.getElementById('excelFile').files;
    for (let i = 0; i < files.length; i++) formData.append('files', files[i]);

    document.getElementById('uploadBtn').disabled = true;
    document.getElementById('uploadSpinner').classList.remove('d-none');

    try {
        const response = await fetch('/upload', { method: 'POST', body: formData });
        const data = await response.json();
        if (response.ok) {
            window.uploadedFiles = data.files;
            const container = document.getElementById('sheetCheckboxesContainer');
            container.innerHTML = '';

            window.uploadedFiles.forEach((file, fIdx) => {
                container.insertAdjacentHTML('beforeend', `<div class="file-section-title">Файл ${fIdx + 1}: ${file.original_name}</div>`);
                file.sheets.forEach(sheet => {
                    const autoChecked = file.sheets.length === 1 ? ' checked' : '';
                    container.insertAdjacentHTML('beforeend', `
                        <div class="form-check ms-3">
                            <input class="form-check-input sheet-checkbox" type="checkbox" value="${sheet}" data-filename="${file.filename}" id="s_${fIdx}_${sheet}"${autoChecked}>
                            <label class="form-check-label" for="s_${fIdx}_${sheet}">${sheet}</label>
                        </div>`);
                });
            });

            goToStep(1);
            updateSheetBtn();
        } else { showToast(data.message, 'danger'); }
    } catch (err) { showToast('Ошибка соединения с сервером', 'danger'); }
    finally {
        updateUploadBtn();
        document.getElementById('uploadSpinner').classList.add('d-none');
    }
});
