// ===================== SHEET FORM =====================
document.getElementById('sheetForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const checkedBoxes = document.querySelectorAll('.sheet-checkbox:checked');
    const filesPayload = [];
    window.uploadedFiles.forEach(f => {
        const selectedSheets = Array.from(checkedBoxes).filter(cb => cb.dataset.filename === f.filename).map(cb => cb.value);
        if (selectedSheets.length > 0) {
            filesPayload.push({ filename: f.filename, sheets: selectedSheets });
        }
    });

    document.getElementById('processSheetsBtn').disabled = true;
    document.getElementById('sheetSpinner').classList.remove('d-none');

    try {
        const response = await fetch('/process_sheets', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ files: filesPayload })
        });
        const data = await response.json();

        if (response.ok) {
            window.processedFiles = data.processed_files;
            window.questionMapping = {};
            window.reportSections = [];

            renderQuestionsStep3();
            goToStep(2);
        } else { showToast(data.message, 'danger'); }
    } catch (err) { showToast('Ошибка соединения с сервером', 'danger'); }
    finally {
        updateSheetBtn();
        document.getElementById('sheetSpinner').classList.add('d-none');
    }
});

document.getElementById('selectAllSheets').addEventListener('change', (e) => {
    document.querySelectorAll('.sheet-checkbox').forEach(cb => cb.checked = e.target.checked);
    updateSheetBtn();
});

document.getElementById('sheetCheckboxesContainer').addEventListener('change', (e) => {
    if (e.target.classList.contains('sheet-checkbox')) {
        const total = document.querySelectorAll('.sheet-checkbox').length;
        const checked = document.querySelectorAll('.sheet-checkbox:checked').length;
        document.getElementById('selectAllSheets').checked = (total === checked && total > 0);
        updateSheetBtn();
    }
});

// ===================== LEGEND SETTINGS =====================
function renderLegendSettings() {
    const block = document.getElementById('legendSettingsBlock');
    const container = document.getElementById('legendInputsContainer');
    if (!window.processedFiles || window.processedFiles.length <= 1) {
        if (block) block.style.display = 'none';
        return;
    }
    if (block) block.style.display = '';
    container.innerHTML = '';
    window.processedFiles.forEach((f, i) => {
        const color = defaultColors[i % defaultColors.length];
        container.innerHTML += `
        <div class="d-flex align-items-center border p-2 rounded gap-2">
            <input type="color" class="form-control form-control-color legend-color-swatch legend-color" data-file="${f.clean_filename}" value="${color}">
            <button type="button" class="btn btn-sm btn-outline-secondary random-legend-color-btn" data-file="${f.clean_filename}" title="Случайный цвет"><i class="fa-solid fa-dice-five"></i></button>
            <input type="text" class="form-control form-control-sm legend-label-input legend-label" data-file="${f.clean_filename}" value="${f.original_name.replace(/\.[^.]+$/, '')}" placeholder="Подпись файла">
        </div>`;
    });
}
