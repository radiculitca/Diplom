window.uploadedFiles = [];
window.processedFiles = [];
window.questionMapping = {};
window.questionSourceFile = {};
window.charts = {};
window.questionMerges = {}; // { "основной вопрос": "подвопрос" }
window.questionSections = []; // [{ id, name, description, color, questions: [colName, ...] }]

// ===================== WIZARD NAVIGATION =====================
let currentWizardStep = 0;

function goToStep(n) {
    const track = document.getElementById('wizardTrack');
    if (!track) return;
    track.style.transform = `translateX(-${n * 100}%)`;
    currentWizardStep = n;
    document.querySelectorAll('.step-dot').forEach((dot, i) => {
        dot.classList.toggle('active', i === n);
        dot.classList.toggle('done', i < n);
    });
    const stepEl = document.getElementById(`wizardStep${n}`);
    if (stepEl) stepEl.scrollTop = 0;
    document.getElementById('navBackBtn').classList.toggle('invisible', n === 0);
    document.querySelectorAll('.nav-step-fwd').forEach(el => {
        el.classList.toggle('d-none', parseInt(el.dataset.fwdStep) !== n);
    });
}

document.getElementById('toStep4Btn').addEventListener('click', () => goToStep(4));
document.getElementById('toStep6Btn').addEventListener('click', () => goToStep(6));

function updateUploadBtn() {
    const f = document.getElementById('excelFile');
    document.getElementById('uploadBtn').disabled = !f || !f.files.length;
}
function updateSheetBtn() {
    document.getElementById('processSheetsBtn').disabled = document.querySelectorAll('.sheet-checkbox:checked').length === 0;
}
function updateQuestionsBtn() {
    const hasQ = document.querySelectorAll('#sortableQuestionsList .question-item').length > 0;
    document.getElementById('toStep4Btn').disabled = !hasQ;
    document.getElementById('analyzeBtn').disabled = !hasQ;
}

// ===================== TOAST NOTIFICATIONS =====================
function showToast(message, type) {
    type = type || 'danger';
    const cfg = {
        danger:  { bg: '#dc3545', icon: 'fa-circle-xmark' },
        warning: { bg: '#e67e22', icon: 'fa-triangle-exclamation' },
        success: { bg: '#198754', icon: 'fa-circle-check' },
        info:    { bg: '#0d6efd', icon: 'fa-circle-info' }
    };
    const c = cfg[type] || cfg.danger;
    let container = document.getElementById('toastContainer');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toastContainer';
        container.className = 'toast-container position-fixed top-0 end-0 p-3';
        container.style.zIndex = '99999';
        document.body.appendChild(container);
    }
    const id = 'toast_' + Date.now();
    container.insertAdjacentHTML('beforeend', `
        <div id="${id}" class="toast align-items-center text-white border-0 shadow" role="alert" style="background:${c.bg};min-width:260px;">
            <div class="d-flex">
                <div class="toast-body fw-medium">
                    <i class="fa-solid ${c.icon} me-2"></i>${message}
                </div>
                <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast"></button>
            </div>
        </div>`);
    const el = document.getElementById(id);
    new bootstrap.Toast(el, { delay: 4500 }).show();
    el.addEventListener('hidden.bs.toast', () => el.remove());
}

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

const defaultColors = ['#FF0000', '#4472C4', '#70AD47', '#FFC000', '#ED7D31', '#A5A5A5', '#5B9BD5', '#C00000', '#00B050', '#7030A0'];

Chart.register(ChartDataLabels);
Chart.defaults.font.family = '"Times New Roman", Times, serif';
Chart.defaults.font.size = 12;
Chart.defaults.color = '#000';

function initTooltips() {
    var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    tooltipTriggerList.map(function (t) { return new bootstrap.Tooltip(t); });
}

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
                container.innerHTML += `<div class="file-section-title">Файл ${fIdx + 1}: ${file.original_name}</div>`;
                file.sheets.forEach(sheet => {
                    container.innerHTML += `
                        <div class="form-check ms-3">
                            <input class="form-check-input sheet-checkbox" type="checkbox" value="${sheet}" data-filename="${file.filename}" id="s_${fIdx}_${sheet}">
                            <label class="form-check-label" for="s_${fIdx}_${sheet}">${sheet}</label>
                        </div>`;
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

            renderQuestionsStep3();
            renderLegendSettings();
            renderMergeStep();
            goToStep(2);
        } else { showToast(data.message, 'danger'); }
    } catch (err) { showToast('Ошибка соединения с сервером', 'danger'); }
    finally {
        updateSheetBtn();
        document.getElementById('sheetSpinner').classList.add('d-none');
    }
});

// ===================== ШАГ 2.5: ОБЪЕДИНЕНИЕ ВОПРОСОВ =====================

function _getMergeAllCols() {
    const allCols = [];
    window.processedFiles.forEach(f => {
        f.columns.forEach(c => {
            if (!c.is_system && !allCols.find(x => x.name === c.name)) {
                allCols.push(c);
            }
        });
    });
    return allCols;
}

function renderMergeStep() {
    window.questionMerges = {};
    window.questionSections = [];
    const container = document.getElementById('mergeQuestionsList');
    container.innerHTML = '';

    const allCols = _getMergeAllCols();

    if (allCols.length === 0) {
        container.innerHTML = '<div class="text-muted p-3 text-center">Нет доступных вопросов</div>';
        return;
    }

    allCols.forEach((col, idx) => {
        const row = document.createElement('div');
        row.className = 'list-group-item d-flex align-items-center gap-2 py-2';
        row.dataset.colname = col.name;
        row.id = `merge_row_${idx}`;
        row.innerHTML = `
            <span class="section-color-dot flex-shrink-0"
                  style="width:10px;height:10px;border-radius:50%;background:transparent;flex-shrink:0;transition:background .2s;"></span>
            <span class="flex-grow-1 text-truncate fw-medium" title="${col.name}">${col.name}</span>
            <span class="merge-badge badge bg-success d-none" data-for="${col.name}"></span>
            <button type="button" class="btn btn-sm btn-outline-primary merge-add-btn flex-shrink-0"
                    data-main="${col.name}" title="Добавить подвопрос">
                <i class="fa-solid fa-plus"></i>
            </button>`;
        container.appendChild(row);
    });

    _refreshMergeRowStates();
}

function _refreshMergeRowStates() {
    const mergedAsChild = new Set(Object.values(window.questionMerges));

    document.querySelectorAll('#mergeQuestionsList .list-group-item').forEach(row => {
        const name = row.dataset.colname;
        const isChild = mergedAsChild.has(name);

        row.classList.toggle('text-muted', isChild);
        row.style.opacity = isChild ? '0.45' : '';
        row.querySelector('.merge-add-btn').classList.toggle('d-none', isChild);

        const badge = row.querySelector('.merge-badge');
        const merge = window.questionMerges[name];
        if (merge) {
            badge.textContent = merge.length > 35 ? merge.substring(0, 35) + '…' : merge;
            badge.classList.remove('d-none');
            if (!row.querySelector('.merge-remove-btn')) {
                const removeBtn = document.createElement('button');
                removeBtn.type = 'button';
                removeBtn.className = 'btn btn-sm btn-outline-danger merge-remove-btn flex-shrink-0';
                removeBtn.dataset.main = name;
                removeBtn.title = 'Убрать объединение';
                removeBtn.innerHTML = '<i class="fa-solid fa-xmark"></i>';
                row.insertBefore(removeBtn, row.querySelector('.merge-add-btn'));
            }
        } else {
            badge.classList.add('d-none');
            row.querySelector('.merge-remove-btn')?.remove();
        }
    });

    _refreshSectionColors();
}

// Покрасить точки и фон строк в соответствии с разделами
function _refreshSectionColors() {
    // Сначала сбросить все
    document.querySelectorAll('#mergeQuestionsList .list-group-item').forEach(row => {
        row.style.borderLeft = '';
        row.querySelector('.section-color-dot').style.background = 'transparent';
    });

    window.questionSections.forEach(sec => {
        sec.questions.forEach(qName => {
            const row = document.querySelector(`#mergeQuestionsList .list-group-item[data-colname="${CSS.escape(qName)}"]`);
            if (row) {
                row.style.borderLeft = `4px solid ${sec.color}`;
                row.querySelector('.section-color-dot').style.background = sec.color;
            }
        });
    });
}

// Открытие модалки выбора подвопроса
let _mergePickMainName = null;

document.getElementById('mergeQuestionsList').addEventListener('click', e => {
    const addBtn = e.target.closest('.merge-add-btn');
    if (addBtn) {
        _mergePickMainName = addBtn.dataset.main;
        document.getElementById('mergePickMainLabel').textContent = _mergePickMainName;
        document.getElementById('mergePickSearch').value = '';
        _renderMergePickList('');
        new bootstrap.Modal(document.getElementById('mergePickModal')).show();
        return;
    }
    const removeBtn = e.target.closest('.merge-remove-btn');
    if (removeBtn) {
        delete window.questionMerges[removeBtn.dataset.main];
        _refreshMergeRowStates();
    }
});

function _renderMergePickList(query) {
    const list = document.getElementById('mergePickList');
    list.innerHTML = '';
    const mergedAsChild = new Set(Object.values(window.questionMerges));
    const allCols = _getMergeAllCols();

    const filtered = allCols.filter(c => {
        if (c.name === _mergePickMainName) return false;
        if (mergedAsChild.has(c.name)) return false;
        if (Object.keys(window.questionMerges).includes(c.name)) return false;
        if (query && !c.name.toLowerCase().includes(query.toLowerCase())) return false;
        return true;
    });

    if (filtered.length === 0) {
        list.innerHTML = '<div class="text-muted p-3 text-center small">Нет доступных вопросов</div>';
        return;
    }

    filtered.forEach(col => {
        const item = document.createElement('div');
        item.className = 'list-group-item list-group-item-action py-2 merge-pick-item';
        item.dataset.colname = col.name;
        item.innerHTML = `<span class="fw-medium">${col.name}</span>`;
        item.addEventListener('click', () => {
            window.questionMerges[_mergePickMainName] = col.name;
            _refreshMergeRowStates();
            bootstrap.Modal.getInstance(document.getElementById('mergePickModal')).hide();
        });
        list.appendChild(item);
    });
}

document.getElementById('mergePickSearch').addEventListener('input', e => {
    _renderMergePickList(e.target.value);
});
document.getElementById('mergePickSearchClear').addEventListener('click', () => {
    document.getElementById('mergePickSearch').value = '';
    _renderMergePickList('');
});

document.getElementById('toStep3Btn').addEventListener('click', () => {
    _applyMergestoProcessedFiles();
    renderQuestionsStep3();
    goToStep(3);
});

function _applyMergestoProcessedFiles() {
    const childSet = new Set(Object.values(window.questionMerges));
    window.processedFiles.forEach(f => {
        f.columns.forEach(c => {
            c.merged_child = childSet.has(c.name);
        });
    });
}

// ===================== РАЗДЕЛЫ =====================

const SECTION_PALETTE = [
    '#4472C4', '#70AD47', '#FFC000', '#ED7D31',
    '#9B59B6', '#1ABC9C', '#E74C3C', '#3498DB',
    '#F39C12', '#16A085'
];

let _sectionEditId = null; // null = создание, иначе id раздела

function _nextSectionColor() {
    const used = window.questionSections.map(s => s.color);
    return SECTION_PALETTE.find(c => !used.includes(c)) || SECTION_PALETTE[window.questionSections.length % SECTION_PALETTE.length];
}

// Открыть модалку создания нового раздела
document.getElementById('addSectionBtn').addEventListener('click', () => {
    _sectionEditId = null;
    document.getElementById('sectionNameInput').value = '';
    document.getElementById('sectionDescInput').value = '';
    document.getElementById('sectionColorInput').value = _nextSectionColor();
    document.getElementById('sectionModalTitle').textContent = 'Новый раздел';
    document.getElementById('sectionDeleteBtn').classList.add('d-none');

    _renderSectionQuestionPicker(null);
    new bootstrap.Modal(document.getElementById('sectionModal')).show();
});

// Открыть модалку редактирования существующего раздела
function openEditSection(sectionId) {
    const sec = window.questionSections.find(s => s.id === sectionId);
    if (!sec) return;
    _sectionEditId = sectionId;
    document.getElementById('sectionNameInput').value = sec.name;
    document.getElementById('sectionDescInput').value = sec.description || '';
    document.getElementById('sectionColorInput').value = sec.color;
    document.getElementById('sectionModalTitle').textContent = 'Редактировать раздел';
    document.getElementById('sectionDeleteBtn').classList.remove('d-none');

    _renderSectionQuestionPicker(sec);
    new bootstrap.Modal(document.getElementById('sectionModal')).show();
}

// Рендер списка вопросов в модалке раздела
function _renderSectionQuestionPicker(existingSec) {
    const container = document.getElementById('sectionQuestionPicker');
    container.innerHTML = '';

    const mergedAsChild = new Set(Object.values(window.questionMerges));
    const allCols = _getMergeAllCols().filter(c => !mergedAsChild.has(c.name));

    // Вопросы уже занятые другими разделами (не текущим)
    const takenByOther = new Set();
    window.questionSections.forEach(s => {
        if (existingSec && s.id === existingSec.id) return;
        s.questions.forEach(q => takenByOther.add(q));
    });

    allCols.forEach(col => {
        const isTaken = takenByOther.has(col.name);
        const isChecked = existingSec ? existingSec.questions.includes(col.name) : false;
        const takenSec = isTaken
            ? window.questionSections.find(s => s.questions.includes(col.name))
            : null;

        const item = document.createElement('div');
        item.className = 'form-check py-1 border-bottom';
        item.innerHTML = `
            <input class="form-check-input section-q-cb" type="checkbox" id="sq_${col.name.replace(/\W/g,'_')}"
                   value="${col.name}" ${isChecked ? 'checked' : ''} ${isTaken ? 'disabled' : ''}>
            <label class="form-check-label ${isTaken ? 'text-muted' : ''}"
                   for="sq_${col.name.replace(/\W/g,'_')}">
                ${col.name}
                ${isTaken && takenSec ? `<span class="badge ms-1" style="background:${takenSec.color};font-size:0.65rem;">${takenSec.name}</span>` : ''}
            </label>`;
        container.appendChild(item);
    });

    // Чекбокс "Все"
    const total = container.querySelectorAll('.section-q-cb:not(:disabled)').length;
    const checked = container.querySelectorAll('.section-q-cb:not(:disabled):checked').length;
    document.getElementById('sectionSelectAll').checked = total > 0 && total === checked;
    document.getElementById('sectionSelectAll').indeterminate = checked > 0 && checked < total;
}

// "Выбрать все" в модалке раздела
document.getElementById('sectionSelectAll').addEventListener('change', e => {
    document.querySelectorAll('#sectionQuestionPicker .section-q-cb:not(:disabled)').forEach(cb => {
        cb.checked = e.target.checked;
    });
});

document.getElementById('sectionQuestionPicker').addEventListener('change', () => {
    const total = document.querySelectorAll('#sectionQuestionPicker .section-q-cb:not(:disabled)').length;
    const checked = document.querySelectorAll('#sectionQuestionPicker .section-q-cb:not(:disabled):checked').length;
    document.getElementById('sectionSelectAll').checked = total > 0 && total === checked;
    document.getElementById('sectionSelectAll').indeterminate = checked > 0 && checked < total;
});

// Сохранить раздел
document.getElementById('saveSectionBtn').addEventListener('click', () => {
    const name = document.getElementById('sectionNameInput').value.trim();
    if (!name) {
        document.getElementById('sectionNameInput').classList.add('is-invalid');
        return;
    }
    document.getElementById('sectionNameInput').classList.remove('is-invalid');

    const description = document.getElementById('sectionDescInput').value.trim();
    const color = document.getElementById('sectionColorInput').value;
    const questions = Array.from(
        document.querySelectorAll('#sectionQuestionPicker .section-q-cb:checked')
    ).map(cb => cb.value);

    if (_sectionEditId !== null) {
        // Редактирование
        const sec = window.questionSections.find(s => s.id === _sectionEditId);
        if (sec) {
            sec.name = name;
            sec.description = description;
            sec.color = color;
            sec.questions = questions;
        }
    } else {
        // Создание
        window.questionSections.push({
            id: Date.now(),
            name,
            description,
            color,
            questions
        });
    }

    bootstrap.Modal.getInstance(document.getElementById('sectionModal')).hide();
    _refreshSectionColors();
    _renderSectionChips();
});

// Удалить раздел
document.getElementById('sectionDeleteBtn').addEventListener('click', () => {
    if (_sectionEditId !== null) {
        window.questionSections = window.questionSections.filter(s => s.id !== _sectionEditId);
        bootstrap.Modal.getInstance(document.getElementById('sectionModal')).hide();
        _refreshSectionColors();
        _renderSectionChips();
    }
});

// Рендер чипов разделов над списком вопросов
function _renderSectionChips() {
    const container = document.getElementById('sectionChips');
    if (!container) return;
    container.innerHTML = '';

    if (window.questionSections.length === 0) {
        container.innerHTML = '<span class="text-muted small fst-italic">Разделы не заданы</span>';
        return;
    }

    window.questionSections.forEach(sec => {
        const chip = document.createElement('span');
        chip.className = 'badge rounded-pill d-inline-flex align-items-center gap-1 section-chip';
        chip.style.cssText = `background:${sec.color};cursor:pointer;font-size:0.8rem;padding:5px 10px;`;
        chip.title = sec.description || sec.name;
        chip.innerHTML = `
            <i class="fa-solid fa-layer-group" style="font-size:0.7rem;"></i>
            ${sec.name}
            <span class="ms-1" style="opacity:0.7;font-size:0.75rem;">${sec.questions.length}</span>`;
        chip.addEventListener('click', () => openEditSection(sec.id));
        container.appendChild(chip);
    });
}

// ===================== ШАГ 3: ВОПРОСЫ =====================

function renderQuestionsStep3() {
    window.questionMapping = {};
    window.questionSourceFile = {};
    document.getElementById('sortableQuestionsList').innerHTML = '<div class="p-3 text-center text-muted" id="emptySortablePlaceholder"><i class="fa-solid fa-hand-pointer me-1"></i>Выберите вопросы на шаге 4, чтобы они появились здесь</div>';

    const fileSelect = document.getElementById('fileSelectStep3');
    const fileSelectContainer = document.getElementById('fileSelectContainer');
    if (window.processedFiles.length > 1) {
        if ($(fileSelect).hasClass('select2-hidden-accessible')) $(fileSelect).select2('destroy');
        fileSelect.innerHTML = window.processedFiles.map((f, i) =>
            `<option value="${i}">${f.original_name}</option>`
        ).join('');
        fileSelectContainer.style.display = '';
        $(fileSelect).select2({ theme: 'bootstrap-5', width: '100%', language: 'ru'});
    } else {
        fileSelectContainer.style.display = 'none';
    }

    document.getElementById('selectAllQuestions').checked = false;
    fileSelect.value = '0';
    _renderQuestionsForFile(0);

    new Sortable(document.getElementById('sortableQuestionsList'), { handle: '.drag-handle', animation: 150, ghostClass: 'sortable-ghost' });
    updateQuestionsBtn();
}

function _getQuestionSectionColor(qName) {
    const sec = window.questionSections.find(s => s.questions.includes(qName));
    return sec ? sec.color : null;
}

function _renderQuestionsForFile(fileIdx) {
    document.getElementById('questionsSearchInput').value = '';
    const container = document.getElementById('allQuestionsList');
    container.innerHTML = '';
    const f = window.processedFiles[fileIdx];

    const selectedQNames = new Set(
        Array.from(document.querySelectorAll('#sortableQuestionsList .question-item')).map(el => el.dataset.col)
    );
    const showHidden = document.getElementById('showHiddenCols').checked;

    const inner = document.createElement('div');
    inner.style.cssText = 'width:max-content;min-width:100%';

    f.columns.forEach((colObj, colIdx) => {
        if (colObj.merged_child) return;
        const isSystem = colObj.is_system;
        const qName = colObj.name;
        const cbId = `qcb_${fileIdx}_${colIdx}`;
        const sectionColor = _getQuestionSectionColor(qName);

        const item = document.createElement('div');
        const hiddenClass = isSystem ? 'system-col' + (showHidden ? '' : ' d-none') : '';
        const labelColor = isSystem ? 'text-secondary' : 'text-dark';
        item.className = `list-group-item d-flex align-items-center gap-2 py-2 q-item-container ${hiddenClass}`;
        item.dataset.qname = qName;

        // Цветовая полоска раздела слева
        if (sectionColor) {
            item.style.borderLeft = `4px solid ${sectionColor}`;
        }

        item.innerHTML = `
            <input class="form-check-input flex-shrink-0 q-checkbox" type="checkbox" id="${cbId}" value="${qName}"${selectedQNames.has(qName) ? ' checked' : ''}>
            <label class="form-check-label fw-medium ${labelColor}" for="${cbId}" style="white-space:nowrap;cursor:pointer" title="${qName}">${qName}</label>`;
        inner.appendChild(item);
    });

    container.appendChild(inner);

    const total = container.querySelectorAll('.q-item-container:not(.d-none):not(.q-filtered)').length;
    const checked = container.querySelectorAll('.q-item-container:not(.d-none):not(.q-filtered) .q-checkbox:checked').length;
    document.getElementById('selectAllQuestions').checked = (total === checked && total > 0);
}

function filterQuestionsList() {
    const query = document.getElementById('questionsSearchInput').value.trim().toLowerCase();
    document.querySelectorAll('#allQuestionsList .q-item-container').forEach(item => {
        const matches = !query || (item.dataset.qname || '').toLowerCase().includes(query);
        item.classList.toggle('q-filtered', !matches);
    });
    const total = document.querySelectorAll('#allQuestionsList .q-item-container:not(.d-none):not(.q-filtered)').length;
    const checked = document.querySelectorAll('#allQuestionsList .q-item-container:not(.d-none):not(.q-filtered) .q-checkbox:checked').length;
    document.getElementById('selectAllQuestions').checked = (total === checked && total > 0);
}

document.getElementById('questionsSearchInput').addEventListener('input', filterQuestionsList);
document.getElementById('questionsSearchClear').addEventListener('click', () => {
    document.getElementById('questionsSearchInput').value = '';
    filterQuestionsList();
});

document.getElementById('fileSelectStep3').addEventListener('change', (e) => {
    document.querySelectorAll('#sortableQuestionsList .question-item').forEach(el => {
        delete window.questionMapping[el.dataset.col];
        delete window.questionSourceFile[el.dataset.col];
    });
    document.getElementById('sortableQuestionsList').innerHTML = '<div class="p-3 text-center text-muted" id="emptySortablePlaceholder"><i class="fa-solid fa-hand-pointer me-1"></i>Выберите вопросы на шаге 4, чтобы они появились здесь</div>';
    document.getElementById('selectAllQuestions').checked = false;
    updateQuestionsBtn();
    _renderQuestionsForFile(parseInt(e.target.value, 10));
});

function renderLegendSettings() {
    const container = document.getElementById('legendInputsContainer');
    container.innerHTML = '';
    window.processedFiles.forEach((f, i) => {
        const color = defaultColors[i % defaultColors.length];
        container.innerHTML += `
        <div class="d-flex align-items-center border p-2 rounded gap-2">
            <input type="color" class="form-control form-control-color legend-color" data-file="${f.clean_filename}" value="${color}" style="width: 36px; height: 30px; padding: 1px 2px; cursor: pointer;">
            <button type="button" class="btn btn-sm btn-outline-secondary random-legend-color-btn" data-file="${f.clean_filename}" title="Случайный цвет"><i class="fa-solid fa-dice-five"></i></button>
            <input type="text" class="form-control form-control-sm legend-label" data-file="${f.clean_filename}" value="${f.original_name.replace(/\.[^.]+$/, '')}" placeholder="Подпись файла" style="flex:1;min-width:200px;">
        </div>`;
    });
}

function autoMapQuestion(qName, sourceFileIdx) {
    sourceFileIdx = sourceFileIdx ?? 0;
    const mapping = {};
    mapping[window.processedFiles[sourceFileIdx].clean_filename] = qName;
    window.questionSourceFile[qName] = sourceFileIdx;

    const missingIn = [];
    for (let i = 0; i < window.processedFiles.length; i++) {
        if (i === sourceFileIdx) continue;
        const f = window.processedFiles[i];
        const found = f.columns.find(c => c.name === qName);
        if (found) {
            mapping[f.clean_filename] = found.name;
        } else {
            missingIn.push(f.original_name);
        }
    }

    window.questionMapping[qName] = mapping;
    return missingIn;
}

function addQuestionToSortable(qName, sourceFileIdx) {
    const sortableContainer = document.getElementById('sortableQuestionsList');
    const alreadyExists = Array.from(sortableContainer.querySelectorAll('.question-item')).some(el => el.getAttribute('data-col') === qName);
    if (alreadyExists) return;

    const missingIn = autoMapQuestion(qName, sourceFileIdx ?? 0);
    const sectionColor = _getQuestionSectionColor(qName);

    const warningHtml = missingIn.length > 0
        ? `<span class="missing-warning ms-2" data-bs-toggle="tooltip" title="Вопрос не найден в: ${missingIn.join(', ')}"><i class="fa-solid fa-circle-exclamation"></i></span>`
        : '';

    const mappingBtnHtml = window.processedFiles.length > 1
        ? `<a class="mapping-btn ms-2" title="Соотнести вручную" onclick="openMappingModal('${qName}')"><i class="fa-solid fa-link"></i></a>`
        : '';

    const colorDot = sectionColor
        ? `<span style="width:8px;height:8px;border-radius:50%;background:${sectionColor};flex-shrink:0;display:inline-block;margin-right:4px;"></span>`
        : '';

    const item = document.createElement('div');
    item.className = 'list-group-item d-flex align-items-center question-item bg-white';
    item.setAttribute('data-col', qName);
    if (sectionColor) item.style.borderLeft = `4px solid ${sectionColor}`;
    item.innerHTML = `
        <span class="drag-handle me-3" title="Потяните, чтобы переместить"><i class="fa-solid fa-grip-lines"></i></span>
        ${colorDot}
        <span class="text-truncate fw-medium text-dark flex-grow-1" title="${qName}">${qName}</span>
        ${warningHtml}
        ${mappingBtnHtml}`;

    sortableContainer.appendChild(item);

    initTooltips();
    checkEmptyPlaceholder();
    updateQuestionsBtn();
}

function removeQuestionFromSortable(qName) {
    document.getElementById('sortableQuestionsList').querySelectorAll('.question-item').forEach(el => {
        if (el.getAttribute('data-col') === qName) el.remove();
    });
    delete window.questionMapping[qName];
    delete window.questionSourceFile[qName];
    checkEmptyPlaceholder();
    updateQuestionsBtn();
}

function checkEmptyPlaceholder() {
    const sortableContainer = document.getElementById('sortableQuestionsList');
    const emptyPlaceholder = document.getElementById('emptySortablePlaceholder');
    const hasItems = sortableContainer.querySelectorAll('.question-item').length > 0;
    if (emptyPlaceholder) emptyPlaceholder.style.display = hasItems ? 'none' : 'block';
}

let currentMappingQName = null;
window.openMappingModal = function (qName) {
    currentMappingQName = qName;
    const mapping = window.questionMapping[qName] || {};
    const sourceFileIdx = window.questionSourceFile[qName] ?? 0;
    const sourceFile = window.processedFiles[sourceFileIdx];
    const body = document.getElementById('mappingModalBody');
    body.innerHTML = `<p class="fw-bold mb-3">Соотнести вопрос:<br><span class="text-primary">${qName}</span><br><small class="text-muted fw-normal">Источник: ${sourceFile.original_name}</small></p>`;

    for (let i = 0; i < window.processedFiles.length; i++) {
        if (i === sourceFileIdx) continue;
        const f = window.processedFiles[i];
        const currentMapped = mapping[f.clean_filename] || "";

        let optionsHtml = `<option value="">-- Не выбрано (Исключить) --</option>`;
        f.columns.forEach(c => {
            const sel = c.name === currentMapped ? 'selected' : '';
            optionsHtml += `<option value="${c.name}" ${sel}>${c.name}</option>`;
        });

        body.innerHTML += `
            <div class="mb-3">
                <label class="form-label small text-muted fw-semibold">В файле: ${f.original_name}</label>
                <select class="mapping-select" data-file="${f.clean_filename}">${optionsHtml}</select>
            </div>`;
    }

    const modalEl = document.getElementById('dataMappingModal');
    $('.mapping-select').select2({
        theme: 'bootstrap-5',
        width: '100%',
        dropdownParent: $(modalEl),
        language: 'ru',
        placeholder: '-- Не выбрано (Исключить) --',
        allowClear: true
    });

    new bootstrap.Modal(modalEl).show();
};

document.getElementById('saveMappingBtn').addEventListener('click', () => {
    const selects = document.querySelectorAll('.mapping-select');
    const sourceFileIdx = window.questionSourceFile[currentMappingQName] ?? 0;
    const sourceFile = window.processedFiles[sourceFileIdx];
    const mapping = { [sourceFile.clean_filename]: currentMappingQName };
    let missingCount = 0;

    selects.forEach(sel => {
        if (sel.value) mapping[sel.getAttribute('data-file')] = sel.value;
        else missingCount++;
    });

    window.questionMapping[currentMappingQName] = mapping;

    const qItem = document.querySelector(`.question-item[data-col="${currentMappingQName}"]`);
    if (qItem) {
        let warnIcon = qItem.querySelector('.missing-warning');
        if (missingCount > 0) {
            if (!warnIcon) {
                qItem.querySelector('.text-truncate').insertAdjacentHTML('afterend', `<span class="missing-warning ms-2" data-bs-toggle="tooltip" title="Вопрос не соотнесен во всех файлах"><i class="fa-solid fa-circle-exclamation"></i></span>`);
                initTooltips();
            }
        } else {
            if (warnIcon) warnIcon.remove();
        }
    }

    bootstrap.Modal.getInstance(document.getElementById('dataMappingModal')).hide();
});

document.getElementById('allQuestionsList').addEventListener('click', (e) => {
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'LABEL') return;
    const container = e.target.closest('.q-item-container');
    if (!container) return;
    const cb = container.querySelector('.q-checkbox');
    if (cb) cb.click();
});

document.getElementById('allQuestionsList').addEventListener('change', (e) => {
    if (e.target.classList.contains('q-checkbox')) {
        const sourceFileIdx = parseInt(document.getElementById('fileSelectStep3').value || '0', 10);
        if (e.target.checked) addQuestionToSortable(e.target.value, sourceFileIdx);
        else removeQuestionFromSortable(e.target.value);

        const total = document.querySelectorAll('.q-item-container:not(.d-none):not(.q-filtered)').length;
        const checked = document.querySelectorAll('.q-item-container:not(.d-none):not(.q-filtered) .q-checkbox:checked').length;
        document.getElementById('selectAllQuestions').checked = (total === checked && total > 0);
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

document.getElementById('selectAllQuestions').addEventListener('change', (e) => {
    const isChecked = e.target.checked;
    const sourceFileIdx = parseInt(document.getElementById('fileSelectStep3').value || '0', 10);
    document.querySelectorAll('.q-item-container:not(.d-none):not(.q-filtered)').forEach(el => {
        const cb = el.querySelector('.q-checkbox');
        if (cb.checked !== isChecked) {
            cb.checked = isChecked;
            if (isChecked) addQuestionToSortable(cb.value, sourceFileIdx);
            else removeQuestionFromSortable(cb.value);
        }
    });
});

document.getElementById('showHiddenCols').addEventListener('change', (e) => {
    const isChecked = e.target.checked;
    document.querySelectorAll('#allQuestionsList .system-col').forEach(el => {
        if (isChecked) el.classList.remove('d-none');
        else {
            el.classList.add('d-none');
            const cb = el.querySelector('.q-checkbox');
            if (cb && cb.checked) { cb.checked = false; removeQuestionFromSortable(cb.value); }
        }
    });
    const total = document.querySelectorAll('.q-item-container:not(.d-none)').length;
    const checked = document.querySelectorAll('.q-item-container:not(.d-none) .q-checkbox:checked').length;
    document.getElementById('selectAllQuestions').checked = (total === checked && total > 0);
});

document.getElementById('sortableQuestionsList').addEventListener('change', (e) => {
    if (e.target.classList.contains('viz-all-cb')) {
        e.target.closest('.dropdown-menu').querySelectorAll('.viz-opt-cb').forEach(cb => cb.checked = e.target.checked);
    }
    if (e.target.classList.contains('viz-opt-cb')) {
        const dp = e.target.closest('.dropdown-menu');
        dp.querySelector('.viz-all-cb').checked = (dp.querySelectorAll('.viz-opt-cb').length === dp.querySelectorAll('.viz-opt-cb:checked').length);
    }
});

// ===================== FUZZY / RANGES =====================

function levenshtein(a, b) {
    if (a === b) return 0;
    if (a.length === 0) return b.length;
    if (b.length === 0) return a.length;
    const dp = Array.from({length: a.length + 1}, (_, i) => [i]);
    for (let j = 1; j <= b.length; j++) dp[0][j] = j;
    for (let i = 1; i <= a.length; i++) {
        for (let j = 1; j <= b.length; j++) {
            dp[i][j] = a[i-1] === b[j-1]
                ? dp[i-1][j-1]
                : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
        }
    }
    return dp[a.length][b.length];
}

const FUZZY_THRESHOLD = 2;
const FUZZY_MIN_LEN = 4;
const RU_PATTERN = /[а-яёА-ЯЁ]/;

function toSlug(str) {
    let s = str.replace(/[\p{Emoji_Presentation}\p{Extended_Pictographic}]/gu, '').trim();
    s = s.replace(/^из\s+/i, '').trim();
    return window.transliterate ? window.transliterate(s).toLowerCase().replace(/[^a-z0-9]/g, '') : s.toLowerCase().replace(/[^a-z0-9а-яё]/g, '');
}

let _fuzzyTargetId = null;
let _pendingMergeData = null;

function _escHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function _fuzzyResetPreview() {
    _pendingMergeData = null;
    document.getElementById('fuzzyConfirmBtn').disabled = true;
    document.getElementById('fuzzyPreviewArea').classList.add('d-none');
    document.getElementById('fuzzyPreviewContent').innerHTML = '';
}

function _fuzzyShowPreview(data, origAnswers) {
    _pendingMergeData = data;
    document.getElementById('fuzzyConfirmBtn').disabled = false;
    const html = data.map(row => {
        const from = origAnswers && origAnswers[row.answer];
        const merged = from && from.length
            ? ` <small class="text-muted">(+ ${from.slice(0, 3).map(a => `«${_escHtml(a)}»`).join(', ')}${from.length > 3 ? ` и ещё ${from.length - 3}` : ''})</small>`
            : '';
        return `<div class="d-flex justify-content-between align-items-baseline py-1 border-bottom">
            <span>${_escHtml(row.answer)}${merged}</span>
            <span class="badge bg-secondary ms-2 flex-shrink-0">${row._total}</span>
        </div>`;
    }).join('');
    document.getElementById('fuzzyPreviewContent').innerHTML = html || '<span class="text-muted">Нет данных.</span>';
    document.getElementById('fuzzyPreviewArea').classList.remove('d-none');
}

window.confirmFuzzyMapping = function(id) {
    _fuzzyTargetId = id;
    if (document.activeElement) document.activeElement.blur();
    document.querySelectorAll('#fuzzyModalTabs [data-fuzzy-tab]').forEach(b => b.classList.remove('active'));
    document.querySelector('#fuzzyModalTabs [data-fuzzy-tab="fuzzy"]').classList.add('active');
    document.getElementById('fuzzyTabPane').classList.remove('d-none');
    document.getElementById('rangesTabPane').classList.add('d-none');
    _fuzzyResetPreview();
    new bootstrap.Modal(document.getElementById('fuzzyConfirmModal')).show();
};

document.getElementById('fuzzyModalTabs').addEventListener('click', e => {
    const btn = e.target.closest('[data-fuzzy-tab]');
    if (!btn) return;
    document.querySelectorAll('[data-fuzzy-tab]').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const tab = btn.dataset.fuzzyTab;
    document.getElementById('fuzzyTabPane').classList.toggle('d-none', tab !== 'fuzzy');
    document.getElementById('rangesTabPane').classList.toggle('d-none', tab !== 'ranges');
    _fuzzyResetPreview();
});

document.getElementById('fuzzyPreviewBtn').addEventListener('click', () => {
    if (!_fuzzyTargetId) return;
    const result = _computeFuzzyMerge(_fuzzyTargetId);
    if (result) _fuzzyShowPreview(result.data, result.origAnswers);
});

document.getElementById('rangesPreviewBtn').addEventListener('click', () => {
    if (!_fuzzyTargetId) return;
    const n = parseInt(document.getElementById('rangeCountInput').value, 10);
    if (n < 2) return;
    const result = _computeRangeMerge(_fuzzyTargetId, n);
    if (result) {
        _fuzzyShowPreview(result.data, null);
    } else {
        document.getElementById('fuzzyPreviewContent').innerHTML = '<span class="text-danger">Нет числовых ответов для разбиения на диапазоны.</span>';
        document.getElementById('fuzzyPreviewArea').classList.remove('d-none');
        _pendingMergeData = null;
        document.getElementById('fuzzyConfirmBtn').disabled = true;
    }
});

document.getElementById('manualRangesBtn').addEventListener('click', () => {
    if (!_fuzzyTargetId) return;
    const n = parseInt(document.getElementById('rangeCountInput').value, 10);
    if (!(n >= 2)) return;

    const dataObj = window.appData[_fuzzyTargetId];
    if (!dataObj) return;

    const activeRows = dataObj.data.filter(r => r.included !== false);
    const numericVals = activeRows.map(r => parseFloat(r.answer)).filter(v => !isNaN(v));

    let preRanges;
    if (numericVals.length >= 2) {
        const minVal = Math.min(...numericVals);
        const maxVal = Math.max(...numericVals);
        const step = minVal === maxVal ? 1 : (maxVal - minVal) / n;
        preRanges = Array.from({ length: n }, (_, i) => ({
            lo: Math.round(minVal + i * step),
            hi: Math.round(i === n - 1 ? maxVal : minVal + (i + 1) * step)
        }));
    } else {
        preRanges = Array.from({ length: n }, (_, i) => ({ lo: i, hi: i + 1 }));
    }

    document.getElementById('manualRangesBody').innerHTML = preRanges.map((r, i) => `
        <div class="d-flex align-items-center gap-2 mb-2">
            <span class="text-muted small" style="min-width:90px">Диапазон ${i + 1}:</span>
            <input type="number" class="form-control form-control-sm manual-range-lo" value="${r.lo}" style="width:90px">
            <span class="text-muted">—</span>
            <input type="number" class="form-control form-control-sm manual-range-hi" value="${r.hi}" style="width:90px">
        </div>`).join('');

    new bootstrap.Modal(document.getElementById('manualRangesModal')).show();
});

document.getElementById('applyManualRangesBtn').addEventListener('click', () => {
    if (!_fuzzyTargetId) return;

    const loInputs = [...document.querySelectorAll('#manualRangesBody .manual-range-lo')];
    const hiInputs = [...document.querySelectorAll('#manualRangesBody .manual-range-hi')];

    loInputs.forEach(el => el.classList.remove('is-invalid'));
    hiInputs.forEach(el => el.classList.remove('is-invalid'));

    const customRanges = loInputs.map((el, i) => ({
        lo: parseFloat(el.value),
        hi: parseFloat(hiInputs[i].value),
        isLast: i === loInputs.length - 1
    }));

    if (customRanges.some(r => isNaN(r.lo) || isNaN(r.hi))) return;

    const invalidIdx = new Set();
    customRanges.forEach((r, i) => { if (r.lo > r.hi) invalidIdx.add(i); });
    for (let i = 0; i < customRanges.length; i++) {
        for (let j = i + 1; j < customRanges.length; j++) {
            const a = customRanges[i], b = customRanges[j];
            if (a.lo <= b.hi && b.lo <= a.hi) { invalidIdx.add(i); invalidIdx.add(j); }
        }
    }
    if (invalidIdx.size > 0) {
        invalidIdx.forEach(i => { loInputs[i].classList.add('is-invalid'); hiInputs[i].classList.add('is-invalid'); });
        showToast('Диапазоны пересекаются или неверно заданы. Исправьте выделенные поля.', 'danger');
        return;
    }

    bootstrap.Modal.getInstance(document.getElementById('manualRangesModal')).hide();

    const result = _computeRangeMergeCustom(_fuzzyTargetId, customRanges);
    if (result) {
        _fuzzyShowPreview(result.data, null);
    } else {
        document.getElementById('fuzzyPreviewContent').innerHTML = '<span class="text-danger">Нет числовых ответов для разбиения на диапазоны.</span>';
        document.getElementById('fuzzyPreviewArea').classList.remove('d-none');
    }
});

document.getElementById('manualRangesBody').addEventListener('input', e => {
    if (e.target.classList.contains('manual-range-lo') || e.target.classList.contains('manual-range-hi')) {
        e.target.classList.remove('is-invalid');
    }
});

document.getElementById('fuzzyConfirmBtn').addEventListener('click', () => {
    bootstrap.Modal.getInstance(document.getElementById('fuzzyConfirmModal')).hide();
    if (_pendingMergeData && _fuzzyTargetId) _applyMergedData(_fuzzyTargetId, _pendingMergeData);
});

function _computeFuzzyMerge(id) {
    const dataObj = window.appData[id];
    if (!dataObj || dataObj.data.length === 0) return null;

    const sorted = [...dataObj.data].sort((a, b) => b._total - a._total);
    const canonicals = [];
    const rowToCanon = new Map();

    for (const row of sorted) {
        const slug = toSlug(row.answer);
        const isoCode = window.getCountryISO ? window.getCountryISO(row.answer) : null;
        let matched = null;

        for (const canon of canonicals) {
            const isoMatch = isoCode && canon.isoCode && isoCode === canon.isoCode;
            const slugMatch = slug.length >= FUZZY_MIN_LEN && canon.slug.length >= FUZZY_MIN_LEN &&
                levenshtein(slug, canon.slug) <= FUZZY_THRESHOLD;
            if (isoMatch || slugMatch) { matched = canon; break; }
        }

        if (matched) {
            if (isoCode && !matched.isoCode) matched.isoCode = isoCode;
            rowToCanon.set(row.answer, matched.answer);
            if (RU_PATTERN.test(row.answer) && !RU_PATTERN.test(matched.answer)) {
                const oldAnswer = matched.answer;
                matched.answer = row.answer;
                matched.slug = slug;
                for (const [k, v] of rowToCanon.entries()) {
                    if (v === oldAnswer) rowToCanon.set(k, row.answer);
                }
                rowToCanon.set(row.answer, row.answer);
            }
        } else {
            canonicals.push({ slug, answer: row.answer, isoCode });
            rowToCanon.set(row.answer, row.answer);
        }
    }

    if (window.getCountryRuName) {
        for (const canon of canonicals) {
            if (!canon.isoCode) continue;
            const ruName = window.getCountryRuName(canon.isoCode);
            if (ruName && canon.answer !== ruName) {
                const oldAnswer = canon.answer;
                canon.answer = ruName;
                for (const [k, v] of rowToCanon.entries()) {
                    if (v === oldAnswer) rowToCanon.set(k, ruName);
                }
            }
        }
    }

    const mergedMap = {};
    const origAnswers = {};
    for (const row of dataObj.data) {
        const canon = rowToCanon.get(row.answer);
        if (!mergedMap[canon]) {
            mergedMap[canon] = {
                answer: canon,
                counts: Object.fromEntries(dataObj.file_keys.map(fk => [fk, 0])),
                included: false,
                _total: 0
            };
            origAnswers[canon] = [];
        }
        const m = mergedMap[canon];
        dataObj.file_keys.forEach(fk => { m.counts[fk] += row.counts[fk] || 0; });
        m._total += row._total;
        m.included = m.included || row.included;
        if (row.answer !== canon) origAnswers[canon].push(row.answer);
    }

    const data = Object.values(mergedMap).sort((a, b) => b._total - a._total);
    return { data, origAnswers };
}

function _computeRangeMerge(id, numRanges) {
    const dataObj = window.appData[id];
    if (!dataObj || dataObj.data.length === 0) return null;

    const activeRows = dataObj.data.filter(r => r.included !== false);
    const numeric = [];
    const other = [];
    for (const row of activeRows) {
        const v = parseFloat(row.answer);
        if (!isNaN(v) && String(row.answer).trim() !== '') numeric.push({ row, v });
        else other.push(row);
    }
    if (numeric.length === 0) return null;

    const minVal = Math.min(...numeric.map(x => x.v));
    const maxVal = Math.max(...numeric.map(x => x.v));
    const step = minVal === maxVal ? 1 : (maxVal - minVal) / numRanges;

    const buckets = Array.from({ length: numRanges }, (_, i) => {
        const lo = minVal + i * step;
        const hi = i === numRanges - 1 ? maxVal : minVal + (i + 1) * step;
        return {
            answer: `${Math.round(lo)} – ${Math.round(hi)}`,
            lo, hi, isLast: i === numRanges - 1,
            counts: Object.fromEntries(dataObj.file_keys.map(fk => [fk, 0])),
            included: true, _total: 0
        };
    });

    for (const { row, v } of numeric) {
        const bucket = buckets.find(b => b.isLast ? v <= b.hi : v >= b.lo && v < b.hi) || buckets[buckets.length - 1];
        dataObj.file_keys.forEach(fk => { bucket.counts[fk] += row.counts[fk] || 0; });
        bucket._total += row._total;
    }

    const data = buckets.map(({ answer, counts, included, _total }) => ({ answer, counts, included, _total }));
    if (other.length > 0) {
        const otherRow = { answer: 'Другое', counts: Object.fromEntries(dataObj.file_keys.map(fk => [fk, 0])), included: true, _total: 0 };
        for (const row of other) {
            dataObj.file_keys.forEach(fk => { otherRow.counts[fk] += row.counts[fk] || 0; });
            otherRow._total += row._total;
        }
        data.push(otherRow);
    }
    return { data };
}

function _computeRangeMergeCustom(id, customRanges) {
    const dataObj = window.appData[id];
    if (!dataObj || dataObj.data.length === 0) return null;

    const activeRows = dataObj.data.filter(r => r.included !== false);
    const numeric = [];
    const other = [];
    for (const row of activeRows) {
        const v = parseFloat(row.answer);
        if (!isNaN(v) && String(row.answer).trim() !== '') numeric.push({ row, v });
        else other.push(row);
    }
    if (numeric.length === 0) return null;

    const buckets = customRanges.map(r => ({
        answer: `${r.lo} – ${r.hi}`,
        lo: r.lo, hi: r.hi, isLast: r.isLast,
        counts: Object.fromEntries(dataObj.file_keys.map(fk => [fk, 0])),
        included: true, _total: 0
    }));

    const unmatched = [];
    for (const { row, v } of numeric) {
        const bucket = buckets.find(b => v >= b.lo && v <= b.hi);
        if (bucket) {
            dataObj.file_keys.forEach(fk => { bucket.counts[fk] += row.counts[fk] || 0; });
            bucket._total += row._total;
        } else { unmatched.push(row); }
    }

    const data = buckets.map(({ answer, counts, included, _total }) => ({ answer, counts, included, _total }));
    const leftover = [...other, ...unmatched];
    if (leftover.length > 0) {
        const otherRow = { answer: 'Другое', counts: Object.fromEntries(dataObj.file_keys.map(fk => [fk, 0])), included: true, _total: 0 };
        for (const row of leftover) {
            dataObj.file_keys.forEach(fk => { otherRow.counts[fk] += row.counts[fk] || 0; });
            otherRow._total += row._total;
        }
        data.push(otherRow);
    }
    return { data };
}

function _applyMergedData(id, newData) {
    const dataObj = window.appData[id];
    dataObj.data = newData;
    renderTable(id);
    drawChart(id);
    drawStackedChart(id);
    drawPieChart(id);
    if (window._chartEditId === id && document.getElementById('chartEditModal').classList.contains('show')) {
        renderChartEditModal(id);
    }
}

function applyFuzzyMapping(id) {
    const result = _computeFuzzyMerge(id);
    if (result) _applyMergedData(id, result.data);
}

// ===================== CHART EDIT MODAL =====================

function renderChartEditModal(id) {
    const dataObj = window.appData[id];
    const body = document.getElementById('chartEditModalBody');

    const activeRows = dataObj.data.filter(r => r.included);
    const totals = {};
    dataObj.file_keys.forEach(fk => {
        totals[fk] = activeRows.reduce((sum, r) => sum + (r.counts[fk] || 0), 0);
    });

    const multiFile = dataObj.file_keys.length > 1;
    const subheaderRow = multiFile
        ? `<tr><th></th><th></th>${dataObj.file_keys.map(() => '<th class="text-center">Кол-во</th><th class="text-center">%</th>').join('')}</tr>`
        : '';
    const theadCols = multiFile
        ? dataObj.file_keys.map(fk => `<th class="text-center" colspan="2">${dataObj.file_labels[fk]}</th>`).join('')
        : '<th class="text-center">Кол-во</th><th class="text-center">%</th>';

    let tbodyHtml = '';
    dataObj.data.forEach((row, idx) => {
        const trClass = row.included ? '' : 'row-excluded';
        let tdHtml = '';
        dataObj.file_keys.forEach(fk => {
            const c = row.counts[fk] || 0;
            const tot = totals[fk];
            let pct = '0';
            if (row.included && tot > 0 && c > 0) {
                const rawPct = (c / tot) * 100;
                pct = (rawPct > 0 && rawPct < 1) ? '<1' : Math.round(rawPct).toString();
            }
            tdHtml += `
                <td class="text-center align-middle"><span contenteditable="true" class="editable-cell modal-answer-count" data-id="${id}" data-index="${idx}" data-file="${fk}">${c}</span></td>
                <td class="text-center align-middle">${pct}</td>`;
        });
        tbodyHtml += `
            <tr class="${trClass}">
                <td class="control-col">
                    <div class="row-toggle-container" style="opacity:1;">
                        <div class="form-check form-switch mb-0">
                            <input class="form-check-input modal-row-toggle" type="checkbox" ${row.included ? 'checked' : ''} data-id="${id}" data-index="${idx}">
                        </div>
                    </div>
                </td>
                <td class="text-start align-middle"><span contenteditable="true" class="editable-cell modal-answer-text" data-id="${id}" data-index="${idx}">${row.answer}</span></td>
                ${tdHtml}
            </tr>`;
    });

    body.innerHTML = `
        <table class="table table-bordered table-hover table-custom-border align-middle mb-0">
            <thead class="table-light">
                <tr>
                    <th class="control-col"></th>
                    <th class="text-start">Ответ</th>
                    ${theadCols}
                </tr>
                ${subheaderRow}
            </thead>
            <tbody>${tbodyHtml}</tbody>
        </table>`;
}

window.openChartEditModal = function(id) {
    window._chartEditId = id;
    document.getElementById('chartEditModalLabel').textContent = `Редактирование: «${window.appData[id].question_name}»`;
    renderChartEditModal(id);
    bootstrap.Modal.getOrCreateInstance(document.getElementById('chartEditModal'), { backdrop: false, keyboard: true }).show();
};

(function () {
    const modalEl = document.getElementById('chartEditModal');
    const dialog = modalEl.querySelector('.modal-dialog');
    const header = modalEl.querySelector('.modal-header');

    modalEl.addEventListener('shown.bs.modal', () => {
        document.body.classList.remove('modal-open');
        document.body.style.overflow = '';
        document.body.style.paddingRight = '';
    });

    document.addEventListener('mousedown', (e) => {
        if (!modalEl.classList.contains('show')) return;
        if (!dialog.contains(e.target)) {
            bootstrap.Modal.getInstance(modalEl)?.hide();
        }
    });

    header.style.cursor = 'grab';
    let dragging = false, ox = 0, oy = 0;

    header.addEventListener('mousedown', (e) => {
        if (e.target.closest('button')) return;
        dragging = true;
        header.style.cursor = 'grabbing';
        const rect = dialog.getBoundingClientRect();
        dialog.style.width = rect.width + 'px';
        dialog.style.position = 'fixed';
        dialog.style.margin = '0';
        dialog.style.left = rect.left + 'px';
        dialog.style.top = rect.top + 'px';
        ox = e.clientX - rect.left;
        oy = e.clientY - rect.top;
        e.preventDefault();
    });

    document.addEventListener('mousemove', (e) => {
        if (!dragging) return;
        dialog.style.left = (e.clientX - ox) + 'px';
        dialog.style.top = (e.clientY - oy) + 'px';
    });

    document.addEventListener('mouseup', () => {
        if (dragging) { dragging = false; header.style.cursor = 'grab'; }
    });
}());

// ===================== HIDE COL MODAL =====================

let currentHideColTableId = null;
window.openHideColModal = function(id) {
    currentHideColTableId = id;
    const opt = window.appData[id].options.hiddenCol || 'none';
    document.getElementById('hideColNone').checked = (opt === 'none');
    document.getElementById('hideColCount').checked = (opt === 'count');
    document.getElementById('hideColPct').checked = (opt === 'percent');
    new bootstrap.Modal(document.getElementById('hideColModal')).show();
};

document.getElementById('saveHideColBtn').addEventListener('click', () => {
    let opt = 'none';
    if (document.getElementById('hideColCount').checked) opt = 'count';
    if (document.getElementById('hideColPct').checked) opt = 'percent';
    window.appData[currentHideColTableId].options.hiddenCol = opt;
    applyHiddenColumns(currentHideColTableId);
    bootstrap.Modal.getInstance(document.getElementById('hideColModal')).hide();
});

window.applyHiddenColumns = function(id) {
    const table = document.getElementById(id);
    if (!table) return;
    const opt = window.appData[id].options.hiddenCol || 'none';
    const isVert = window.appData[id].options.tableVertical;

    table.querySelectorAll(`.count-col-${id}`).forEach(el => el.classList.toggle('d-none', opt === 'count'));
    table.querySelectorAll(`.pct-col-${id}`).forEach(el => el.classList.toggle('d-none', opt === 'percent'));

    const fileHeaders = table.querySelectorAll(`.file-header-${id}`);
    const colspanVal = opt === 'none' ? '2' : '1';
    fileHeaders.forEach(el => el.setAttribute('colspan', colspanVal));

    if (!isVert) {
        const row1Ths = table.querySelectorAll(`.main-th-${id}`);
        const row2 = document.getElementById(`thead_row2_${id}`);
        if (opt === 'none') {
            if(row2) row2.classList.remove('d-none');
            row1Ths.forEach(el => el.setAttribute('rowspan', '2'));
        } else {
            if(row2) row2.classList.add('d-none');
            row1Ths.forEach(el => el.setAttribute('rowspan', '1'));
        }
    }
};

// ===================== MERGE ANSWER MODAL =====================

let _mergeAnswerTableId = null;
let _mergeAnswerSourceIdx = null;
let _mergeAnswerTargetIdx = null;

window.openMergeAnswerModal = function(tableId, sourceIdx) {
    _mergeAnswerTableId = tableId;
    _mergeAnswerSourceIdx = sourceIdx;
    _mergeAnswerTargetIdx = null;

    const dataObj = window.appData[tableId];
    const sourceRow = dataObj.data[sourceIdx];

    document.getElementById('mergeAnswerSourceLabel').textContent =
        sourceRow.answer.length > 60 ? sourceRow.answer.substring(0, 60) + '…' : sourceRow.answer;
    document.getElementById('mergeAnswerConfirmBtn').disabled = true;

    const list = document.getElementById('mergeAnswerPickList');
    list.innerHTML = '';

    dataObj.data.forEach((row, idx) => {
        if (idx === sourceIdx) return; // сам себя нельзя
        const item = document.createElement('div');
        item.className = 'list-group-item list-group-item-action py-2 d-flex align-items-center gap-2';
        item.dataset.idx = idx;
        item.innerHTML = `
            <span class="flex-grow-1">${row.answer}</span>
            <span class="badge bg-secondary">${row._total}</span>`;
        item.addEventListener('click', () => {
            list.querySelectorAll('.list-group-item').forEach(el => el.classList.remove('active'));
            item.classList.add('active');
            _mergeAnswerTargetIdx = idx;
            document.getElementById('mergeAnswerConfirmBtn').disabled = false;
        });
        list.appendChild(item);
    });

    new bootstrap.Modal(document.getElementById('mergeAnswerModal')).show();
};

document.getElementById('mergeAnswerConfirmBtn').addEventListener('click', () => {
    if (_mergeAnswerTargetIdx === null || _mergeAnswerTableId === null) return;

    const dataObj = window.appData[_mergeAnswerTableId];
    const src = dataObj.data[_mergeAnswerSourceIdx];
    const tgt = dataObj.data[_mergeAnswerTargetIdx];

    // Суммируем counts
    dataObj.file_keys.forEach(fk => {
        tgt.counts[fk] = (tgt.counts[fk] || 0) + (src.counts[fk] || 0);
    });
    tgt._total += src._total;

    // Удаляем исходную строку
    dataObj.data.splice(_mergeAnswerSourceIdx, 1);

    bootstrap.Modal.getInstance(document.getElementById('mergeAnswerModal')).hide();

    renderTable(_mergeAnswerTableId);
    drawChart(_mergeAnswerTableId);
    drawStackedChart(_mergeAnswerTableId);
    drawPieChart(_mergeAnswerTableId);
});

// ===================== PIE / BAR / STACKED CHARTS =====================

const PIE_COLORS = [
    '#dc3545','#0d6efd','#198754','#ffc107','#6f42c1',
    '#fd7e14','#20c997','#0dcaf0','#6c757d','#343a40',
    '#e15759','#4e79a7','#59a14f','#edc948','#b07aa1',
    '#003f5c','#2f4b7c','#665191','#a05195','#d45087',
    '#f95d6a','#ff7c43','#ffa600','#b5bd00','#00b050'
];

function drawPieChart(id) {
    const dataObj = window.appData[id];
    if (!dataObj || !window.pieChartsData[id]) return;

    const MAX_PIE = 25;
    const activeData = dataObj.data.filter(r => r.included);
    const labels = activeData.map(r => r.answer.length > 40 ? r.answer.substring(0, 40) + '…' : r.answer);

    while (dataObj.pieColors.length < activeData.length) {
        dataObj.pieColors.push(PIE_COLORS[dataObj.pieColors.length % PIE_COLORS.length]);
    }

    const editor = document.getElementById(`pie_color_editor_${id}`);
    if (editor) {
        editor.innerHTML = `<div class="d-flex flex-wrap gap-1 align-items-center ui-system-font mb-2">
            <small class="fw-medium me-1">Цвета секторов:</small>
            ${activeData.map((row, i) => `<input type="color" class="form-control form-control-color pie-answer-color" data-id="${id}" data-index="${i}" value="${dataObj.pieColors[i] || PIE_COLORS[i % PIE_COLORS.length]}" style="width:26px;height:22px;padding:1px 2px;cursor:pointer;" title="${row.answer}">`).join('')}
            <button class="btn btn-sm btn-outline-secondary random-pie-colors-btn" data-id="${id}" title="Случайные цвета"><i class="fa-solid fa-dice-five"></i></button>
        </div>`;
    }

    dataObj.file_keys.forEach((fileKey, fi) => {
        const chartKey = `pie_${id}_${fi}`;
        if (window.charts[chartKey]) window.charts[chartKey].destroy();

        const canvas = document.getElementById(`pie_canvas_${id}_${fi}`);
        if (!canvas) return;
        const msgEl = document.getElementById(`pie_msg_${id}_${fi}`);

        const counts = activeData.map(r => r.counts[fileKey] || 0);
        const total = counts.reduce((a, b) => a + b, 0);
        const nonZero = counts.filter(v => v > 0).length;

        if (activeData.length === 0 || nonZero > MAX_PIE) {
            canvas.style.display = 'none';
            if (msgEl) { msgEl.textContent = activeData.length === 0 ? 'Нет активных вариантов' : `Слишком много вариантов (${nonZero}) для круговой диаграммы`; msgEl.classList.remove('d-none'); }
            return;
        }
        canvas.style.display = '';
        if (msgEl) msgEl.classList.add('d-none');
        canvas.width = 380;
        canvas.height = 380;

        const colors = activeData.map((_, i) => dataObj.pieColors[i] || PIE_COLORS[i % PIE_COLORS.length]);
        const ctx = canvas.getContext('2d');
        window.charts[chartKey] = new Chart(ctx, {
            type: 'pie',
            data: { labels, datasets: [{ data: counts, backgroundColor: colors, borderWidth: 1 }] },
            options: {
                responsive: false,
                plugins: {
                    legend: { display: dataObj.options.showLegend !== false, position: 'bottom', labels: { font: { family: '"Times New Roman", Times, serif', size: 11 } } },
                    tooltip: { enabled: true, displayColors: false, callbacks: { title: (items) => activeData[items[0].dataIndex]?.answer || '', label: (ctx) => ctx.raw } },
                    datalabels: {
                        color: '#fff',
                        font: { family: '"Times New Roman", Times, serif', size: 13, weight: 'bold' },
                        formatter: (value) => {
                            if (total === 0 || value === 0) return '';
                            const pct = (value / total) * 100;
                            return pct < 1 ? '<1%' : Math.round(pct) + '%';
                        }
                    }
                }
            }
        });
    });
}

function drawChart(id) {
    const dataObj = window.appData[id];
    if (!dataObj || !window.chartsData[id]) return;

    if (window.charts[id]) window.charts[id].destroy();

    const activeData = dataObj.data.filter(r => r.included).sort((a, b) => b._total - a._total);
    if (activeData.length === 0) return;

    const labels = activeData.map(r => r.answer.length > 50 ? r.answer.substring(0, 50) + '...' : r.answer);
    const isHorizontal = dataObj.options.chartDirection === 'y';
    const topN = dataObj.options.highlightTop ? Math.min(dataObj.options.topN, activeData.length * dataObj.file_keys.length) : 0;
    const HIGHLIGHT_COLOR = dataObj.options.highlightColor || '#0d6efd';

    const fileTotals = {};
    dataObj.file_keys.forEach(fk => {
        fileTotals[fk] = dataObj.data.reduce((sum, r) => sum + (r.counts[fk] || 0), 0);
    });

    const allBars = [];
    activeData.forEach(r => {
        dataObj.file_keys.forEach(fk => {
            const c = r.counts[fk] || 0;
            const ft = fileTotals[fk];
            allBars.push({ answer: r.answer, fileKey: fk, pct: ft > 0 ? c / ft * 100 : 0 });
        });
    });
    allBars.sort((a, b) => b.pct - a.pct);
    const topBarSet = new Set(allBars.slice(0, topN).map(b => `${b.answer}__${b.fileKey}`));

    const isSingleFile = dataObj.file_keys.length === 1;

    if (isSingleFile) {
        while (dataObj.barColors.length < activeData.length) {
            dataObj.barColors.push(PIE_COLORS[dataObj.barColors.length % PIE_COLORS.length]);
        }
        const editor = document.getElementById(`bar_color_editor_${id}`);
        if (editor) {
            editor.innerHTML = `<div class="d-flex flex-wrap gap-1 align-items-center ui-system-font mb-2">
                <small class="fw-medium me-1">Цвета столбиков:</small>
                ${activeData.map((row, i) => `<input type="color" class="form-control form-control-color bar-answer-color" data-id="${id}" data-index="${i}" value="${dataObj.barColors[i] || PIE_COLORS[i % PIE_COLORS.length]}" style="width:26px;height:22px;padding:1px 2px;cursor:pointer;" title="${row.answer.replace(/"/g, '&quot;')}">`).join('')}
                <button class="btn btn-sm btn-outline-secondary random-bar-colors-btn" data-id="${id}" title="Случайные цвета"><i class="fa-solid fa-dice-five"></i></button>
            </div>`;
        }
    } else {
        const editor = document.getElementById(`bar_color_editor_${id}`);
        if (editor) editor.innerHTML = '';
    }

    const datasets = dataObj.file_keys.map(fileKey => {
        const ft = fileTotals[fileKey];
        const rawCounts = activeData.map(r => r.counts[fileKey] || 0);
        return {
            label: dataObj.file_labels[fileKey],
            backgroundColor: activeData.map((r, barIdx) => {
                if (topN > 0 && topBarSet.has(`${r.answer}__${fileKey}`)) return HIGHLIGHT_COLOR;
                return isSingleFile ? (dataObj.barColors[barIdx] || PIE_COLORS[barIdx % PIE_COLORS.length]) : dataObj.file_colors[fileKey];
            }),
            data: rawCounts.map(c => ft > 0 ? (c / ft) * 100 : 0),
            rawCounts,
            barPercentage: 0.8
        };
    });

    const ctx = document.getElementById(`canvas_${id}`).getContext('2d');
    ctx.canvas.height = isHorizontal ? Math.max(200, labels.length * 36 + (datasets.length > 1 ? 40 : 0)) : 400;

    window.charts[id] = new Chart(ctx, {
        type: 'bar',
        data: { labels, datasets },
        options: {
            indexAxis: dataObj.options.chartDirection,
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: !!dataObj.options.showLegend,
                    position: 'bottom',
                    labels: {
                        font: { family: '"Times New Roman", Times, serif', size: 12 },
                        generateLabels: (chart) => dataObj.file_keys.map((fk, i) => {
                            const ds = chart.data.datasets[i];
                            return { text: ds.label, fillStyle: dataObj.file_colors[fk], strokeStyle: dataObj.file_colors[fk], lineWidth: 1, hidden: !chart.isDatasetVisible(i), datasetIndex: i };
                        })
                    }
                },
                tooltip: { enabled: true, displayColors: datasets.length > 1, callbacks: { title: (items) => activeData[items[0].dataIndex]?.answer || '', label: (ctx) => (datasets.length > 1 ? ctx.dataset.label + ': ' : '') + ctx.dataset.rawCounts[ctx.dataIndex] } },
                datalabels: {
                    color: '#000', anchor: 'end', align: isHorizontal ? 'right' : 'top', offset: 4,
                    font: { family: '"Times New Roman", Times, serif', size: 14, weight: 'bold' },
                    formatter: (value) => { if (!value || value === 0) return ''; if (value > 0 && value < 1) return '<1%'; return Math.round(value) + '%'; }
                }
            },
            scales: {
                x: { display: !isHorizontal, min: isHorizontal ? 0 : undefined, grid: { display: false }, border: { display: false } },
                y: { display: isHorizontal, min: !isHorizontal ? 0 : undefined, grid: { display: false }, border: { display: false } }
            },
            layout: { padding: isHorizontal ? { right: 50 } : { top: 30 } }
        }
    });
}

function drawStackedChart(id) {
    const dataObj = window.appData[id];
    if (!dataObj || !window.stackedChartsData[id]) return;

    if (window.charts['stacked_' + id]) window.charts['stacked_' + id].destroy();

    const activeData = dataObj.data.filter(r => r.included).sort((a, b) => b._total - a._total);
    if (activeData.length === 0) return;

    const fileTotals = {};
    dataObj.file_keys.forEach(fk => {
        fileTotals[fk] = activeData.reduce((sum, r) => sum + (r.counts[fk] || 0), 0);
    });

    const labels = activeData.map(r => r.answer.length > 50 ? r.answer.substring(0, 50) + '...' : r.answer);

    const datasets = dataObj.file_keys.map(fileKey => {
        const actualPcts = activeData.map(r => { const count = r.counts[fileKey] || 0; const ft = fileTotals[fileKey]; return ft > 0 ? (count / ft) * 100 : 0; });
        const actualCounts = activeData.map(r => r.counts[fileKey] || 0);
        return {
            label: dataObj.file_labels[fileKey],
            backgroundColor: dataObj.file_colors[fileKey],
            actualPcts, actualCounts,
            data: activeData.map((r, rIdx) => {
                const pct = actualPcts[rIdx];
                const answerSum = dataObj.file_keys.reduce((sum, fk) => { const c = r.counts[fk] || 0; const ft = fileTotals[fk]; return sum + (ft > 0 ? (c / ft) * 100 : 0); }, 0);
                return answerSum > 0 ? (pct / answerSum) * 100 : 0;
            }),
            barPercentage: 0.7
        };
    });

    const isVertStacked = dataObj.options.chartDirection === 'x';
    const ctx = document.getElementById(`stacked_canvas_${id}`).getContext('2d');
    ctx.canvas.height = isVertStacked ? 400 : Math.max(150, labels.length * 25);

    window.charts['stacked_' + id] = new Chart(ctx, {
        type: 'bar',
        data: { labels, datasets },
        options: {
            indexAxis: isVertStacked ? 'x' : 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: !!dataObj.options.showLegend, position: 'bottom', labels: { font: { family: '"Times New Roman", Times, serif', size: 12 } } },
                tooltip: { enabled: true, displayColors: datasets.length > 1, callbacks: { title: (items) => activeData[items[0].dataIndex]?.answer || '', label: (ctx) => (datasets.length > 1 ? ctx.dataset.label + ': ' : '') + ctx.dataset.actualCounts[ctx.dataIndex] } },
                datalabels: {
                    color: '#fff', anchor: 'center', align: 'center',
                    font: { family: '"Times New Roman", Times, serif', size: 13, weight: 'bold' },
                    formatter: (value, context) => { const pct = context.dataset.actualPcts[context.dataIndex]; if (!pct || pct === 0) return ''; const rounded = Math.round(pct); return rounded < 1 ? '<1%' : rounded + '%'; }
                }
            },
            scales: isVertStacked
                ? { x: { stacked: true, grid: { display: false }, border: { display: false } }, y: { stacked: true, display: false, max: 100, grid: { display: false }, border: { display: false } } }
                : { x: { stacked: true, display: false, max: 100, grid: { display: false }, border: { display: false } }, y: { stacked: true, grid: { display: false }, border: { display: false } } }
        }
    });
}

// ===================== TABLE RENDER =====================

function renderTable(tableId) {
    const dataObj = window.appData[tableId];
    const tableEl = document.getElementById(tableId);
    if (!tableEl) return;

    const isVert = dataObj.options.tableVertical;
    const activeRows = dataObj.data.filter(r => r.included);

    const totals = {};
    dataObj.file_keys.forEach(fk => {
        totals[fk] = dataObj.data.reduce((sum, r) => sum + (r.counts[fk] || 0), 0);
    });

    const topN = dataObj.options.highlightTop ? Math.min(dataObj.options.topN, activeRows.length) : 0;
    const highlightColor = dataObj.options.highlightColor || '#dc3545';
    let threshold = -1;

    if(dataObj.options.highlightTop && activeRows.length > 0) {
        const uniqueCounts = [...new Set(activeRows.map(r => r._total))].sort((a, b) => b - a);
        threshold = uniqueCounts[Math.min(topN, uniqueCounts.length) - 1];
    }

    if (isVert) {
        const isSingleFileVert = dataObj.file_keys.length === 1;
        let theadHtml1 = isSingleFileVert
            ? `<tr id="thead_row1_${tableId}">`
            : `<tr id="thead_row1_${tableId}"><th class="text-start align-middle main-th-${tableId}" rowspan="2" style="width: 15%;"><span contenteditable="true" class="editable-cell text-dark" data-id="${tableId}" data-header="h1">Файл</span></th>`;
        let theadHtml2 = `<tr id="thead_row2_${tableId}">`;

        dataObj.data.forEach((row, origIdx) => {
            const isTop = row.included && topN > 0 && row._total >= threshold;
            const excluded = !row.included;
            const textClass = isTop ? 'text-danger' : (excluded ? 'text-muted' : '');
            const thStyle = excluded ? 'text-decoration:line-through;text-decoration-thickness:2px;text-decoration-color:#868e96;color:#6c757d;background-color:#e9ecef;' : '';

            theadHtml1 += `
                <th class="text-center file-header-${tableId} ${textClass} ${excluded ? 'col-excluded' : ''}" colspan="2" style="${thStyle}">
                    <div class="d-flex align-items-center justify-content-center gap-1">
                        <div class="row-toggle-container">
                            <div class="form-check form-switch mb-0">
                                <input class="form-check-input row-toggle ui-system-font" type="checkbox" ${row.included ? 'checked' : ''} data-id="${tableId}" data-index="${origIdx}" title="${excluded ? 'Включить' : 'Исключить'} ответ">
                            </div>
                        </div>
                        <span contenteditable="true" class="editable-cell answer-text" data-id="${tableId}" data-index="${origIdx}">${row.answer}</span>
                    </div>
                </th>`;
            theadHtml2 += `
                <th class="text-center count-col-${tableId}"><span contenteditable="true" class="editable-cell text-dark" data-id="${tableId}" data-header="h2">${dataObj.headers.h2 || 'Кол-во'}</span></th>
                <th class="text-center pct-col-${tableId}"><span contenteditable="true" class="editable-cell text-dark" data-id="${tableId}" data-header="h3">${dataObj.headers.h3 || '%'}</span></th>`;
        });

        if (dataObj.options.showTotal) {
            theadHtml1 += `<th class="text-center file-header-${tableId}" colspan="2">Всего</th>`;
            theadHtml2 += `<th class="text-center count-col-${tableId}">Кол-во</th><th class="text-center pct-col-${tableId}">%</th>`;
        }
        theadHtml1 += `</tr>`;
        theadHtml2 += `</tr>`;

        tableEl.querySelector('thead').innerHTML = theadHtml1 + theadHtml2;

        let tbodyHtml = '';
        dataObj.file_keys.forEach(fk => {
            const tot = totals[fk];
            let rowHtml = isSingleFileVert ? `<tr>` : `<tr><td class="text-start fw-medium">${dataObj.file_labels[fk]}</td>`;

            dataObj.data.forEach((row, origIdx) => {
                const c = row.counts[fk] || 0;
                const excluded = !row.included;
                let pct = '0';
                if (!excluded && tot > 0 && c > 0) {
                    const rawPct = (c / tot) * 100;
                    pct = (rawPct > 0 && rawPct < 1) ? '<1' : Math.round(rawPct).toString();
                }
                const isTop = row.included && topN > 0 && row._total >= threshold;
                const cellStyle = excluded ? 'text-decoration:line-through;text-decoration-thickness:2px;text-decoration-color:#868e96;color:#6c757d;background-color:#e9ecef;' : '';
                const textClass = isTop ? 'text-danger' : '';

                rowHtml += `
                    <td class="text-center align-middle count-col-${tableId}" style="${cellStyle}"><span contenteditable="true" class="editable-cell answer-count ${textClass}" data-id="${tableId}" data-index="${origIdx}" data-file="${fk}">${c}</span></td>
                    <td class="text-center align-middle pct-col-${tableId} ${textClass}" style="${cellStyle}">${pct}</td>`;
            });

            if (dataObj.options.showTotal) {
                rowHtml += `<td class="text-center fw-bold count-col-${tableId}">${tot}</td><td class="text-center fw-bold pct-col-${tableId}">${tot > 0 ? '100' : '0'}</td>`;
            }
            rowHtml += `</tr>`;
            tbodyHtml += rowHtml;
        });

        tableEl.querySelector('tbody').innerHTML = tbodyHtml;
        tableEl.querySelector('tfoot').innerHTML = '';
        tableEl.querySelector('tfoot').classList.add('d-none');

    } else {
        const isSingleFile = dataObj.file_keys.length === 1;
        let theadHtml;
        if (isSingleFile) {
            theadHtml = `
                <tr id="thead_row1_${tableId}">
                    <th class="control-col main-th-${tableId}"></th>
                    <th class="text-start align-middle main-th-${tableId}" style="width: 40%;"><span contenteditable="true" class="editable-cell text-dark" data-id="${tableId}" data-header="h1">${dataObj.headers.h1 || 'Ответ'}</span></th>
                    <th class="text-center count-col-${tableId}"><span contenteditable="true" class="editable-cell text-dark" data-id="${tableId}" data-header="h2">${dataObj.headers.h2 || 'Кол-во ответивших'}</span></th>
                    <th class="text-center pct-col-${tableId}"><span contenteditable="true" class="editable-cell text-dark" data-id="${tableId}" data-header="h3">${dataObj.headers.h3 || '% от числа ответивших'}</span></th>
                </tr>`;
        } else {
            theadHtml = `
                <tr id="thead_row1_${tableId}">
                    <th class="control-col main-th-${tableId}" rowspan="2"></th>
                    <th class="text-start align-middle main-th-${tableId}" rowspan="2" style="width: 40%;"><span contenteditable="true" class="editable-cell text-dark" data-id="${tableId}" data-header="h1">${dataObj.headers.h1 || 'Ответ'}</span></th>
                    ${dataObj.file_keys.map(fk => `<th class="text-center file-header-${tableId}" colspan="2">${dataObj.file_labels[fk]}</th>`).join('')}
                </tr>
                <tr id="thead_row2_${tableId}">
                    ${dataObj.file_keys.map(() => `
                        <th class="text-center count-col-${tableId}"><span contenteditable="true" class="editable-cell text-dark" data-id="${tableId}" data-header="h2">${dataObj.headers.h2 || 'Кол-во ответивших'}</span></th>
                        <th class="text-center pct-col-${tableId}"><span contenteditable="true" class="editable-cell text-dark" data-id="${tableId}" data-header="h3">${dataObj.headers.h3 || '% от числа ответивших'}</span></th>
                    `).join('')}
                </tr>`;
        }
        tableEl.querySelector('thead').innerHTML = theadHtml;

        let tbodyHtml = '';
        dataObj.data.forEach((row, idx) => {
            const isTop = topN > 0 && row.included && row._total >= threshold;
            const trClass = row.included ? (isTop ? 'row-top-highlight' : '') : 'row-excluded';
            const trStyle = isTop ? ` style="--top-color:${highlightColor};"` : '';
            let tdHtml = '';

            dataObj.file_keys.forEach(fk => {
                const c = row.counts[fk] || 0;
                const tot = totals[fk];
                let pct = '0';
                if (row.included && tot > 0 && c > 0) {
                    const rawPct = (c / tot) * 100;
                    pct = (rawPct > 0 && rawPct < 1) ? '<1' : Math.round(rawPct).toString();
                }
                tdHtml += `
                    <td class="text-center align-middle count-col-${tableId}"><span contenteditable="true" class="editable-cell answer-count" data-id="${tableId}" data-index="${idx}" data-file="${fk}">${c}</span></td>
                    <td class="text-center align-middle pct-col-${tableId}">${pct}</td>`;
            });

            tbodyHtml += `
                <tr class="${trClass}"${trStyle}>
                    <td class="control-col">
                        <div class="row-toggle-container">
                            <div class="form-check form-switch mb-0">
                                <input class="form-check-input row-toggle ui-system-font" type="checkbox" ${row.included ? 'checked' : ''} data-id="${tableId}" data-index="${idx}" title="Включить/исключить ответ">
                            </div>
                        </div>
                    </td>
                    <td class="text-start align-middle">
                        <div class="d-flex align-items-center gap-1">
                            <span contenteditable="true" class="editable-cell answer-text flex-grow-1" data-id="${tableId}" data-index="${idx}">${row.answer}</span>
                            <button type="button" class="btn btn-sm btn-link text-muted p-0 flex-shrink-0 merge-answer-open-btn"
                                    onclick="openMergeAnswerModal('${tableId}', ${idx})"
                                    title="Присоединить к другому ответу"
                                    style="font-size:0.75rem;line-height:1;opacity:0.5;" onmouseenter="this.style.opacity='1'" onmouseleave="this.style.opacity='0.5'">
                                <i class="fa-solid fa-plus-circle"></i>
                            </button>
                        </div>
                    </td>
                    ${tdHtml}
                </tr>`;
        });
        tableEl.querySelector('tbody').innerHTML = tbodyHtml;

        if (dataObj.options.showTotal) {
            let footTd = '';
            dataObj.file_keys.forEach(fk => {
                footTd += `<td class="text-center count-col-${tableId}">${totals[fk]}</td><td class="text-center pct-col-${tableId}">${totals[fk] > 0 ? '100' : '0'}</td>`;
            });
            tableEl.querySelector('tfoot').innerHTML = `<tr class="fw-bold bg-light"><td class="control-col"></td><td class="text-start">Всего</td>${footTd}</tr>`;
            tableEl.querySelector('tfoot').classList.remove('d-none');
        } else { tableEl.querySelector('tfoot').classList.add('d-none'); }
    }

    applyHiddenColumns(tableId);
}

// ===================== SETTINGS EVENTS =====================

document.addEventListener('change', (e) => {
    if (e.target.classList.contains('setting-show-total')) { window.appData[e.target.dataset.id].options.showTotal = e.target.checked; renderTable(e.target.dataset.id); }
    if (e.target.classList.contains('setting-highlight-top')) { window.appData[e.target.dataset.id].options.highlightTop = e.target.checked; drawChart(e.target.dataset.id); renderTable(e.target.dataset.id); }
    if (e.target.classList.contains('setting-vertical')) {
        const id = e.target.dataset.id;
        window.appData[id].options.tableVertical = e.target.checked;
        window.appData[id].options.chartDirection = e.target.checked ? 'x' : 'y';
        renderTable(id); drawChart(id); drawStackedChart(id);
    }
    if (e.target.classList.contains('setting-highlight-color')) { window.appData[e.target.dataset.id].options.highlightColor = e.target.value; drawChart(e.target.dataset.id); renderTable(e.target.dataset.id); }
    if (e.target.classList.contains('pie-answer-color')) { const id = e.target.dataset.id; window.appData[id].pieColors[parseInt(e.target.dataset.index)] = e.target.value; drawPieChart(id); }
    if (e.target.classList.contains('bar-answer-color')) { const id = e.target.dataset.id; window.appData[id].barColors[parseInt(e.target.dataset.index)] = e.target.value; drawChart(id); }
    if (e.target.classList.contains('row-toggle')) { window.appData[e.target.dataset.id].data[parseInt(e.target.dataset.index)].included = e.target.checked; renderTable(e.target.dataset.id); drawChart(e.target.dataset.id); drawStackedChart(e.target.dataset.id); drawPieChart(e.target.dataset.id); }
    if (e.target.classList.contains('modal-row-toggle')) { const id = e.target.dataset.id; window.appData[id].data[parseInt(e.target.dataset.index)].included = e.target.checked; renderChartEditModal(id); renderTable(id); drawChart(id); drawStackedChart(id); drawPieChart(id); }
    if (e.target.classList.contains('setting-show-legend')) { const id = e.target.dataset.id; window.appData[id].options.showLegend = e.target.checked; drawChart(id); drawStackedChart(id); drawPieChart(id); }
});

document.addEventListener('input', (e) => {
    if (e.target.classList.contains('setting-top-n')) { const val = parseInt(e.target.value); if (!isNaN(val) && val >= 1) { window.appData[e.target.dataset.id].options.topN = val; if (window.appData[e.target.dataset.id].options.highlightTop) { drawChart(e.target.dataset.id); renderTable(e.target.dataset.id); } } }
    if (e.target.classList.contains('answer-text')) { window.appData[e.target.dataset.id].data[parseInt(e.target.dataset.index)].answer = e.target.innerText; }
    if (e.target.classList.contains('answer-count')) { const val = parseInt(e.target.innerText); if (!isNaN(val)) { const fk = e.target.getAttribute('data-file'); const rowData = window.appData[e.target.dataset.id].data[parseInt(e.target.dataset.index)]; rowData.counts[fk] = val; rowData._total = Object.values(rowData.counts).reduce((a, b) => a + b, 0); } }
    if (e.target.hasAttribute('data-header')) { window.appData[e.target.dataset.id].headers[e.target.getAttribute('data-header')] = e.target.innerText; }
    if (e.target.classList.contains('modal-answer-text')) { window.appData[e.target.dataset.id].data[parseInt(e.target.dataset.index)].answer = e.target.innerText; }
    if (e.target.classList.contains('modal-answer-count')) { const val = parseInt(e.target.innerText); if (!isNaN(val)) { const fk = e.target.getAttribute('data-file'); const rowData = window.appData[e.target.dataset.id].data[parseInt(e.target.dataset.index)]; rowData.counts[fk] = val; rowData._total = Object.values(rowData.counts).reduce((a, b) => a + b, 0); } }
});

document.addEventListener('focusout', (e) => {
    if (e.target.classList.contains('answer-count') || e.target.classList.contains('answer-text') || e.target.hasAttribute('data-header')) { renderTable(e.target.dataset.id); drawChart(e.target.dataset.id); drawStackedChart(e.target.dataset.id); drawPieChart(e.target.dataset.id); }
    if (e.target.classList.contains('modal-answer-count') || e.target.classList.contains('modal-answer-text')) { const id = e.target.dataset.id; renderChartEditModal(id); renderTable(id); drawChart(id); drawStackedChart(id); drawPieChart(id); }
});

document.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && e.target.classList.contains('editable-cell')) { e.preventDefault(); e.target.blur(); }
});

// ===================== ANALYZE FORM =====================

document.getElementById('analyzeForm').addEventListener('submit', async (e) => {
    e.preventDefault();

    const fileLabels = {};
    const fileColors = {};
    document.querySelectorAll('.legend-label').forEach(el => { fileLabels[el.dataset.file] = el.value || el.placeholder; });
    document.querySelectorAll('.legend-color').forEach(el => { fileColors[el.dataset.file] = el.value; });

    const items = document.querySelectorAll('#sortableQuestionsList .question-item');
    const configs = [];
    const ALL_VIZ = ['Таблица', 'Столбчатая диаграмма', 'Накопленная диаграмма', 'Круговая диаграмма'];
    items.forEach(item => {
        const colName = item.getAttribute('data-col');
        const subColName = window.questionMerges[colName];
        let mergedSub = null;
        if (subColName) {
            mergedSub = {};
            window.processedFiles.forEach(f => {
                const subExists = f.columns.find(c => c.name === subColName);
                if (subExists) mergedSub[f.clean_filename] = subColName;
            });
        }

        // Найти раздел для этого вопроса
        const sec = window.questionSections.find(s => s.questions.includes(colName));
        const sectionInfo = sec
            ? { name: sec.name, description: sec.description || '', color: sec.color }
            : null;

        configs.push({
            column: colName,
            viz_type: ALL_VIZ,
            file_mapping: window.questionMapping[colName],
            merged_sub: mergedSub,
            section: sectionInfo
        });
    });
    const payload = { file_labels: fileLabels, file_colors: fileColors, configs };

    document.getElementById('analyzeBtn').disabled = true;
    document.getElementById('analyzeSpinner').classList.remove('d-none');

    try {
        const response = await fetch('/analyze', {
            method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload)
        });
        const data = await response.json();

        if (response.ok) {
            const reportContent = document.getElementById('reportContent');
            reportContent.innerHTML = '';
            window.appData = {};
            window.chartsData = {};
            window.stackedChartsData = {};
            window.pieChartsData = {};
            window.renderedTabs = {};

            let tableCounter = 1;
            let figureCounter = 1;

            data.results.forEach((item, i) => {
                const id = `item_${i}`;
                window.renderedTabs[id] = { bar: false, stacked: false, pie: false };
                window.chartsData[id] = true;
                window.stackedChartsData[id] = true;
                window.pieChartsData[id] = true;

                window.appData[id] = {
                    question_name: item.col_name,
                    options: { showTotal: true, highlightTop: false, topN: 1, chartDirection: 'y', highlightColor: '#dc3545', hiddenCol: 'none', tableVertical: false, showLegend: item.file_keys.length > 1 },
                    headers: { h1: "Ответ", h2: "Кол-во ответивших", h3: "% от числа ответивших" },
                    data: item.data,
                    file_keys: item.file_keys,
                    file_labels: item.file_labels,
                    file_colors: item.file_colors,
                    pieColors: [...PIE_COLORS],
                    barColors: [...PIE_COLORS]
                };

                const tNum = tableCounter++;
                const fNumBar = figureCounter++;
                const fNumStacked = figureCounter++;
                const fNumPie = figureCounter;
                figureCounter += item.file_keys.length;

                const pieCanvasesHtml = item.file_keys.map((fk, fi) => `
                    <div class="text-center flex-fill">
                        <div class="fw-medium mb-1 ui-system-font small">${item.file_labels[fk]}</div>
                        <div class="pie-container">
                            <canvas id="pie_canvas_${id}_${fi}"></canvas>
                            <p class="text-muted small py-3 d-none" id="pie_msg_${id}_${fi}"></p>
                        </div>
                    </div>`).join('');

                // Цветовая полоска раздела в заголовке вопроса в отчёте
                const secInfo = item.section;
                const sectionBorder = secInfo ? `border-left: 4px solid ${secInfo.color}; padding-left: 8px;` : '';

                const html = `<h5 class="fw-semibold text-dark mb-2" style="${sectionBorder}">${item.col_name}</h5><div class="result-item">
                    <ul class="nav nav-tabs ui-system-font" id="tabs_${id}">
                        <li class="nav-item"><button class="nav-link active viz-tab-btn" data-id="${id}" data-tab="table"><i class="fa-solid fa-table me-1"></i>Таблица</button></li>
                        <li class="nav-item"><button class="nav-link viz-tab-btn" data-id="${id}" data-tab="bar"><i class="fa-solid fa-chart-column me-1"></i>Столбчатая</button></li>
                        <li class="nav-item"><button class="nav-link viz-tab-btn" data-id="${id}" data-tab="stacked"><i class="fa-solid fa-chart-bar me-1"></i>Накопленная</button></li>
                        <li class="nav-item"><button class="nav-link viz-tab-btn" data-id="${id}" data-tab="pie"><i class="fa-solid fa-chart-pie me-1"></i>Круговая</button></li>
                    </ul>
                    <div class="ui-system-font bg-white p-2 mb-3 border border-top-0 border-secondary-subtle rounded-bottom d-flex flex-wrap gap-3 align-items-center shadow-sm" id="settings_${id}">
                        <div class="form-check mb-0" data-vis-tabs="table"><input class="form-check-input setting-show-total" type="checkbox" id="total_${id}" data-id="${id}" checked><label class="form-check-label small fw-medium" for="total_${id}"><i class="fa-solid fa-sigma me-1 text-muted"></i>Добавить строку "Всего"</label></div>
                        <div class="form-check form-switch mb-0" data-vis-tabs="table bar"><input class="form-check-input setting-highlight-top" type="checkbox" id="hl_${id}" data-id="${id}"><label class="form-check-label small fw-medium" for="hl_${id}"><i class="fa-solid fa-trophy me-1 text-muted"></i>Выделить топ:</label></div>
                        <input type="number" class="form-control form-control-sm setting-top-n" data-id="${id}" value="1" min="1" max="${item.data.length * item.file_keys.length}" style="width: 70px;" data-vis-tabs="table bar">
                        <input type="color" class="form-control form-control-color setting-highlight-color" data-id="${id}" value="#dc3545" style="width:28px;height:28px;padding:1px 2px;cursor:pointer;" title="Цвет выделения топа" data-vis-tabs="table bar">
                        <button type="button" class="btn btn-sm btn-outline-secondary random-highlight-color-btn" data-id="${id}" title="Случайный цвет" data-vis-tabs="table bar"><i class="fa-solid fa-dice-five"></i></button>
                        <div class="vr" data-vis-tabs="table bar"></div>
                        <div class="form-check form-switch mb-0" data-vis-tabs="table bar stacked"><input class="form-check-input setting-vertical" type="checkbox" id="vert_${id}" data-id="${id}"><label class="form-check-label small fw-medium" for="vert_${id}"><i class="fa-solid fa-rotate me-1 text-muted"></i>Вертикальный</label></div>
                        <div class="form-check mb-0" data-vis-tabs="bar stacked pie"><input class="form-check-input setting-show-legend" type="checkbox" id="legend_${id}" data-id="${id}" ${item.file_keys.length > 1 ? 'checked' : ''}><label class="form-check-label small fw-medium" for="legend_${id}"><i class="fa-solid fa-list-ul me-1 text-muted"></i>Легенда</label></div>
                        <button type="button" class="btn btn-sm btn-outline-secondary ms-auto" onclick="openHideColModal('${id}')" title="Скрыть столбцы" data-vis-tabs="table"><i class="fa-solid fa-eye-slash"></i></button>
                        <button type="button" class="btn btn-sm btn-outline-secondary" onclick="openChartEditModal('${id}')" title="Редактировать данные диаграммы" data-vis-tabs="bar stacked pie"><i class="fa-solid fa-pen"></i></button>
                        <button type="button" class="btn btn-sm btn-outline-secondary" onclick="confirmFuzzyMapping('${id}')" title="Сгруппировать похожие ответы"><i class="fa-solid fa-shuffle"></i></button>
                    </div>
                    <div id="pane_table_${id}">
                        <div class="mb-3">Таблица ${tNum} – Распределение ответов респондентов на вопрос: «${item.col_name}»</div>
                        <div class="table-responsive mb-4">
                            <table class="table table-bordered table-hover table-custom-border align-middle mb-0" id="${id}">
                                <thead></thead><tbody></tbody><tfoot></tfoot>
                            </table>
                        </div>
                    </div>
                    <div id="pane_bar_${id}" class="d-none">
                        <div id="bar_color_editor_${id}"></div>
                        <div class="chart-container mb-2"><canvas id="canvas_${id}"></canvas></div>
                        <div class="text-center mb-4">Рисунок ${fNumBar} – Распределение ответов респондентов на вопрос: «${item.col_name}»</div>
                    </div>
                    <div id="pane_stacked_${id}" class="d-none">
                        <div class="stacked-container mb-2"><canvas id="stacked_canvas_${id}"></canvas></div>
                        <div class="text-center mb-4">Рисунок ${fNumStacked} – Распределение ответов респондентов на вопрос: «${item.col_name}»</div>
                    </div>
                    <div id="pane_pie_${id}" class="d-none">
                        <div id="pie_color_editor_${id}"></div>
                        <div class="d-flex gap-3 flex-wrap">${pieCanvasesHtml}</div>
                        <div class="text-center mb-4">Рисунок ${fNumPie} – Распределение ответов респондентов на вопрос: «${item.col_name}»</div>
                    </div>
                </div>`;

                reportContent.innerHTML += html;
                document.getElementById(`settings_${id}`)?.querySelectorAll('[data-vis-tabs]').forEach(el => {
                    el.classList.toggle('d-none', !el.dataset.visTabs.split(' ').includes('table'));
                });
                renderTable(id);
            });
            goToStep(5);
        } else { showToast(data.message, 'danger'); }
    } catch (err) { showToast('Ошибка соединения с сервером', 'danger'); }
    finally {
        updateQuestionsBtn();
        document.getElementById('analyzeSpinner').classList.add('d-none');
    }
});

// ===================== CLICK DELEGATION =====================

function randomColor() {
    return '#' + Math.floor(Math.random() * 0xffffff).toString(16).padStart(6, '0');
}

document.addEventListener('click', e => {
    if (e.target.closest('.wizard-back-btn')) {
        if (currentWizardStep > 0) goToStep(currentWizardStep - 1);
        return;
    }
    const rhlBtn = e.target.closest('.random-highlight-color-btn');
    if (rhlBtn) { const input = document.querySelector(`.setting-highlight-color[data-id="${rhlBtn.dataset.id}"]`); if (input) { input.value = randomColor(); input.dispatchEvent(new Event('change', { bubbles: true })); } return; }
    const rpcBtn = e.target.closest('.random-pie-colors-btn');
    if (rpcBtn) { const did = rpcBtn.dataset.id; const dataObj = window.appData[did]; if (dataObj) { dataObj.pieColors = dataObj.pieColors.map(() => randomColor()); window.renderedTabs[did].pie = false; drawPieChart(did); } return; }
    const rbcBtn = e.target.closest('.random-bar-colors-btn');
    if (rbcBtn) { const did = rbcBtn.dataset.id; const dataObj = window.appData[did]; if (dataObj) { dataObj.barColors = dataObj.barColors.map(() => randomColor()); drawChart(did); } return; }
    const rLegBtn = e.target.closest('.random-legend-color-btn');
    if (rLegBtn) { const input = document.querySelector(`.legend-color[data-file="${rLegBtn.dataset.file}"]`); if (input) input.value = randomColor(); return; }

    const btn = e.target.closest('.viz-tab-btn');
    if (!btn) return;
    const id = btn.dataset.id;
    const tab = btn.dataset.tab;

    document.querySelectorAll(`#tabs_${id} .viz-tab-btn`).forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    ['table', 'bar', 'stacked', 'pie'].forEach(t => { const pane = document.getElementById(`pane_${t}_${id}`); if (pane) pane.classList.toggle('d-none', t !== tab); });

    const settingsEl = document.getElementById(`settings_${id}`);
    if (settingsEl) { settingsEl.querySelectorAll('[data-vis-tabs]').forEach(el => { const tabs = el.dataset.visTabs.split(' '); el.classList.toggle('d-none', !tabs.includes(tab)); }); }

    if (tab === 'bar' && !window.renderedTabs[id].bar) { window.renderedTabs[id].bar = true; setTimeout(() => drawChart(id), 50); }
    if (tab === 'stacked' && !window.renderedTabs[id].stacked) { window.renderedTabs[id].stacked = true; setTimeout(() => drawStackedChart(id), 50); }
    if (tab === 'pie' && !window.renderedTabs[id].pie) { window.renderedTabs[id].pie = true; setTimeout(() => drawPieChart(id), 50); }
});

// ===================== ЭКСПОРТ =====================

document.getElementById('downloadCleanedBtn').addEventListener('click', async () => {
    const ids = Object.keys(window.appData || {}).sort((a, b) => parseInt(a.split('_')[1]) - parseInt(b.split('_')[1]));

    if (ids.length === 0) { showToast('Нет данных для экспорта. Сначала постройте отчёт на шаге 4.', 'danger'); return; }

    const questions = [];
    let tableNum = 1;

    ids.forEach(id => {
        const dataObj = window.appData[id];
        if (!dataObj) return;

        const activeRows = dataObj.data.filter(r => r.included !== false);
        if (activeRows.length === 0) return;

        const fileTotals = {};
        dataObj.file_keys.forEach(fk => {
            fileTotals[fk] = dataObj.data.reduce((sum, r) => sum + (r.counts[fk] || 0), 0);
        });

        // Найти раздел из window.questionSections по имени вопроса
        const sec = window.questionSections.find(s => s.questions.includes(dataObj.question_name));

        questions.push({
            table_num: tableNum++,
            question_name: dataObj.question_name,
            h1: dataObj.headers.h1 || 'Ответ',
            h2: dataObj.headers.h2 || 'Кол-во ответивших',
            h3: dataObj.headers.h3 || '% от числа ответивших',
            file_keys: dataObj.file_keys,
            file_labels: dataObj.file_labels,
            rows: activeRows.map(r => ({ answer: String(r.answer), counts: Object.fromEntries(dataObj.file_keys.map(fk => [fk, r.counts[fk] || 0])) })),
            file_totals: fileTotals,
            show_total: dataObj.options.showTotal !== false,
            section: sec ? { name: sec.name, description: sec.description || '', color: sec.color } : null
        });
    });

    if (questions.length === 0) { showToast('Все строки исключены — нечего экспортировать.', 'danger'); return; }

    const btn = document.getElementById('downloadCleanedBtn');
    const origHtml = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Генерация...';

    // Показать прогресс-бар
    const progressContainer = document.getElementById('exportProgressContainer');
    const progressBar = document.getElementById('exportProgressBar');
    const progressLabel = document.getElementById('exportProgressLabel');
    if (progressContainer) progressContainer.classList.remove('d-none');
    if (progressBar) { progressBar.style.width = '0%'; progressBar.textContent = ''; }
    if (progressLabel) progressLabel.textContent = 'Подготовка...';

    try {
        const response = await fetch('/export_docx_stream', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ questions })
        });

        if (!response.ok) {
            const data = await response.json();
            showToast(data.message || 'Ошибка генерации документа', 'danger');
            return;
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            buffer += decoder.decode(value, { stream: true });

            const lines = buffer.split('\n\n');
            buffer = lines.pop(); // неполная строка

            for (const line of lines) {
                if (!line.startsWith('data: ')) continue;
                const msg = JSON.parse(line.slice(6));

                if (msg.type === 'progress') {
                    const pct = Math.round((msg.current / msg.total) * 100);
                    if (progressBar) { progressBar.style.width = pct + '%'; progressBar.textContent = `${msg.current}/${msg.total}`; }
                    if (progressLabel) progressLabel.textContent = `Раздел ${msg.current} из ${msg.total}: ${msg.label}`;
                }

                if (msg.type === 'done') {
                    // Скачать файл данных
                    const dataBytes = Uint8Array.from(atob(msg.data), c => c.charCodeAt(0));
                    const dataBlob = new Blob([dataBytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
                    const dataUrl = URL.createObjectURL(dataBlob);
                    const a1 = document.createElement('a');
                    a1.href = dataUrl; a1.download = 'report_data.docx';
                    document.body.appendChild(a1); a1.click();
                    document.body.removeChild(a1); URL.revokeObjectURL(dataUrl);

                    // Небольшая задержка чтобы браузер не блокировал второе скачивание
                    await new Promise(r => setTimeout(r, 400));

                    // Скачать файл аналитики
                    const analysisBytes = Uint8Array.from(atob(msg.analysis), c => c.charCodeAt(0));
                    const analysisBlob = new Blob([analysisBytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
                    const analysisUrl = URL.createObjectURL(analysisBlob);
                    const a2 = document.createElement('a');
                    a2.href = analysisUrl; a2.download = 'report_analysis.docx';
                    document.body.appendChild(a2); a2.click();
                    document.body.removeChild(a2); URL.revokeObjectURL(analysisUrl);

                    showToast(`Готово: скачано 2 файла (данные + аналитика)`, 'success');
                }

                if (msg.type === 'error') {
                    showToast(msg.message || 'Ошибка генерации', 'danger');
                }
            }
        }
    } catch (err) {
        showToast('Ошибка соединения с сервером', 'danger');
    } finally {
        btn.disabled = false;
        btn.innerHTML = origHtml;
        if (progressContainer) setTimeout(() => progressContainer.classList.add('d-none'), 2000);
    }
});

initTooltips();