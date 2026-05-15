// ===================== ШАГ 4: СТРУКТУРА ОТЧЁТА =====================
window.reportSections = window.reportSections || [];
let _sectionIdCounter = 0;

function _getAllStep3QNames() {
    return Array.from(document.querySelectorAll('#sortableQuestionsList .question-item')).map(el => el.dataset.col);
}

function _getAssignedQNames() {
    const s = new Set();
    (window.reportSections || []).forEach(sec => sec.questions.forEach(q => s.add(q.qName)));
    return s;
}

function _updateToStep5Btn() {
    const btn = document.getElementById('toStep5Btn');
    if (!btn) return;
    const hasAnything = (window.reportSections && window.reportSections.length > 0)
        || document.querySelectorAll('#availableQuestionsList .available-q-item').length > 0
        || document.querySelectorAll('#sortableQuestionsList .question-item').length > 0;
    btn.disabled = !hasAnything;
    btn.title = hasAnything ? '' : 'Выберите вопросы на шаге 3';
}

function renderStep4() {
    _renderSectionsList();
    _renderAvailableQuestions();
}

function _renderSectionsList() {
    const container = document.getElementById('sectionsList');
    _updateToStep5Btn();
    if (!window.reportSections || !window.reportSections.length) {
        container.innerHTML = '<p class="text-muted small text-center mt-3 px-2 mb-0"></p>';
        return;
    }
    container.innerHTML = '';
    window.reportSections.forEach(sec => {
        container.insertAdjacentHTML('beforeend', _buildSectionCardHtml(sec));
    });
    window.reportSections.forEach(sec => {
        const listEl = container.querySelector(`.section-questions-list[data-section-id="${sec.id}"]`);
        if (!listEl) return;
        new Sortable(listEl, {
            handle: '.section-q-drag',
            animation: 150,
            ghostClass: 'sortable-ghost',
            filter: '.section-empty-hint',
            onEnd: evt => {
                const s = window.reportSections.find(x => x.id === sec.id);
                if (!s) return;
                const moved = s.questions.splice(evt.oldIndex, 1)[0];
                s.questions.splice(evt.newIndex, 0, moved);
            }
        });
    });
}

function _buildSectionCardHtml(sec) {
    const qHtml = sec.questions.length === 0
        ? `<div class="section-empty-hint">Перетащите вопросы сюда или нажмите «+»</div>`
        : sec.questions.map(q => _buildSectionQuestionHtml(sec.id, q)).join('');
    const collapsed = window._collapsedSections && window._collapsedSections.has(sec.id);
    const sectionColor = sec.color || '#a0bce5';
        return `
        <div class="card mb-2 section-card" data-section-id="${_escAttr(sec.id)}" style="border-left: 3px solid ${sectionColor};">
            <div class="card-body py-2 px-3">
                <div class="d-flex align-items-center gap-2 mb-1">
                    <button type="button" class="section-collapse-btn" data-section-id="${_escAttr(sec.id)}" title="${collapsed ? 'Развернуть' : 'Свернуть'}" style="color:${sectionColor};">
                        <i class="fa-solid ${collapsed ? 'fa-chevron-up' : 'fa-chevron-down'}"></i>
                    </button>
                    <span class="fw-semibold flex-grow-1 text-truncate" title="${_escAttr(sec.name)}" style="color:${sectionColor};">${_escHtml(sec.name)}</span>
                <button type="button" class="btn btn-sm btn-outline-primary py-0 px-2 add-qs-to-section-btn" data-section-id="${_escAttr(sec.id)}" title="Добавить выбранные вопросы">
                    <i class="fa-solid fa-plus"></i>
                </button>
                <button type="button" class="btn btn-sm btn-outline-secondary py-0 px-2 edit-section-btn" data-section-id="${_escAttr(sec.id)}" title="Редактировать раздел">
                    <i class="fa-solid fa-pen"></i>
                </button>
                <button type="button" class="btn btn-sm btn-outline-danger py-0 px-2 delete-section-btn" data-section-id="${_escAttr(sec.id)}" title="Удалить раздел">
                    <i class="fa-solid fa-trash-can"></i>
                </button>
            </div>
            <div class="section-questions-wrapper${collapsed ? ' collapsed' : ''}" data-section-id="${_escAttr(sec.id)}">
                <div class="section-questions-list border rounded p-1" data-section-id="${_escAttr(sec.id)}">
                    ${qHtml}
                </div>
            </div>
        </div>
    </div>`;
}

function _isMissingMapping(qName) {
    if (!window.processedFiles || window.processedFiles.length <= 1) return false;
    const mapping = window.questionMapping[qName] || {};
    return Object.keys(mapping).length < window.processedFiles.length;
}

function _buildSectionQuestionHtml(sectionId, q) {
    const multiFile = window.processedFiles && window.processedFiles.length > 1;
    const missing = multiFile && _isMissingMapping(q.qName);
    const mappingBtn = multiFile
        ? `<button type="button" class="section-q-action section-q-mapping-btn${missing ? ' text-danger' : ''}" data-qname="${_escAttr(q.qName)}" title="${missing ? 'Вопрос не соотнесён во всех файлах — нажмите, чтобы исправить' : 'Соотнести вручную'}"><i class="fa-solid fa-link"></i></button>`
        : '';
    const itemClass = missing ? ' q-item-missing' : '';
    return `
    <div class="section-question-item${itemClass}" data-qname="${_escAttr(q.qName)}" data-section-id="${_escAttr(sectionId)}">
        <span class="section-q-drag" title="Переместить"><i class="fa-solid fa-grip-lines"></i></span>
        <span class="flex-grow-1 text-truncate small fw-medium" title="${_escAttr(q.qName)}">${_escHtml(q.qName)}</span>
        ${mappingBtn}
        <button type="button" class="section-q-remove-btn" data-qname="${_escAttr(q.qName)}" data-section-id="${_escAttr(sectionId)}" title="Убрать из раздела">−</button>
    </div>`;
}

function _renderAvailableQuestions() {
    const container = document.getElementById('availableQuestionsList');
    const assigned = _getAssignedQNames();

    // Собираем все имена доноров которые уже склеены с кем-то
    const donorNames = new Set();
    Object.values(window.questionMerges || {}).forEach(donors => {
        donors.forEach(d => donorNames.add(d));
    });

    const available = _getAllStep3QNames().filter(q => !assigned.has(q) && !donorNames.has(q));

    if (!available.length) {
        container.innerHTML = '<p class="text-muted small text-center mt-2 mb-0">Все вопросы распределены по разделам</p>';
        return;
    }

    container.innerHTML = available.map(qName => `
        <div class="available-q-item" data-qname="${_escAttr(qName)}">
            <input type="checkbox" class="form-check-input flex-shrink-0 avail-q-cb cursor-pointer" data-qname="${_escAttr(qName)}">
            <span class="text-truncate small" title="${_escAttr(qName)}">${_escHtml(qName)}</span>
            <button type="button" class="btn btn-sm btn-outline-secondary ms-1 merge-question-btn" data-qname="${_escAttr(qName)}" title="Склеить с другим вопросом"><i class="fa-solid fa-link"></i></button>
        </div>`).join('');

    const selectAll = document.getElementById('selectAllAvailableQ');
    if (selectAll) selectAll.checked = false;

    container.querySelectorAll('.available-q-item').forEach(el => {
        el.setAttribute('draggable', 'true');
        el.addEventListener('click', e => {
            if (e.target.classList.contains('avail-q-cb')) return;
            const cb = el.querySelector('.avail-q-cb');
            if (cb) cb.checked = !cb.checked;
            _syncSelectAllAvailableQ();
        });
        el.addEventListener('dragstart', e => {
            el.classList.add('dragging');
            const cb = el.querySelector('.avail-q-cb');
            const checkedNames = Array.from(document.querySelectorAll('#availableQuestionsList .avail-q-cb:checked')).map(c => c.dataset.qname);
            window._step4DragNames = (cb && cb.checked && checkedNames.length > 0) ? checkedNames : [el.dataset.qname];
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/plain', window._step4DragNames.join('\n'));
        });
        el.addEventListener('dragend', () => el.classList.remove('dragging'));
        el.querySelector('.avail-q-cb').addEventListener('change', _syncSelectAllAvailableQ);
    });
}

function _syncSelectAllAvailableQ() {
    const all = Array.from(document.querySelectorAll('#availableQuestionsList .avail-q-cb'));
    const selectAll = document.getElementById('selectAllAvailableQ');
    if (!selectAll || !all.length) return;
    const checkedCount = all.filter(cb => cb.checked).length;
    selectAll.indeterminate = checkedCount > 0 && checkedCount < all.length;
    selectAll.checked = checkedCount === all.length;
}

document.getElementById('selectAllAvailableQ').addEventListener('change', function () {
    document.querySelectorAll('#availableQuestionsList .avail-q-cb').forEach(cb => { cb.checked = this.checked; });
});

function _addQNamesToSection(sectionId, qNames) {
    const sec = window.reportSections.find(s => s.id === sectionId);
    if (!sec) return;
    const assigned = _getAssignedQNames();
    let added = false;
    qNames.forEach(qName => {
        if (qName && !assigned.has(qName) && !sec.questions.find(q => q.qName === qName)) {
            sec.questions.push({ qName, visualize: false });
            assigned.add(qName);
            added = true;
        }
    });
    if (added) renderStep4();
}

// Drag-drop на зону раздела
document.getElementById('sectionsList').addEventListener('dragover', e => {
    if (!window._step4DragNames || !window._step4DragNames.length) return;
    const zone = e.target.closest('.section-questions-list');
    if (zone) { e.preventDefault(); zone.classList.add('drag-over'); }
});
document.getElementById('sectionsList').addEventListener('dragleave', e => {
    if (!window._step4DragNames || !window._step4DragNames.length) return;
    const zone = e.target.closest('.section-questions-list');
    if (zone && !zone.contains(e.relatedTarget)) zone.classList.remove('drag-over');
});
document.getElementById('sectionsList').addEventListener('drop', e => {
    if (!window._step4DragNames || !window._step4DragNames.length) return;
    const zone = e.target.closest('.section-questions-list');
    if (!zone) return;
    e.preventDefault();
    zone.classList.remove('drag-over');
    const names = window._step4DragNames;
    window._step4DragNames = [];
    _addQNamesToSection(zone.dataset.sectionId, names);
});

// Клики в панели разделов
document.getElementById('sectionsList').addEventListener('click', e => {
    const colBtn = e.target.closest('.section-collapse-btn');
    if (colBtn) {
        const secId = colBtn.dataset.sectionId;
        const wrapperEl = document.querySelector(`.section-questions-wrapper[data-section-id="${secId}"]`);
        const isNowCollapsed = wrapperEl.classList.toggle('collapsed');
        colBtn.querySelector('i').className = `fa-solid ${isNowCollapsed ? 'fa-chevron-up' : 'fa-chevron-down'}`;
        colBtn.title = isNowCollapsed ? 'Развернуть' : 'Свернуть';
        if (!window._collapsedSections) window._collapsedSections = new Set();
        if (isNowCollapsed) window._collapsedSections.add(secId);
        else window._collapsedSections.delete(secId);
        return;
    }
    const addBtn = e.target.closest('.add-qs-to-section-btn');
    if (addBtn) {
        const checked = Array.from(document.querySelectorAll('#availableQuestionsList .avail-q-cb:checked')).map(cb => cb.dataset.qname);
        if (!checked.length) { showToast('Выберите вопросы в правом списке', 'warning'); return; }
        _addQNamesToSection(addBtn.dataset.sectionId, checked);
        return;
    }
    const editBtn = e.target.closest('.edit-section-btn');
    if (editBtn) {
        const sec = window.reportSections.find(s => s.id === editBtn.dataset.sectionId);
        if (!sec) return;
        document.getElementById('editSectionId').value = sec.id;
        document.getElementById('editSectionName').value = sec.name;
        document.getElementById('editSectionName').classList.remove('is-invalid');
        document.getElementById('editSectionDesc').value = sec.description || '';
        document.getElementById('editSectionColor').value = sec.color || '#a0bce5';
        new bootstrap.Modal(document.getElementById('editSectionModal')).show();
        return;
    }
    const delBtn = e.target.closest('.delete-section-btn');
    if (delBtn) {
        const sec = window.reportSections.find(s => s.id === delBtn.dataset.sectionId);
        if (!sec) return;
        document.getElementById('deleteSectionNameLabel').textContent = `«${sec.name}»`;
        document.getElementById('confirmDeleteSectionBtn').dataset.sectionId = sec.id;
        new bootstrap.Modal(document.getElementById('confirmDeleteSectionModal')).show();
        return;
    }
    const remBtn = e.target.closest('.section-q-remove-btn');
    if (remBtn) {
        const sec = window.reportSections.find(s => s.id === remBtn.dataset.sectionId);
        if (sec) sec.questions = sec.questions.filter(q => q.qName !== remBtn.dataset.qname);
        renderStep4();
        return;
    }
    const mapBtn = e.target.closest('.section-q-mapping-btn');
    if (mapBtn) { openMappingModal(mapBtn.dataset.qname); return; }

    
});

// Модалка добавления раздела
document.getElementById('addSectionBtn').addEventListener('click', () => {
    document.getElementById('newSectionName').value = '';
    document.getElementById('newSectionName').classList.remove('is-invalid');
    document.getElementById('newSectionDesc').value = '';
    new bootstrap.Modal(document.getElementById('addSectionModal')).show();
});

document.getElementById('newSectionName').addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); document.getElementById('confirmAddSectionBtn').click(); }
});

document.getElementById('confirmAddSectionBtn').addEventListener('click', () => {
    const name = document.getElementById('newSectionName').value.trim();
    if (!name) { document.getElementById('newSectionName').classList.add('is-invalid'); return; }
    document.getElementById('newSectionName').classList.remove('is-invalid');
    window.reportSections.push({
        id: 'sec_' + (++_sectionIdCounter),
        name,
        description: document.getElementById('newSectionDesc').value.trim(),
        color: document.getElementById('newSectionColor').value,
        questions: []
    });
    bootstrap.Modal.getInstance(document.getElementById('addSectionModal')).hide();
    _renderSectionsList();
});

document.getElementById('confirmEditSectionBtn').addEventListener('click', () => {
    const name = document.getElementById('editSectionName').value.trim();
    if (!name) { document.getElementById('editSectionName').classList.add('is-invalid'); return; }
    document.getElementById('editSectionName').classList.remove('is-invalid');
    const secId = document.getElementById('editSectionId').value;
    const sec = window.reportSections.find(s => s.id === secId);
    if (sec) {
        sec.name = name;
        sec.description = document.getElementById('editSectionDesc').value.trim();
        sec.color = document.getElementById('editSectionColor').value;
    }
    bootstrap.Modal.getInstance(document.getElementById('editSectionModal')).hide();
    _renderSectionsList();
});

document.getElementById('confirmDeleteSectionBtn').addEventListener('click', function () {
    const secId = this.dataset.sectionId;
    window.reportSections = window.reportSections.filter(s => s.id !== secId);
    bootstrap.Modal.getInstance(document.getElementById('confirmDeleteSectionModal')).hide();
    renderStep4();
});

// ===================== СКЛЕЙКА ВОПРОСОВ =====================
let _mergeQuestionsSourceName = null;

// Делегирование на availableQuestionsList
document.getElementById('availableQuestionsList').addEventListener('click', (e) => {
    const btn = e.target.closest('.merge-question-btn');
    if (!btn) return;
    _mergeQuestionsSourceName = btn.dataset.qname;
    openMergeQuestionsModal(_mergeQuestionsSourceName);
});

function openMergeQuestionsModal(sourceName) {
    document.getElementById('mergeQuestionsSourceLabel').textContent = `«${sourceName}»`;
    document.getElementById('mergeQuestionsSearch').value = '';

    _renderAttachedDonors(sourceName);

    const allQNames = Array.from(
        document.querySelectorAll('#availableQuestionsList .available-q-item')
    ).map(el => el.dataset.qname).filter(n => n !== sourceName);

    _renderMergeQuestionsList(allQNames, '');

    // Переподвешиваем обработчик поиска (удаляем старый через clone)
    const searchEl = document.getElementById('mergeQuestionsSearch');
    const newSearch = searchEl.cloneNode(true);
    searchEl.parentNode.replaceChild(newSearch, searchEl);
    newSearch.addEventListener('input', function () {
        _renderMergeQuestionsList(allQNames, this.value.trim().toLowerCase());
    });

    new bootstrap.Modal(document.getElementById('mergeQuestionsModal')).show();
}

function _renderMergeQuestionsList(names, filter) {
    const container = document.getElementById('mergeQuestionsList');
    const filtered = filter ? names.filter(n => n.toLowerCase().includes(filter)) : names;

    if (!filtered.length) {
        container.innerHTML = '<p class="text-muted small text-center py-3 mb-0">Нет доступных вопросов</p>';
        return;
    }

    container.innerHTML = filtered.map(name => {
        // Показываем какие ответы уже склеены с этим вопросом (если есть)
        const merged = window.questionMerges && window.questionMerges[name];
        const mergedHint = merged && merged.length
            ? `<small class="text-muted ms-2">(уже склеен с: ${merged.slice(0,2).map(m => `«${_escHtml(m)}»`).join(', ')}${merged.length > 2 ? ` +${merged.length - 2}` : ''})</small>`
            : '';
        return `
        <div class="form-check py-1 border-bottom d-flex align-items-start gap-2">
            <input class="form-check-input flex-shrink-0 mt-1 merge-q-checkbox"
                   type="checkbox" id="mqcb_${_escAttr(name)}" value="${_escAttr(name)}">
            <label class="form-check-label small fw-medium cursor-pointer flex-grow-1"
                   for="mqcb_${_escAttr(name)}">${_escHtml(name)}${mergedHint}</label>
        </div>`;
    }).join('');
}

document.getElementById('applyMergeQuestionsBtn').addEventListener('click', () => {
    const checked = Array.from(
        document.querySelectorAll('#mergeQuestionsList .merge-q-checkbox:checked')
    ).map(cb => cb.value);

    if (!checked.length) {
        showToast('Выберите хотя бы один вопрос для склейки', 'warning');
        return;
    }

    // Сохраняем маппинг склейки: sourceName -> [donorName, ...]
    if (!window.questionMerges) window.questionMerges = {};
    if (!window.questionMerges[_mergeQuestionsSourceName]) {
        window.questionMerges[_mergeQuestionsSourceName] = [];
    }
    checked.forEach(name => {
        if (!window.questionMerges[_mergeQuestionsSourceName].includes(name)) {
            window.questionMerges[_mergeQuestionsSourceName].push(name);
        }
    });

    // Убираем доноров из availableQuestionsList
    checked.forEach(name => {
        const item = document.querySelector(`#availableQuestionsList .available-q-item[data-qname="${CSS.escape(name)}"]`);
        if (item) item.remove();
    });

    // Обновляем иконку на кнопке источника — показываем что есть склейка
    _updateMergeIndicator(_mergeQuestionsSourceName);

    bootstrap.Modal.getInstance(document.getElementById('mergeQuestionsModal')).hide();
    showToast(`Склеено вопросов: ${checked.length}`, 'success');
});

function _updateMergeIndicator(sourceName) {
    const item = document.querySelector(
        `#availableQuestionsList .available-q-item[data-qname="${CSS.escape(sourceName)}"]`
    );
    if (!item) return;
    const btn = item.querySelector('.merge-question-btn');
    if (!btn) return;
    const merges = window.questionMerges && window.questionMerges[sourceName];
    if (merges && merges.length) {
        btn.classList.remove('btn-outline-secondary');
        btn.classList.add('btn-warning');
        btn.title = `Склеен с: ${merges.map(m => `«${m}»`).join(', ')} (нажмите, чтобы изменить)`;
    } else {
        btn.classList.add('btn-outline-secondary');
        btn.classList.remove('btn-warning');
        btn.title = 'Склеить с другим вопросом';
    }
}
// Обработчики случайного цвета
document.getElementById('randomNewSectionColor').addEventListener('click', () => {
    document.getElementById('newSectionColor').value = randomColor();
});
document.getElementById('randomEditSectionColor').addEventListener('click', () => {
    document.getElementById('editSectionColor').value = randomColor();
});


function _renderAttachedDonors(sourceName) {
    // Ищем или создаём блок прикреплённых доноров внутри модалки
    let attachedBlock = document.getElementById('mergeQuestionsAttached');
    if (!attachedBlock) {
        // Вставляем перед строкой поиска
        document.getElementById('mergeQuestionsSearch').insertAdjacentHTML('beforebegin', `
            <div id="mergeQuestionsAttached" class="mb-2"></div>
        `);
        attachedBlock = document.getElementById('mergeQuestionsAttached');
    }

    const donors = (window.questionMerges && window.questionMerges[sourceName]) || [];
    if (!donors.length) {
        attachedBlock.innerHTML = '';
        return;
    }

    attachedBlock.innerHTML = `
        <div class="small fw-semibold text-muted mb-1"><i class="fa-solid fa-paperclip me-1"></i>Уже прикреплены:</div>
        ${donors.map(d => `
            <div class="d-flex align-items-center justify-content-between py-1 border-bottom gap-2 attached-donor-row" data-donor="${_escAttr(d)}">
                <span class="small text-truncate flex-grow-1" title="${_escAttr(d)}">${_escHtml(d)}</span>
                <button type="button" class="btn btn-sm btn-outline-danger py-0 px-1 detach-donor-btn flex-shrink-0"
                        data-source="${_escAttr(sourceName)}" data-donor="${_escAttr(d)}" title="Открепить вопрос">
                    <i class="fa-solid fa-xmark"></i>
                </button>
            </div>`).join('')}
    `;
}

document.getElementById('mergeQuestionsModal').addEventListener('click', (e) => {
    const btn = e.target.closest('.detach-donor-btn');
    if (!btn) return;

    const sourceName = btn.dataset.source;
    const donorName = btn.dataset.donor;

    // Удаляем донора из массива склеек
    if (window.questionMerges && window.questionMerges[sourceName]) {
        window.questionMerges[sourceName] = window.questionMerges[sourceName].filter(d => d !== donorName);
        if (window.questionMerges[sourceName].length === 0) {
            delete window.questionMerges[sourceName];
        }
    }

    // Возвращаем донора в правую панель
    const container = document.getElementById('availableQuestionsList');
    // Проверяем что его там нет
    const alreadyThere = container.querySelector(`.available-q-item[data-qname="${CSS.escape(donorName)}"]`);
    if (!alreadyThere) {
        container.insertAdjacentHTML('beforeend', `
            <div class="available-q-item" data-qname="${_escAttr(donorName)}">
                <input type="checkbox" class="form-check-input flex-shrink-0 avail-q-cb cursor-pointer" data-qname="${_escAttr(donorName)}">
                <span class="text-truncate small" title="${_escAttr(donorName)}">${_escHtml(donorName)}</span>
                <button type="button" class="btn btn-sm btn-outline-secondary ms-1 merge-question-btn" data-qname="${_escAttr(donorName)}" title="Склеить с другим вопросом"><i class="fa-solid fa-link"></i></button>
            </div>`);
        // Восстанавливаем draggable и обработчики
        const newItem = container.querySelector(`.available-q-item[data-qname="${CSS.escape(donorName)}"]`);
        if (newItem) {
            newItem.setAttribute('draggable', 'true');
            newItem.addEventListener('click', e => {
                if (e.target.classList.contains('avail-q-cb')) return;
                const cb = newItem.querySelector('.avail-q-cb');
                if (cb) cb.checked = !cb.checked;
                _syncSelectAllAvailableQ();
            });
            newItem.addEventListener('dragstart', e => {
                newItem.classList.add('dragging');
                window._step4DragNames = [donorName];
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text/plain', donorName);
            });
            newItem.addEventListener('dragend', () => newItem.classList.remove('dragging'));
            newItem.querySelector('.avail-q-cb').addEventListener('change', _syncSelectAllAvailableQ);
        }
    }

    // Обновляем индикатор кнопки источника
    _updateMergeIndicator(sourceName);

    // Перерисовываем блок прикреплённых
    _renderAttachedDonors(sourceName);
});