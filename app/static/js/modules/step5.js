// ===================== ШАГ 5: ВИЗУАЛИЗАЦИЯ =====================
window.chartsData = {};
window.stackedChartsData = {};
window.pieChartsData = {};
window.renderedTabs = {};

// Collapse report sections
document.getElementById('reportContent').addEventListener('click', e => {
    const colBtn = e.target.closest('.report-section-collapse-btn');
    if (!colBtn) return;
    const secId = colBtn.dataset.sectionId;
    const wrapper = document.querySelector(`.report-section-wrapper[data-section-id="${secId}"]`);
    if (!wrapper) return;
    const isNowCollapsed = wrapper.classList.toggle('collapsed');
    colBtn.querySelector('i').className = `fa-solid ${isNowCollapsed ? 'fa-chevron-up' : 'fa-chevron-down'}`;
    colBtn.title = isNowCollapsed ? 'Развернуть' : 'Свернуть';
    if (!window._collapsedReportSections) window._collapsedReportSections = new Set();
    if (isNowCollapsed) window._collapsedReportSections.add(secId);
    else window._collapsedReportSections.delete(secId);
});

document.getElementById('analyzeBtn').addEventListener('click', async () => {

    const fileLabels = {};
    const fileColors = {};
    if (window.processedFiles && window.processedFiles.length === 1) {
        const f = window.processedFiles[0];
        fileLabels[f.clean_filename] = f.original_name.replace(/\.[^.]+$/, '');
        fileColors[f.clean_filename] = defaultColors[0];
    } else {
        document.querySelectorAll('.legend-label').forEach(el => { fileLabels[el.dataset.file] = el.value || el.placeholder; });
        document.querySelectorAll('.legend-color').forEach(el => { fileColors[el.dataset.file] = el.value; });
    }

    const ALL_VIZ = ['Таблица', 'Столбчатая диаграмма', 'Накопленная диаграмма', 'Круговая диаграмма'];
    const configs = [];
    const addedQNames = new Set();

    // Сначала вопросы из разделов
    (window.reportSections || []).forEach(sec => {
        sec.questions.forEach(q => {
            if (addedQNames.has(q.qName)) return;
            addedQNames.add(q.qName);
            const mergedColumns = (window.questionMerges && window.questionMerges[q.qName]) || [];
            configs.push({
                column: q.qName,
                viz_type: ALL_VIZ,
                file_mapping: window.questionMapping[q.qName],
                merged_columns: mergedColumns,
                section_id: sec.id
            });
        });
    });
    
    // Затем нераспределённые вопросы (справа на шаге 4)
    document.querySelectorAll('#availableQuestionsList .available-q-item').forEach(el => {
        const qName = el.dataset.qname;
        if (addedQNames.has(qName)) return;
        addedQNames.add(qName);
        const mergedColumns = (window.questionMerges && window.questionMerges[qName]) || [];
        configs.push({
            column: qName,
            viz_type: ALL_VIZ,
            file_mapping: window.questionMapping[qName],
            merged_columns: mergedColumns,
            section_id: null
        });
    });

    if (!configs.length) {
        showToast('Нет вопросов. Выберите вопросы на шаге 3.', 'warning');
        return;
    }

    const payload = { file_labels: fileLabels, file_colors: fileColors, configs: configs };

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
            let globalItemIdx = 0;

            const renderItem = (item, container) => {
                const id = `item_${globalItemIdx++}`;
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
                window.appData[id].includeInReport = true;

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

                const html = `
                <div class="d-flex align-items-center gap-3 mb-2">
                    <h5 class="fw-semibold text-dark mb-0 flex-grow-1">${item.col_name}</h5>
                    <div class="form-check mb-0 ui-system-font flex-shrink-0">
                        <input class="form-check-input viz-include-cb" type="checkbox" id="vizinclude_${id}" data-id="${id}" checked>
                        <label class="form-check-label small fw-medium text-muted" for="vizinclude_${id}">
                            <i class="fa-solid fa-file-export me-1"></i>В отчёт
                        </label>
                    </div>
                </div>
                <div class="result-item">
                    <ul class="nav nav-tabs ui-system-font" id="tabs_${id}">
                        <li class="nav-item"><button class="nav-link active viz-tab-btn" data-id="${id}" data-tab="table"><i class="fa-solid fa-table me-1"></i>Таблица</button></li>
                        <li class="nav-item"><button class="nav-link viz-tab-btn" data-id="${id}" data-tab="bar"><i class="fa-solid fa-chart-column me-1"></i>Столбчатая</button></li>
                        <li class="nav-item"><button class="nav-link viz-tab-btn" data-id="${id}" data-tab="stacked"><i class="fa-solid fa-chart-bar me-1"></i>Накопленная</button></li>
                        <li class="nav-item"><button class="nav-link viz-tab-btn" data-id="${id}" data-tab="pie"><i class="fa-solid fa-chart-pie me-1"></i>Круговая</button></li>
                    </ul>
                    <div class="ui-system-font bg-white p-2 mb-3 border border-top-0 border-secondary-subtle rounded-bottom d-flex flex-wrap gap-3 align-items-center shadow-sm" id="settings_${id}">
                        <div class="form-check mb-0" data-vis-tabs="table">
                            <input class="form-check-input setting-show-total" type="checkbox" id="total_${id}" data-id="${id}" checked>
                            <label class="form-check-label small fw-medium" for="total_${id}"><i class="fa-solid fa-sigma me-1 text-muted"></i>Добавить строку "Всего"</label>
                        </div>
                        <div class="form-check form-switch mb-0" data-vis-tabs="table bar">
                            <input class="form-check-input setting-highlight-top" type="checkbox" id="hl_${id}" data-id="${id}">
                            <label class="form-check-label small fw-medium" for="hl_${id}"><i class="fa-solid fa-trophy me-1 text-muted"></i>Выделить топ:</label>
                        </div>
                        <input type="number" class="form-control form-control-sm setting-top-n input-w-70" data-id="${id}" value="1" min="1" max="${item.data.length * item.file_keys.length}" data-vis-tabs="table bar">
                        <input type="color" class="form-control form-control-color setting-highlight-color color-input-sm" data-id="${id}" value="#dc3545" title="Цвет выделения топа" data-vis-tabs="table bar">
                        <button type="button" class="btn btn-sm btn-outline-secondary random-highlight-color-btn" data-id="${id}" title="Случайный цвет" data-vis-tabs="table bar"><i class="fa-solid fa-dice-five"></i></button>
                        <div class="vr" data-vis-tabs="table bar"></div>
                        <div class="form-check form-switch mb-0" data-vis-tabs="table bar stacked">
                            <input class="form-check-input setting-vertical" type="checkbox" id="vert_${id}" data-id="${id}">
                            <label class="form-check-label small fw-medium" for="vert_${id}"><i class="fa-solid fa-rotate me-1 text-muted"></i>Вертикальный</label>
                        </div>
                        <div class="form-check mb-0" data-vis-tabs="bar stacked pie">
                            <input class="form-check-input setting-show-legend" type="checkbox" id="legend_${id}" data-id="${id}" ${item.file_keys.length > 1 ? 'checked' : ''}>
                            <label class="form-check-label small fw-medium" for="legend_${id}"><i class="fa-solid fa-list-ul me-1 text-muted"></i>Легенда</label>
                        </div>
                        <button type="button" class="btn btn-sm btn-outline-secondary ms-auto" onclick="openHideColModal('${id}')" title="Скрыть столбцы" data-vis-tabs="table"><i class="fa-solid fa-eye-slash"></i></button>
                        <button type="button" class="btn btn-sm btn-outline-secondary" onclick="openChartEditModal('${id}')" title="Редактировать данные диаграммы" data-vis-tabs="bar stacked pie"><i class="fa-solid fa-pen"></i></button>
                        <button type="button" class="btn btn-sm btn-outline-secondary" onclick="confirmFuzzyMapping('${id}')" title="Сгруппировать похожие ответы"><i class="fa-solid fa-shuffle"></i></button>
                    </div>
                    <div id="pane_table_${id}">
                        <div class="mb-3">Таблица ${tNum} – Распределение ответов респондентов на вопрос: «${item.col_name}»</div>
                        <div class="table-responsive mb-4">
                            <table class="table table-bordered table-hover table-custom-border align-middle mb-0" id="${id}">
                                <thead></thead>
                                <tbody></tbody>
                                <tfoot></tfoot>
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

                container.insertAdjacentHTML('beforeend', html);
                document.getElementById(`settings_${id}`)?.querySelectorAll('[data-vis-tabs]').forEach(el => {
                    el.classList.toggle('d-none', !el.dataset.visTabs.split(' ').includes('table'));
                });
                renderTable(id);
            };

            const resultsByName = {};
            data.results.forEach(item => {
                if (item && item.col_name) resultsByName[item.col_name] = item;
            });

            const assignedQNames = new Set((window.reportSections || []).flatMap(s => s.questions.map(q => q.qName)));

            (window.reportSections || []).forEach(sec => {
                if (!sec.questions || sec.questions.length === 0) return;

                const rscCollapsed = window._collapsedReportSections && window._collapsedReportSections.has(sec.id);
                const secColor = sec.color || '#a0bce5';
                reportContent.insertAdjacentHTML('beforeend', `
                    <div class="report-section-card mb-4" style="border-left: 3px solid ${secColor}; padding-left: 8px;">
                        <div class="report-section-header">
                            <i class="fa-solid fa-layer-group me-2" style="color:${secColor};"></i>
                            <span class="flex-grow-1" style="color:${secColor};">${_escHtml(sec.name)}</span>
                            <button type="button" class="report-section-collapse-btn" data-section-id="${_escAttr(sec.id)}" title="${rscCollapsed ? 'Развернуть' : 'Свернуть'}">
                                <i class="fa-solid ${rscCollapsed ? 'fa-chevron-up' : 'fa-chevron-down'}"></i>
                            </button>
                        </div>
                        <div class="report-section-wrapper${rscCollapsed ? ' collapsed' : ''}" data-section-id="${_escAttr(sec.id)}">
                            <div class="report-section-body" id="rsc_${_escAttr(sec.id)}"></div>
                        </div>
                    </div>`);
                const sectionBody = document.getElementById(`rsc_${sec.id}`);

                sec.questions.forEach(q => {
                    const item = resultsByName[q.qName];
                    if (!item || !sectionBody) return;
                    renderItem(item, sectionBody);
                });
            });

            const unassignedItems = configs
                .filter(cfg => !assignedQNames.has(cfg.column))
                .map(cfg => resultsByName[cfg.column])
                .filter(Boolean);

            if (unassignedItems.length) {
                reportContent.insertAdjacentHTML('beforeend', `
                    <div class="report-section-card mb-4" style="border-left: 3px solid #6c757d; padding-left: 8px;">
                        <div class="report-section-header" style="color:#6c757d;">
                            <i class="fa-solid fa-layer-group me-2"></i>
                            <span class="flex-grow-1">Без раздела</span>
                        </div>
                        <div class="report-section-body" id="rsc_unassigned"></div>
                    </div>`);
                const unassignedBody = document.getElementById('rsc_unassigned');
                unassignedItems.forEach(item => renderItem(item, unassignedBody));
            }

        } else { showToast(data.message, 'danger'); }
    } catch (err) { showToast('Ошибка соединения с сервером', 'danger'); }
    finally {
        document.getElementById('analyzeBtn').disabled = false;
        document.getElementById('analyzeSpinner').classList.add('d-none');
    }
});

// ===================== VIZ INTERACTION HANDLERS =====================
document.addEventListener('click', e => {
    const rhlBtn = e.target.closest('.random-highlight-color-btn');
    if (rhlBtn) {
        const input = document.querySelector(`.setting-highlight-color[data-id="${rhlBtn.dataset.id}"]`);
        if (input) { input.value = randomColor(); input.dispatchEvent(new Event('change', { bubbles: true })); }
        return;
    }

    const rpcBtn = e.target.closest('.random-pie-colors-btn');
    if (rpcBtn) {
        const did = rpcBtn.dataset.id;
        const dataObj = window.appData[did];
        if (dataObj) {
            dataObj.pieColors = dataObj.pieColors.map(() => randomColor());
            window.renderedTabs[did].pie = false;
            drawPieChart(did);
        }
        return;
    }

    const rbcBtn = e.target.closest('.random-bar-colors-btn');
    if (rbcBtn) {
        const did = rbcBtn.dataset.id;
        const dataObj = window.appData[did];
        if (dataObj) {
            dataObj.barColors = dataObj.barColors.map(() => randomColor());
            drawChart(did);
        }
        return;
    }

    const rLegBtn = e.target.closest('.random-legend-color-btn');
    if (rLegBtn) {
        const input = document.querySelector(`.legend-color[data-file="${rLegBtn.dataset.file}"]`);
        if (input) input.value = randomColor();
        return;
    }

    const mergeBtn = e.target.closest('.btn-merge-answer');
    if (mergeBtn) {
        const id = mergeBtn.dataset.id;
        const srcIdx = parseInt(mergeBtn.dataset.index);
        const dataObj = window.appData[id];
        window._mergeState = { id, srcIdx };
        document.getElementById('mergeSourceLabel').textContent = dataObj.data[srcIdx].answer;
        const sel = document.getElementById('mergeTargetSelect');
        if ($(sel).hasClass('select2-hidden-accessible')) $(sel).select2('destroy');
        $(sel).empty();
        dataObj.data.forEach((r, i) => {
            if (i === srcIdx) return;
            $(sel).append(new Option(r.answer, i, false, false));
        });
        $(sel).select2({ theme: 'bootstrap-5', width: '100%', language: 'ru', dropdownParent: $('#mergeAnswerModal') });
        new bootstrap.Modal(document.getElementById('mergeAnswerModal')).show();
        return;
    }

    const btn = e.target.closest('.viz-tab-btn');
    if (!btn) return;
    const id = btn.dataset.id;
    const tab = btn.dataset.tab;

    document.querySelectorAll(`#tabs_${id} .viz-tab-btn`).forEach(b => b.classList.remove('active'));
    btn.classList.add('active');

    ['table', 'bar', 'stacked', 'pie'].forEach(t => {
        const pane = document.getElementById(`pane_${t}_${id}`);
        if (pane) pane.classList.toggle('d-none', t !== tab);
    });

    const settingsEl = document.getElementById(`settings_${id}`);
    if (settingsEl) {
        settingsEl.querySelectorAll('[data-vis-tabs]').forEach(el => {
            const tabs = el.dataset.visTabs.split(' ');
            el.classList.toggle('d-none', !tabs.includes(tab));
        });
    }

    if (tab === 'bar' && !window.renderedTabs[id].bar) {
        window.renderedTabs[id].bar = true;
        setTimeout(() => drawChart(id), 50);
    }
    if (tab === 'stacked' && !window.renderedTabs[id].stacked) {
        window.renderedTabs[id].stacked = true;
        setTimeout(() => drawStackedChart(id), 50);
    }
    if (tab === 'pie' && !window.renderedTabs[id].pie) {
        window.renderedTabs[id].pie = true;
        setTimeout(() => drawPieChart(id), 50);
    }
});

document.getElementById('reportContent').addEventListener('change', e => {
    if (e.target.classList.contains('viz-include-cb')) {
        const id = e.target.dataset.id;
        if (window.appData[id]) {
            window.appData[id].includeInReport = e.target.checked;
        }
    }
});
