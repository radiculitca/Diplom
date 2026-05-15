// ===================== CHART.JS INIT =====================
Chart.register(ChartDataLabels);
Chart.defaults.font.family = '"Times New Roman", Times, serif';
Chart.defaults.font.size = 12;
Chart.defaults.color = '#000';

const PIE_COLORS = [
    '#dc3545','#0d6efd','#198754','#ffc107','#6f42c1',
    '#fd7e14','#20c997','#0dcaf0','#6c757d','#343a40',
    '#e15759','#4e79a7','#59a14f','#edc948','#b07aa1',
    '#003f5c','#2f4b7c','#665191','#a05195','#d45087',
    '#f95d6a','#ff7c43','#ffa600','#b5bd00','#00b050'
];

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
                    <div class="row-toggle-container row-toggle-visible">
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

// ===================== HIDE COLUMNS MODAL =====================
let currentHideColTableId = null;
window.openHideColModal = function(id) {
    currentHideColTableId = id;
    const opt = window.appData[id].options.hiddenCol || 'none';
    document.getElementById('hideColNone').checked = (opt === 'none');
    document.getElementById('hideColCount').checked = (opt === 'count');
    document.getElementById('hideColPct').checked = (opt === 'percent');
    new bootstrap.Modal(document.getElementById('hideColModal')).show();
};

document.getElementById('applyMergeBtn').addEventListener('click', () => {
    if (!window._mergeState) return;
    const { id, srcIdx } = window._mergeState;
    const targetIdx = parseInt($('#mergeTargetSelect').val());
    if (isNaN(targetIdx)) return;
    const dataObj = window.appData[id];
    const srcRow = dataObj.data[srcIdx];
    const tgtRow = dataObj.data[targetIdx];

    dataObj.file_keys.forEach(fk => {
        tgtRow.counts[fk] = (tgtRow.counts[fk] || 0) + (srcRow.counts[fk] || 0);
    });
    tgtRow._total = Object.values(tgtRow.counts).reduce((a, b) => a + b, 0);

    dataObj.data.splice(srcIdx, 1);
    if (Array.isArray(dataObj.pieColors) && dataObj.pieColors.length > dataObj.data.length) dataObj.pieColors.splice(srcIdx, 1);
    if (Array.isArray(dataObj.barColors) && dataObj.barColors.length > dataObj.data.length) dataObj.barColors.splice(srcIdx, 1);

    window._mergeState = null;
    bootstrap.Modal.getInstance(document.getElementById('mergeAnswerModal')).hide();
    renderTable(id);
    drawChart(id);
    drawStackedChart(id);
    drawPieChart(id);
});

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

// ===================== PIE CHART =====================
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
                    tooltip: {
                        enabled: true,
                        displayColors: false,
                        callbacks: {
                            title: (items) => activeData[items[0].dataIndex]?.answer || '',
                            label: (ctx) => ctx.raw
                        }
                    },
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

// ===================== BAR CHART =====================
function drawChart(id) {
    const dataObj = window.appData[id];
    if (!dataObj || !window.chartsData[id]) return;

    if (window.charts[id]) window.charts[id].destroy();

    const activeData = dataObj.data.filter(r => r.included).sort((a, b) => b._total - a._total);
    const canvasBar = document.getElementById(`canvas_${id}`);
    if (activeData.length === 0) {
        if (canvasBar) {
            canvasBar.style.display = 'none';
            const wrap = canvasBar.closest('.chart-container');
            if (wrap && !wrap.querySelector('.no-data-msg')) wrap.insertAdjacentHTML('beforeend', '<p class="no-data-msg text-muted text-center py-4 small mb-0">Нет данных для отображения</p>');
        }
        return;
    }
    if (canvasBar) {
        canvasBar.style.display = '';
        canvasBar.closest('.chart-container')?.querySelector('.no-data-msg')?.remove();
    }

    const labels = activeData.map(r => r.answer.length > 50 ? r.answer.substring(0, 50) + '...' : r.answer);
    const isHorizontal = dataObj.options.chartDirection === 'y';
    const topN = dataObj.options.highlightTop ? Math.min(dataObj.options.topN, activeData.length * dataObj.file_keys.length) : 0;
    const HIGHLIGHT_COLOR = dataObj.options.highlightColor || '#0d6efd';

    const fileTotals = {};
    dataObj.file_keys.forEach(fk => {
        fileTotals[fk] = activeData.reduce((sum, r) => sum + (r.counts[fk] || 0), 0);
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
        data: { labels: labels, datasets: datasets },
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
                            return {
                                text: ds.label,
                                fillStyle: dataObj.file_colors[fk],
                                strokeStyle: dataObj.file_colors[fk],
                                lineWidth: 1,
                                hidden: !chart.isDatasetVisible(i),
                                datasetIndex: i
                            };
                        })
                    }
                },
                tooltip: {
                    enabled: true,
                    displayColors: datasets.length > 1,
                    callbacks: {
                        title: (items) => activeData[items[0].dataIndex]?.answer || '',
                        label: (ctx) => (datasets.length > 1 ? ctx.dataset.label + ': ' : '') + ctx.dataset.rawCounts[ctx.dataIndex]
                    }
                },
                datalabels: {
                    color: '#000',
                    anchor: 'end',
                    align: isHorizontal ? 'right' : 'top',
                    offset: 4,
                    font: { family: '"Times New Roman", Times, serif', size: 14, weight: 'bold' },
                    formatter: (value) => {
                        if (!value || value === 0) return '';
                        if (value > 0 && value < 1) return '<1%';
                        return Math.round(value) + '%';
                    }
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

// ===================== STACKED BAR CHART =====================
function drawStackedChart(id) {
    const dataObj = window.appData[id];
    if (!dataObj || !window.stackedChartsData[id]) return;

    if (window.charts['stacked_' + id]) window.charts['stacked_' + id].destroy();

    const activeData = dataObj.data.filter(r => r.included).sort((a, b) => b._total - a._total);
    const canvasStacked = document.getElementById(`stacked_canvas_${id}`);
    if (activeData.length === 0) {
        if (canvasStacked) {
            canvasStacked.style.display = 'none';
            const wrap = canvasStacked.closest('.stacked-container');
            if (wrap && !wrap.querySelector('.no-data-msg')) wrap.insertAdjacentHTML('beforeend', '<p class="no-data-msg text-muted text-center py-4 small mb-0">Нет данных для отображения</p>');
        }
        return;
    }
    if (canvasStacked) {
        canvasStacked.style.display = '';
        canvasStacked.closest('.stacked-container')?.querySelector('.no-data-msg')?.remove();
    }

    const fileTotals = {};
    dataObj.file_keys.forEach(fk => {
        fileTotals[fk] = activeData.reduce((sum, r) => sum + (r.counts[fk] || 0), 0);
    });

    const labels = activeData.map(r => r.answer.length > 50 ? r.answer.substring(0, 50) + '...' : r.answer);

    const datasets = dataObj.file_keys.map(fileKey => {
        const actualPcts = activeData.map(r => {
            const count = r.counts[fileKey] || 0;
            const ft = fileTotals[fileKey];
            return ft > 0 ? (count / ft) * 100 : 0;
        });
        const actualCounts = activeData.map(r => r.counts[fileKey] || 0);
        return {
            label: dataObj.file_labels[fileKey],
            backgroundColor: dataObj.file_colors[fileKey],
            actualPcts,
            actualCounts,
            data: activeData.map((r, rIdx) => {
                const pct = actualPcts[rIdx];
                const answerSum = dataObj.file_keys.reduce((sum, fk) => {
                    const c = r.counts[fk] || 0;
                    const ft = fileTotals[fk];
                    return sum + (ft > 0 ? (c / ft) * 100 : 0);
                }, 0);
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
                legend: {
                    display: !!dataObj.options.showLegend,
                    position: 'bottom',
                    labels: { font: { family: '"Times New Roman", Times, serif', size: 12 } }
                },
                tooltip: {
                    enabled: true,
                    displayColors: datasets.length > 1,
                    callbacks: {
                        title: (items) => activeData[items[0].dataIndex]?.answer || '',
                        label: (ctx) => (datasets.length > 1 ? ctx.dataset.label + ': ' : '') + ctx.dataset.actualCounts[ctx.dataIndex]
                    }
                },
                datalabels: {
                    color: '#fff',
                    anchor: 'center',
                    align: 'center',
                    font: { family: '"Times New Roman", Times, serif', size: 13, weight: 'bold' },
                    formatter: (value, context) => {
                        const pct = context.dataset.actualPcts[context.dataIndex];
                        if (!pct || pct === 0) return '';
                        const rounded = Math.round(pct);
                        return rounded < 1 ? '<1%' : rounded + '%';
                    }
                }
            },
            scales: isVertStacked
                ? {
                    x: { stacked: true, grid: { display: false }, border: { display: false } },
                    y: { stacked: true, display: false, max: 100, grid: { display: false }, border: { display: false } }
                }
                : {
                    x: { stacked: true, display: false, max: 100, grid: { display: false }, border: { display: false } },
                    y: { stacked: true, grid: { display: false }, border: { display: false } }
                }
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
        totals[fk] = activeRows.reduce((sum, r) => sum + (r.counts[fk] || 0), 0);
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
                </tr>
            `;
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
                </tr>
            `;
        }
        tableEl.querySelector('thead').innerHTML = theadHtml;

        let tbodyHtml = '';
        const lastIdx = dataObj.data.length - 1;
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

            const addBtn = idx === lastIdx
                ? `<button type="button" class="add-row-btn" data-id="${tableId}" data-index="${idx}" title="Добавить строку ниже"><i class="fa-solid fa-plus"></i></button>`
                : '';

            tbodyHtml += `
                <tr class="${trClass}"${trStyle}>
                    <td class="control-col">
                        <div class="row-toggle-container">
                            <div class="form-check form-switch mb-0">
                                <input class="form-check-input row-toggle ui-system-font" type="checkbox" ${row.included ? 'checked' : ''} data-id="${tableId}" data-index="${idx}" title="Включить/исключить ответ">
                            </div>
                            <button type="button" class="btn-merge-answer" data-id="${tableId}" data-index="${idx}" title="Слить с другим ответом"><i class="fa-solid fa-plus"></i></button>
                        </div>
                    </td>
                    <td class="text-start align-middle" style="position:relative"><span contenteditable="true" class="editable-cell answer-text" data-id="${tableId}" data-index="${idx}">${row.answer}</span>${addBtn}</td>
                    ${tdHtml}
                </tr>`;
        });
        if (!tbodyHtml) {
            const colSpan = isSingleFile ? 4 : 2 + dataObj.file_keys.length * 2;
            tbodyHtml = `<tr><td colspan="${colSpan}" class="text-center text-muted py-3 small">Нет данных</td></tr>`;
        }
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

// ===================== TABLE SETTINGS EVENT HANDLERS =====================
document.addEventListener('change', (e) => {
    if (e.target.classList.contains('setting-show-total')) {
        window.appData[e.target.dataset.id].options.showTotal = e.target.checked;
        renderTable(e.target.dataset.id);
    }
    if (e.target.classList.contains('setting-highlight-top')) {
        window.appData[e.target.dataset.id].options.highlightTop = e.target.checked;
        drawChart(e.target.dataset.id);
        renderTable(e.target.dataset.id);
    }
    if (e.target.classList.contains('setting-vertical')) {
        const id = e.target.dataset.id;
        window.appData[id].options.tableVertical = e.target.checked;
        window.appData[id].options.chartDirection = e.target.checked ? 'x' : 'y';
        renderTable(id);
        drawChart(id);
        drawStackedChart(id);
    }
    if (e.target.classList.contains('setting-highlight-color')) {
        window.appData[e.target.dataset.id].options.highlightColor = e.target.value;
        drawChart(e.target.dataset.id);
        renderTable(e.target.dataset.id);
    }
    if (e.target.classList.contains('pie-answer-color')) {
        const id = e.target.dataset.id;
        window.appData[id].pieColors[parseInt(e.target.dataset.index)] = e.target.value;
        drawPieChart(id);
    }
    if (e.target.classList.contains('bar-answer-color')) {
        const id = e.target.dataset.id;
        window.appData[id].barColors[parseInt(e.target.dataset.index)] = e.target.value;
        drawChart(id);
    }
    if (e.target.classList.contains('row-toggle')) {
        window.appData[e.target.dataset.id].data[parseInt(e.target.dataset.index)].included = e.target.checked;
        renderTable(e.target.dataset.id);
        drawChart(e.target.dataset.id);
        drawStackedChart(e.target.dataset.id);
        drawPieChart(e.target.dataset.id);
    }
    if (e.target.classList.contains('modal-row-toggle')) {
        const id = e.target.dataset.id;
        window.appData[id].data[parseInt(e.target.dataset.index)].included = e.target.checked;
        renderChartEditModal(id);
        renderTable(id);
        drawChart(id);
        drawStackedChart(id);
        drawPieChart(id);
    }
    if (e.target.classList.contains('setting-show-legend')) {
        const id = e.target.dataset.id;
        window.appData[id].options.showLegend = e.target.checked;
        drawChart(id);
        drawStackedChart(id);
        drawPieChart(id);
    }
});

document.addEventListener('input', (e) => {
    if (e.target.classList.contains('setting-top-n')) {
        const val = parseInt(e.target.value);
        if (!isNaN(val) && val >= 1) {
            window.appData[e.target.dataset.id].options.topN = val;
            if (window.appData[e.target.dataset.id].options.highlightTop) {
                drawChart(e.target.dataset.id);
                renderTable(e.target.dataset.id);
            }
        }
    }
    if (e.target.classList.contains('answer-text')) {
        window.appData[e.target.dataset.id].data[parseInt(e.target.dataset.index)].answer = e.target.innerText;
    }
    if (e.target.classList.contains('answer-count')) {
        const val = parseInt(e.target.innerText);
        if (!isNaN(val)) {
            const fk = e.target.getAttribute('data-file');
            const rowData = window.appData[e.target.dataset.id].data[parseInt(e.target.dataset.index)];
            rowData.counts[fk] = val;
            rowData._total = Object.values(rowData.counts).reduce((a, b) => a + b, 0);
        }
    }
    if (e.target.hasAttribute('data-header')) {
        window.appData[e.target.dataset.id].headers[e.target.getAttribute('data-header')] = e.target.innerText;
    }
    if (e.target.classList.contains('modal-answer-text')) {
        window.appData[e.target.dataset.id].data[parseInt(e.target.dataset.index)].answer = e.target.innerText;
    }
    if (e.target.classList.contains('modal-answer-count')) {
        const val = parseInt(e.target.innerText);
        if (!isNaN(val)) {
            const fk = e.target.getAttribute('data-file');
            const rowData = window.appData[e.target.dataset.id].data[parseInt(e.target.dataset.index)];
            rowData.counts[fk] = val;
            rowData._total = Object.values(rowData.counts).reduce((a, b) => a + b, 0);
        }
    }
});

document.addEventListener('focusout', (e) => {
    if (e.target.classList.contains('answer-count') || e.target.classList.contains('answer-text') || e.target.hasAttribute('data-header')) {
        renderTable(e.target.dataset.id);
        drawChart(e.target.dataset.id);
        drawStackedChart(e.target.dataset.id);
        drawPieChart(e.target.dataset.id);
    }
    if (e.target.classList.contains('modal-answer-count') || e.target.classList.contains('modal-answer-text')) {
        const id = e.target.dataset.id;
        renderChartEditModal(id);
        renderTable(id);
        drawChart(id);
        drawStackedChart(id);
        drawPieChart(id);
    }
});

document.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && e.target.classList.contains('editable-cell')) { e.preventDefault(); e.target.blur(); }
});

document.addEventListener('click', (e) => {
    const btn = e.target.closest('.add-row-btn');
    if (!btn) return;
    const id = btn.dataset.id;
    const idx = parseInt(btn.dataset.index);
    const dataObj = window.appData[id];
    if (!dataObj) return;
    const newRow = { answer: '', counts: Object.fromEntries(dataObj.file_keys.map(fk => [fk, 0])), included: true, _total: 0 };
    dataObj.data.splice(idx + 1, 0, newRow);
    renderTable(id);
    setTimeout(() => {
        const cell = document.querySelector(`.answer-text[data-id="${id}"][data-index="${idx + 1}"]`);
        if (cell) { cell.focus(); document.execCommand('selectAll', false, null); }
    }, 30);
});
