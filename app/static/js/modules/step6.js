// ===================== ШАГ 6: ГЕНЕРАЦИЯ ИИ-ОТЧЁТА =====================

async function _doExport() {
    const sections = window.reportSections || [];
    if (!window.appData || Object.keys(window.appData).length === 0) {
        showToast('Нет данных анализа. Сначала запустите анализ на шаге 5.', 'danger');
        return;
    }

    // Индекс question_name → id в appData для быстрого поиска
    const nameToId = {};
    for (const id of Object.keys(window.appData)) {
        const qName = window.appData[id].question_name;
        if (qName) nameToId[qName] = id;
    }

    const questions = [];
    let tableNum = 1;

    const processedQNames = new Set();

    // Из разделов
    for (const sec of sections) {
        for (const qEntry of (sec.questions || [])) {
            const qName = qEntry.qName;
            const id = nameToId[qName];
            if (!id || processedQNames.has(qName)) continue;
            processedQNames.add(qName);

            const dataObj = window.appData[id];
            const activeRows = dataObj.data.filter(r => r.included !== false);
            if (activeRows.length === 0) continue;

            const fileTotals = {};
            dataObj.file_keys.forEach(fk => {
                fileTotals[fk] = dataObj.data.reduce((sum, r) => sum + (r.counts[fk] || 0), 0);
            });

            const opts = dataObj.options || {};
            const includeInReport = dataObj.includeInReport !== false;
            let vizTab = null;
            if (includeInReport) {
                const activeBtn = document.querySelector(`#tabs_${id} .viz-tab-btn.active`);
                vizTab = activeBtn ? activeBtn.dataset.tab : 'table';
            }

            questions.push({
                table_num: tableNum++,
                question_name: dataObj.question_name,
                h1: dataObj.headers.h1 || 'Ответ',
                h2: dataObj.headers.h2 || 'Кол-во ответивших',
                h3: dataObj.headers.h3 || '% от числа ответивших',
                file_keys: dataObj.file_keys,
                file_labels: dataObj.file_labels,
                rows: activeRows.map(r => ({
                    answer: String(r.answer),
                    counts: Object.fromEntries(dataObj.file_keys.map(fk => [fk, r.counts[fk] || 0]))
                })),
                file_totals: fileTotals,
                show_total: opts.showTotal !== false,
                section: { name: sec.name, description: sec.description || '', color: sec.color || '' },
                viz_tab: vizTab,
                chart_direction: opts.chartDirection || 'y',
                show_legend: opts.showLegend !== false,
                hidden_col: opts.hiddenCol || 'none',
            });
        }
    }

    // Нераспределённые вопросы (без раздела)
    for (const [id, dataObj] of Object.entries(window.appData)) {
        const qName = dataObj.question_name;
        if (processedQNames.has(qName)) continue;
        processedQNames.add(qName);

        const activeRows = dataObj.data.filter(r => r.included !== false);
        if (activeRows.length === 0) continue;

        const fileTotals = {};
        dataObj.file_keys.forEach(fk => {
            fileTotals[fk] = dataObj.data.reduce((sum, r) => sum + (r.counts[fk] || 0), 0);
        });

        const opts = dataObj.options || {};
        const includeInReport = dataObj.includeInReport !== false;
        let vizTab = null;
        if (includeInReport) {
            const activeBtn = document.querySelector(`#tabs_${id} .viz-tab-btn.active`);
            vizTab = activeBtn ? activeBtn.dataset.tab : 'table';
        }

        questions.push({
            table_num: tableNum++,
            question_name: dataObj.question_name,
            h1: dataObj.headers.h1 || 'Ответ',
            h2: dataObj.headers.h2 || 'Кол-во ответивших',
            h3: dataObj.headers.h3 || '% от числа ответивших',
            file_keys: dataObj.file_keys,
            file_labels: dataObj.file_labels,
            rows: activeRows.map(r => ({
                answer: String(r.answer),
                counts: Object.fromEntries(dataObj.file_keys.map(fk => [fk, r.counts[fk] || 0]))
            })),
            file_totals: fileTotals,
            show_total: opts.showTotal !== false,
            section: null,  // без раздела
            viz_tab: vizTab,
            chart_direction: opts.chartDirection || 'y',
            show_legend: opts.showLegend !== false,
            hidden_col: opts.hiddenCol || 'none',
        });
    }

    if (questions.length === 0) {
        showToast('Ни один вопрос из разделов не попал в анализ. Запустите анализ на шаге 5.', 'danger');
        return;
    }

    const btn = document.getElementById('downloadApiBtn');
    const origHtml = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Генерация...';

    const progressContainer = document.getElementById('exportProgressContainer');
    const progressBar = document.getElementById('exportProgressBar');
    const progressLabel = document.getElementById('exportProgressLabel');
    if (progressContainer) progressContainer.classList.remove('d-none');
    if (progressBar) { progressBar.style.width = '0%'; progressBar.textContent = ''; }
    if (progressLabel) progressLabel.textContent = '';

    const _exportStartTime = Date.now();
    const _exportTimerInterval = setInterval(() => {
        const elapsed = Math.floor((Date.now() - _exportStartTime) / 1000);
        const m = Math.floor(elapsed / 60).toString().padStart(2, '0');
        const s = (elapsed % 60).toString().padStart(2, '0');
        const timerEl = document.getElementById('exportTimerLabel');
        if (timerEl) timerEl.textContent = `${m}:${s}`;
    }, 1000);

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
            buffer = lines.pop();

            for (const line of lines) {
                if (!line.startsWith('data: ')) continue;
                const msg = JSON.parse(line.slice(6));

                if (msg.type === 'progress') {
                    const pct = Math.round((msg.current / msg.total) * 100);
                    if (progressBar) { progressBar.style.width = pct + '%'; progressBar.textContent = `${msg.current}/${msg.total}`; }
                    if (progressLabel) progressLabel.textContent = `Вопрос ${msg.current} из ${msg.total}: ${msg.label}`;
                }

                if (msg.type === 'done') {
                    const fileBytes = Uint8Array.from(atob(msg.file), c => c.charCodeAt(0));
                    const fileBlob = new Blob([fileBytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
                    const fileUrl = URL.createObjectURL(fileBlob);
                    const a = document.createElement('a');
                    a.href = fileUrl; a.download = msg.filename || 'report_analysis.docx';
                    document.body.appendChild(a); a.click();
                    document.body.removeChild(a); URL.revokeObjectURL(fileUrl);

                    showToast('Готово: отчёт скачан', 'success');
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
        clearInterval(_exportTimerInterval);
        if (progressContainer) setTimeout(() => progressContainer.classList.add('d-none'), 2000);
    }
}

document.getElementById('downloadApiBtn').addEventListener('click', () => {
    _doExport();
});

// ===================== TOOLTIP INIT =====================
document.querySelectorAll('[data-bs-toggle="tooltip"]').forEach(el => new bootstrap.Tooltip(el, { trigger: 'hover' }));

initTooltips();
