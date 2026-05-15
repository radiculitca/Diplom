// ===================== FUZZY / RANGE GROUPING =====================

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
        let merged = '';
        if (from && from.length) {
            const visible = from.slice(0, 3).map(a => `«${_escHtml(a)}»`).join(', ');
            const rest = from.slice(3);
            const moreHtml = rest.length > 0
                ? `, <span class="fuzzy-more-toggle text-primary" style="cursor:pointer">и ещё ${rest.length}</span>`
                  + `<span class="fuzzy-more-items d-none">, ${rest.map(a => `«${_escHtml(a)}»`).join(', ')}</span>`
                : '';
            merged = ` <small class="text-muted">(+ ${visible}${moreHtml})</small>`;
        }
        return `<div class="d-flex justify-content-between align-items-baseline py-1 border-bottom">
            <span>${_escHtml(row.answer)}${merged}</span>
            <span class="badge bg-secondary ms-2 flex-shrink-0">${row._total}</span>
        </div>`;
    }).join('');
    document.getElementById('fuzzyPreviewContent').innerHTML = html || '<span class="text-muted">Нет данных.</span>';
    document.getElementById('fuzzyPreviewArea').classList.remove('d-none');
}

document.getElementById('fuzzyPreviewContent').addEventListener('click', e => {
    const toggle = e.target.closest('.fuzzy-more-toggle');
    if (!toggle) return;
    toggle.classList.add('d-none');
    toggle.nextElementSibling.classList.remove('d-none');
});

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

// Tab switching — clears preview
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

// "Сгруппировать" for fuzzy tab
document.getElementById('fuzzyPreviewBtn').addEventListener('click', () => {
    if (!_fuzzyTargetId) return;
    const result = _computeFuzzyMerge(_fuzzyTargetId);
    if (result) _fuzzyShowPreview(result.data, result.origAnswers);
});

// "Сгруппировать" for ranges tab
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

// "Вручную" — открыть модалку ручной настройки диапазонов
document.getElementById('manualRangesBtn').addEventListener('click', () => {
    if (!_fuzzyTargetId) return;
    const n = parseInt(document.getElementById('rangeCountInput').value, 10);
    if (!(n >= 2)) return;

    const dataObj = window.appData[_fuzzyTargetId];
    if (!dataObj) return;

    const activeRows = dataObj.data.filter(r => r.included !== false);
    const numericVals = activeRows
        .map(r => parseFloat(r.answer))
        .filter(v => !isNaN(v));

    let preRanges;
    if (numericVals.length >= 2) {
        const minVal = Math.min(...numericVals);
        const maxVal = Math.max(...numericVals);
        const step = minVal === maxVal ? 1 : (maxVal - minVal) / n;
        preRanges = Array.from({ length: n }, (_, i) => {
            const loRaw = minVal + i * step;
            const hiRaw = i === n - 1 ? maxVal : minVal + (i + 1) * step;
            const isLast = i === n - 1;
            const uvalsSorted = numericVals.filter(v => v >= loRaw && (isLast ? v <= hiRaw : v < hiRaw)).sort((a, b) => a - b);
            const unique = [...new Set(uvalsSorted)];
            const isSingle = unique.length === 1;
            const lo = unique.length >= 1 ? Math.round(unique[0]) : Math.round(loRaw);
            const hi = unique.length >= 2 ? Math.round(unique[unique.length - 1]) : (isSingle ? lo : Math.round(hiRaw));
            return { lo, hi, isSingle, singleVal: isSingle ? lo : null };
        });
    } else {
        preRanges = Array.from({ length: n }, (_, i) => ({ lo: i, hi: i + 1, isSingle: false, singleVal: null }));
    }

    document.getElementById('manualRangesBody').innerHTML = preRanges.map((r, i) => `
        <div class="d-flex align-items-center gap-2 mb-2 manual-range-row">
            <span class="text-muted small" style="min-width:90px">Диапазон ${i + 1}:</span>
            <input type="number" class="form-control form-control-sm manual-range-lo" value="${r.isSingle ? r.singleVal : r.lo}" style="width:90px">
            <span class="text-muted manual-range-sep"${r.isSingle ? ' style="display:none"' : ''}>—</span>
            <input type="number" class="form-control form-control-sm manual-range-hi" value="${r.hi}" style="width:90px${r.isSingle ? ';display:none' : ''}">
            <div class="manual-range-single-wrap d-flex align-items-center gap-1 ms-2" style="visibility:${r.isSingle ? 'visible' : 'hidden'}">
                <div class="form-check form-switch mb-0">
                    <input class="form-check-input manual-range-single-switch" type="checkbox" role="switch" id="singleSwitch_${i}"${r.isSingle ? ' checked' : ''}>
                    <label class="form-check-label small text-muted" for="singleSwitch_${i}">Одно число</label>
                </div>
            </div>
        </div>`).join('');

    document.querySelectorAll('#manualRangesBody .manual-range-row').forEach(row => {
        const wrap = row.querySelector('.manual-range-single-wrap');
        const sw = row.querySelector('.manual-range-single-switch');
        const sep = row.querySelector('.manual-range-sep');
        const hiEl = row.querySelector('.manual-range-hi');

        row.addEventListener('mouseenter', () => { wrap.style.visibility = 'visible'; });
        row.addEventListener('mouseleave', () => { if (!sw.checked) wrap.style.visibility = 'hidden'; });
        sw.addEventListener('change', () => {
            sep.style.display = sw.checked ? 'none' : '';
            hiEl.style.display = sw.checked ? 'none' : '';
        });
    });

    new bootstrap.Modal(document.getElementById('manualRangesModal')).show();
});

// "Применить" в модалке ручных диапазонов
document.getElementById('applyManualRangesBtn').addEventListener('click', () => {
    if (!_fuzzyTargetId) return;

    const loInputs = [...document.querySelectorAll('#manualRangesBody .manual-range-lo')];
    const hiInputs = [...document.querySelectorAll('#manualRangesBody .manual-range-hi')];
    const singleSwitches = [...document.querySelectorAll('#manualRangesBody .manual-range-single-switch')];

    loInputs.forEach(el => el.classList.remove('is-invalid'));
    hiInputs.forEach(el => el.classList.remove('is-invalid'));

    const customRanges = loInputs.map((el, i) => {
        const lo = parseFloat(el.value);
        const isSingle = singleSwitches[i] && singleSwitches[i].checked;
        const hi = isSingle ? lo : parseFloat(hiInputs[i].value);
        return { lo, hi, isSingle, isLast: i === loInputs.length - 1 };
    });

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
        invalidIdx.forEach(i => {
            loInputs[i].classList.add('is-invalid');
            if (!customRanges[i].isSingle) hiInputs[i].classList.add('is-invalid');
        });
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

// "Применить"
document.getElementById('fuzzyConfirmBtn').addEventListener('click', () => {
    bootstrap.Modal.getInstance(document.getElementById('fuzzyConfirmModal')).hide();
    if (_pendingMergeData && _fuzzyTargetId) _applyMergedData(_fuzzyTargetId, _pendingMergeData);
});

// ИИ-группировка через API
document.getElementById('fuzzyTabPane').addEventListener('click', async e => {
    const btn = e.target.closest('[data-ai-backend]');
    if (!btn || !_fuzzyTargetId) return;

    const dataObj = window.appData[_fuzzyTargetId];
    if (!dataObj || dataObj.data.length === 0) return;

    const answers = dataObj.data.map(r => String(r.answer));
    const questionName = dataObj.question_name || '';

    const progressModal = new bootstrap.Modal(document.getElementById('aiGroupProgressModal'));
    progressModal.show();

    try {
        const resp = await fetch('/ai_group_answers', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ answers, question_name: questionName })
        });
        const result = await resp.json();
        progressModal.hide();

        if (!resp.ok) {
            showToast(result.message || 'Ошибка ИИ-группировки', 'danger');
            return;
        }

        const toCanon = {};
        for (const grp of result.groups) {
            for (const member of grp.members) {
                toCanon[member] = grp.canonical;
            }
        }

        const mergedMap = {};
        const origAnswers = {};
        for (const row of dataObj.data) {
            const canon = toCanon[row.answer] ?? row.answer;
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
        _fuzzyShowPreview(data, origAnswers);
    } catch (err) {
        progressModal.hide();
        showToast('Ошибка соединения с сервером', 'danger');
    }
});

// ===================== COMPUTE FUNCTIONS =====================

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
        if (!isNaN(v) && String(row.answer).trim() !== '') {
            numeric.push({ row, v });
        } else {
            other.push(row);
        }
    }

    if (numeric.length === 0) return null;

    const minVal = Math.min(...numeric.map(x => x.v));
    const maxVal = Math.max(...numeric.map(x => x.v));
    const step = minVal === maxVal ? 1 : (maxVal - minVal) / numRanges;

    const buckets = Array.from({ length: numRanges }, (_, i) => {
        const lo = minVal + i * step;
        const hi = i === numRanges - 1 ? maxVal : minVal + (i + 1) * step;
        return {
            lo, hi, isLast: i === numRanges - 1,
            counts: Object.fromEntries(dataObj.file_keys.map(fk => [fk, 0])),
            included: true,
            _total: 0,
            _uniqueVals: new Set()
        };
    });

    for (const { row, v } of numeric) {
        const bucket = buckets.find(b => b.isLast ? v <= b.hi : v >= b.lo && v < b.hi) || buckets[buckets.length - 1];
        dataObj.file_keys.forEach(fk => { bucket.counts[fk] += row.counts[fk] || 0; });
        bucket._total += row._total;
        bucket._uniqueVals.add(v);
    }

    const data = buckets.map(({ lo, hi, counts, included, _total, _uniqueVals }) => {
        const uvals = [..._uniqueVals].sort((a, b) => a - b);
        let answer;
        if (uvals.length === 0) answer = `${Math.round(lo)} – ${Math.round(hi)}`;
        else if (uvals.length === 1) answer = String(Math.round(uvals[0]));
        else answer = `${Math.round(uvals[0])} – ${Math.round(uvals[uvals.length - 1])}`;
        return { answer, counts, included, _total };
    });

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
        if (!isNaN(v) && String(row.answer).trim() !== '') {
            numeric.push({ row, v });
        } else {
            other.push(row);
        }
    }
    if (numeric.length === 0) return null;

    const buckets = customRanges.map(r => ({
        lo: r.lo, hi: r.hi, isLast: r.isLast,
        counts: Object.fromEntries(dataObj.file_keys.map(fk => [fk, 0])),
        included: true,
        _total: 0,
        _uniqueVals: new Set()
    }));

    const unmatched = [];
    for (const { row, v } of numeric) {
        const bucket = buckets.find(b => v >= b.lo && v <= b.hi);
        if (bucket) {
            dataObj.file_keys.forEach(fk => { bucket.counts[fk] += row.counts[fk] || 0; });
            bucket._total += row._total;
            bucket._uniqueVals.add(v);
        } else {
            unmatched.push(row);
        }
    }

    const data = buckets.map(({ lo, hi, counts, included, _total, _uniqueVals }) => {
        const uvals = [..._uniqueVals].sort((a, b) => a - b);
        let answer;
        if (uvals.length === 0) answer = `${lo} – ${hi}`;
        else if (uvals.length === 1) answer = String(Math.round(uvals[0]));
        else answer = `${Math.round(uvals[0])} – ${Math.round(uvals[uvals.length - 1])}`;
        return { answer, counts, included, _total };
    });

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
