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
}

function initTooltips() {
    var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    tooltipTriggerList.map(function (t) { return new bootstrap.Tooltip(t); });
}

// Step navigation buttons
document.getElementById('toStep4Btn').addEventListener('click', () => { goToStep(3); renderStep4(); });
document.getElementById('toStep5Btn').addEventListener('click', () => { renderLegendSettings(); goToStep(4); });
document.getElementById('toStep6Btn').addEventListener('click', () => goToStep(5));

function _resetFromStep(newStep) {
    // Шаг 4 (визуализация): уничтожить графики, очистить данные анализа
    if (newStep < 4) {
        Object.values(window.charts || {}).forEach(c => { try { c.destroy(); } catch (e) {} });
        window.charts = {};
        window.appData = {};
        window.chartsData = {};
        window.stackedChartsData = {};
        window.pieChartsData = {};
        window.renderedTabs = {};
        document.getElementById('reportContent').innerHTML = '';
    }

    // Шаг 3 (структура отчёта): очистить разделы
    if (newStep < 3) {
        window.reportSections = [];
        document.getElementById('sectionsList').innerHTML = '<p class="text-muted small text-center mt-3 px-2 mb-0"></p>';
        document.getElementById('availableQuestionsList').innerHTML = '<p class="text-muted small text-center mt-2 mb-0">Выберите вопросы на шаге 3</p>';
        const selAllQ = document.getElementById('selectAllAvailableQ');
        if (selAllQ) selAllQ.checked = false;
    }

    // Шаг 2 (выбор вопросов): очистить маппинг и списки
    if (newStep < 2) {
        window.questionMapping = {};
        window.questionSourceFile = {};
        document.getElementById('sortableQuestionsList').innerHTML = '<div id="emptySortablePlaceholder"></div>';
        document.getElementById('allQuestionsList').innerHTML = '';
        document.getElementById('selectAllQuestions').checked = false;
        document.getElementById('questionsSearchInput').value = '';
        const fileSelect = document.getElementById('fileSelectStep3');
        if (fileSelect) fileSelect.innerHTML = '';
        document.getElementById('fileSelectContainer').style.display = 'none';
        updateQuestionsBtn();
    }

    // Шаг 1 (листы): очистить обработанные файлы и список листов
    if (newStep < 1) {
        window.processedFiles = [];
        document.getElementById('sheetCheckboxesContainer').innerHTML = '';
        updateSheetBtn();
    }
}

// Back button (global delegation)
document.addEventListener('click', e => {
    if (e.target.closest('.wizard-back-btn')) {
        if (currentWizardStep > 0) {
            const newStep = currentWizardStep - 1;
            _resetFromStep(newStep);
            goToStep(newStep);
            if (newStep === 3) renderStep4();
        }
    }
});
