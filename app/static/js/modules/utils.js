// ===================== GLOBAL STATE =====================
window.uploadedFiles = [];
window.processedFiles = [];
window.questionMapping = {};
window.questionSourceFile = {};
window.charts = {};
window.appData = {};
window.questionMerges = {};

const defaultColors = ['#FF0000', '#4472C4', '#70AD47', '#FFC000', '#ED7D31', '#A5A5A5', '#5B9BD5', '#C00000', '#00B050', '#7030A0'];

// ===================== UTILITY FUNCTIONS =====================

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

function _escHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function _escAttr(s) {
    return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function randomColor() {
    return '#' + Math.floor(Math.random() * 0xffffff).toString(16).padStart(6, '0');
}
