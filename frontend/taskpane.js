/**
 * Excel Commander - Task Pane JavaScript
 * Handles Office.js interactions and API calls.
 */

// ============ Configuration ============
const API_BASE_URL = 'https://excel-commander.onrender.com';

// ============ State ============
let isOfficeReady = false;

// ============ UI Elements ============
const elements = {
    statusBadge: null,
    commandInput: null,
    runBtn: null,
    outputPanel: null,
    loadingOverlay: null
};

// ============ Initialization ============
Office.onReady((info) => {
    // Cache DOM elements
    elements.statusBadge = document.getElementById('status-badge');
    elements.commandInput = document.getElementById('command-input');
    elements.runBtn = document.getElementById('run-btn');
    elements.outputPanel = document.getElementById('output-panel');
    elements.loadingOverlay = document.getElementById('loading-overlay');

    if (info.host === Office.HostType.Excel) {
        isOfficeReady = true;
        updateStatus('online', '✅ Bağlandı');
        console.log('Excel Commander: Office.js ready');
    } else {
        updateStatus('offline', '⚠️ Excel dışı mod');
    }

    // Event Listeners
    elements.runBtn.onclick = handleCommand;
    elements.commandInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && e.ctrlKey) {
            handleCommand();
        }
    });

    // Check API health
    checkApiHealth();
});

// ============ API Functions ============

async function checkApiHealth() {
    try {
        const response = await fetch(`${API_BASE_URL}/`);
        if (response.ok) {
            console.log('API connection successful');
        }
    } catch (error) {
        console.warn('API not reachable:', error);
        updateStatus('offline', '⚠️ API Bağlantısı Yok');
    }
}

async function apiCall(endpoint, method = 'GET', body = null) {
    const options = {
        method,
        headers: { 'Content-Type': 'application/json' }
    };
    if (body) {
        options.body = JSON.stringify(body);
    }
    const response = await fetch(`${API_BASE_URL}${endpoint}`, options);
    return response.json();
}

// ============ Command Handler ============

async function handleCommand() {
    const command = elements.commandInput.value.trim();
    if (!command) return;

    showLoading(true);

    try {
        const result = await apiCall('/api/formula/generate', 'POST', {
            description: command,
            language: 'tr'
        });

        if (result.success) {
            showOutput('formula', result.formula, result.explanation);

            // Write to Excel if ready
            if (isOfficeReady) {
                await writeToActiveCell(result.formula);
            }
        } else {
            showOutput('error', null, result.error || 'Bir hata oluştu.');
        }
    } catch (error) {
        showOutput('error', null, `Bağlantı hatası: ${error.message}`);
    }

    showLoading(false);
}

// ============ Action Handlers ============

const actions = {
    async generateFormula() {
        const desc = prompt('Formülü açıkla (Örn: A sütunundaki sayıları topla):');
        if (!desc) return;

        elements.commandInput.value = desc;
        await handleCommand();
    },

    async explainFormula() {
        if (!isOfficeReady) {
            showOutput('error', null, 'Excel bağlantısı gerekli.');
            return;
        }

        showLoading(true);

        try {
            const formula = await getActiveCell();

            if (!formula || !formula.startsWith('=')) {
                showOutput('error', null, 'Lütfen bir formül içeren hücre seçin.');
                showLoading(false);
                return;
            }

            const result = await apiCall('/api/formula/explain', 'POST', {
                formula: formula,
                language: 'tr'
            });

            if (result.success) {
                showOutput('explanation', formula, result.explanation);
            } else {
                showOutput('error', null, result.error);
            }
        } catch (error) {
            showOutput('error', null, error.message);
        }

        showLoading(false);
    },

    async cleanData() {
        if (!isOfficeReady) {
            showOutput('error', null, 'Excel bağlantısı gerekli.');
            return;
        }

        showLoading(true);

        try {
            const data = await getSelectedRangeData();

            if (!data || data.length === 0) {
                showOutput('error', null, 'Lütfen temizlenecek veri aralığını seçin.');
                showLoading(false);
                return;
            }

            const result = await apiCall('/api/formula/clean', 'POST', {
                data: data
            });

            if (result.success) {
                await writeToSelectedRange(result.cleaned_data);
                showOutput('success', null, `✅ Veri temizlendi! ${result.changes_made?.length || 0} değişiklik yapıldı.`);
            } else {
                showOutput('error', null, result.error);
            }
        } catch (error) {
            showOutput('error', null, error.message);
        }

        showLoading(false);
    },

    async generateSlide() {
        if (!isOfficeReady) {
            showOutput('error', null, 'Excel bağlantısı gerekli.');
            return;
        }

        showLoading(true);

        try {
            const data = await getSelectedRangeData();

            if (!data || data.length < 2) {
                showOutput('error', null, 'Lütfen en az başlık + 1 satır veri seçin.');
                showLoading(false);
                return;
            }

            const title = prompt('Sunum Başlığı:', 'Excel Analiz Raporu');

            const result = await apiCall('/api/presentation/generate', 'POST', {
                data: data,
                title: title || 'Analiz Raporu',
                insights_count: 3,
                include_chart: true,
                chart_type: 'chart_bar'
            });

            if (result.success) {
                // Create download link
                const downloadUrl = `${API_BASE_URL}${result.file_url}`;

                let insightsHtml = result.insights?.map(i => `<li>${i}</li>`).join('') || '';

                elements.outputPanel.innerHTML = `
                    <div class="output-success">
                        <strong>✅ Sunum Hazır!</strong>
                    </div>
                    <p style="margin: 8px 0;">
                        <a href="${downloadUrl}" download style="color: var(--color-secondary); font-weight: 600;">
                            📥 Sunumu İndir (PPTX)
                        </a>
                    </p>
                    ${insightsHtml ? `<p style="font-size: 12px; color: #666;"><strong>Bulunan İçgörüler:</strong></p><ul style="font-size: 12px; margin-left: 16px;">${insightsHtml}</ul>` : ''}
                `;
            } else {
                showOutput('error', null, result.error);
            }
        } catch (error) {
            showOutput('error', null, error.message);
        }

        showLoading(false);
    },

    showHelp() {
        elements.outputPanel.innerHTML = `
            <strong>🆘 Yardım</strong>
            <ul style="margin: 8px 0 0 16px; font-size: 13px;">
                <li><strong>Formül Yaz:</strong> İstediğinizi yazın, AI formülü üretsin.</li>
                <li><strong>Formül Açıkla:</strong> Bir formül seçin, AI açıklasın.</li>
                <li><strong>Veri Temizle:</strong> Veri aralığı seçin, isimler düzeltilsin.</li>
                <li><strong>Sunum Yap:</strong> Veri seçin, PPT otomatik oluşsun!</li>
            </ul>
            <p style="margin-top: 12px; font-size: 11px; color: #999;">Kısayol: Ctrl+Enter</p>
        `;
    }
};

// ============ Excel Helpers ============

async function getActiveCell() {
    return Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('values');
        await context.sync();
        return range.values[0][0]?.toString() || '';
    });
}

async function writeToActiveCell(value) {
    return Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.values = [[value]];
        await context.sync();
    });
}

async function getSelectedRangeData() {
    return Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load('values');
        await context.sync();
        return range.values;
    });
}

async function writeToSelectedRange(data) {
    return Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.values = data;
        await context.sync();
    });
}

// ============ UI Helpers ============

function updateStatus(type, text) {
    elements.statusBadge.textContent = text;
    elements.statusBadge.className = `badge badge-${type === 'online' ? 'online' : 'offline'}`;
}

function showLoading(show) {
    elements.loadingOverlay.classList.toggle('hidden', !show);
}

function showOutput(type, primary, secondary) {
    let html = '';

    switch (type) {
        case 'formula':
            html = `
                <div class="output-formula">${escapeHtml(primary)}</div>
                ${secondary ? `<div class="output-explanation">${escapeHtml(secondary)}</div>` : ''}
            `;
            break;
        case 'explanation':
            html = `
                <div class="output-formula">${escapeHtml(primary)}</div>
                <div class="output-explanation">${escapeHtml(secondary)}</div>
            `;
            break;
        case 'success':
            html = `<div class="output-success">${secondary}</div>`;
            break;
        case 'error':
            html = `<div class="output-error">❌ ${escapeHtml(secondary)}</div>`;
            break;
        default:
            html = secondary;
    }

    elements.outputPanel.innerHTML = html;
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
