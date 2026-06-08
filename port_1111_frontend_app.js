// ============================================
// Web Monitoring — Local Data Visualizer + QoS
// ============================================
'use strict';

const API_BASE = '';

// ============================================
// DOM Elements
// ============================================
const elements = {
    // Upload
    fileInput: document.getElementById('fileInput'),
    uploadLabel: document.getElementById('uploadLabel'),
    clearBtn: document.getElementById('clearBtn'),
    fileInfo: document.getElementById('fileInfo'),
    fileName: document.getElementById('fileName'),
    fileRows: document.getElementById('fileRows'),
    dropOverlay: document.getElementById('dropOverlay'),
    location: document.getElementById('location'),

    // Cards
    cards: document.querySelectorAll('.skeleton-card'),
    pressureNeedle: document.getElementById('pressureNeedle'),
    thermoFill: document.getElementById('thermoFill'),
    humidityDrop: document.getElementById('humidityDrop'),
    compassArrow: document.getElementById('compassArrow'),

    // Table
    dataTableBody: document.getElementById('dataTableBody'),
    dataRange: document.getElementById('dataRange'),
    exportBtn: document.getElementById('exportBtn'),
    pagination: document.getElementById('pagination'),
    prevBtn: document.getElementById('prevBtn'),
    nextBtn: document.getElementById('nextBtn'),
    pageInfo: document.getElementById('pageInfo'),
    pageSizeSelect: document.getElementById('pageSizeSelect'),

    // Tabs
    tabData: document.getElementById('tabData'),
    tabQos: document.getElementById('tabQos'),
    tabContentData: document.getElementById('tabContentData'),
    tabContentQos: document.getElementById('tabContentQos'),

    // QoS
    qosEmpty: document.getElementById('qosEmpty'),
    qosContent: document.getElementById('qosContent'),
    intervalSelect: document.getElementById('intervalSelect'),
    recalcBtn: document.getElementById('recalcBtn'),

    // QoS Latency
    latencyNoData: document.getElementById('latencyNoData'),
    latencyResult: document.getElementById('latencyResult'),
    latencyValue: document.getElementById('latencyValue'),
    latencyBadge: document.getElementById('latencyBadge'),
    latencyFormula: document.getElementById('latencyFormula'),
    latencyMin: document.getElementById('latencyMin'),
    latencyMax: document.getElementById('latencyMax'),
    latencyPackets: document.getElementById('latencyPackets'),

    // QoS Packet Loss
    packetLossNoData: document.getElementById('packetLossNoData'),
    packetLossResult: document.getElementById('packetLossResult'),
    packetLossValue: document.getElementById('packetLossValue'),
    packetLossBadge: document.getElementById('packetLossBadge'),
    packetLossFormula: document.getElementById('packetLossFormula'),
    pSend: document.getElementById('pSend'),
    pReceive: document.getElementById('pReceive'),
    pLost: document.getElementById('pLost')
};

// ============================================
// State
// ============================================
let currentPage = 1;
let pageSize = parseInt(elements.pageSizeSelect?.value) || 20;
let hasData = false;

// ============================================
// Tab Navigation
// ============================================
function initTabs() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const tab = btn.dataset.tab;

            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

            btn.classList.add('active');
            if (tab === 'data') {
                elements.tabContentData.classList.add('active');
            } else if (tab === 'qos') {
                elements.tabContentQos.classList.add('active');
            }
        });
    });
}

// ============================================
// File Upload
// ============================================
async function uploadFile(file) {
    const formData = new FormData();
    formData.append('file', file);

    // Pass interval setting
    const interval = elements.intervalSelect?.value || '300';
    formData.append('interval', interval);

    elements.uploadLabel.classList.add('uploading');

    try {
        const response = await fetch(`${API_BASE}/api/upload`, {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (!result.success) {
            alert('Error: ' + result.message);
            return;
        }

        // Update file info
        elements.fileInfo.classList.add('has-file');
        elements.fileName.textContent = result.filename;
        elements.fileRows.textContent = `${result.totalRows} baris`;
        elements.clearBtn.style.display = 'flex';
        elements.exportBtn.disabled = false;
        hasData = true;

        // Update location
        if (result.location) {
            elements.location.textContent = `📍 ${result.location}`;
        } else {
            elements.location.textContent = '📍 Lokasi tidak terdeteksi';
        }

        // Update cards
        if (result.summary) {
            updateCards(result.summary);
            if (result.summary.firstTime && result.summary.lastTime) {
                elements.dataRange.textContent = `${result.summary.firstTime} — ${result.summary.lastTime}`;
            }
        }

        // Fetch paginated data
        currentPage = 1;
        await fetchData();

        // Update QoS
        updateQoS(result.qos);

    } catch (error) {
        console.error('Upload error:', error);
        alert('Gagal upload file. Pastikan server berjalan.');
    } finally {
        elements.uploadLabel.classList.remove('uploading');
    }
}

// ============================================
// Fetch paginated table data
// ============================================
async function fetchData() {
    try {
        const response = await fetch(`${API_BASE}/api/data?page=${currentPage}&limit=${pageSize}`);
        const result = await response.json();

        if (result.success) {
            updateTable(result.data, result.pagination);
        }
    } catch (error) {
        console.error('Fetch data error:', error);
    }
}

// ============================================
// Update sensor cards — show null as "--"
// ============================================
function updateCards(summary) {
    elements.cards.forEach(card => card.classList.add('loaded'));

    const fields = [
        { key: 'airpressure', decimals: 2, cardId: 'cardPressure' },
        { key: 'airtemperature', decimals: 1, cardId: 'cardTemperature' },
        { key: 'airhumidity', decimals: 1, cardId: 'cardHumidity' },
        { key: 'windspeed', decimals: 2, cardId: 'cardWindspeed' },
        { key: 'winddirection', decimals: 0, cardId: 'cardWinddirection' }
    ];

    fields.forEach(({ key, decimals, cardId }) => {
        const el = document.getElementById(key);
        const cardEl = document.getElementById(cardId);
        const value = summary[key];

        if (value === null || value === undefined) {
            if (el) el.textContent = '--';
            if (cardEl) cardEl.classList.remove('warning');
        } else {
            const newVal = parseFloat(value).toFixed(decimals);
            if (el && el.textContent !== newVal) {
                el.textContent = newVal;
                triggerAnimation(el, 'value-updated', 400);
                if (cardEl) triggerAnimation(cardEl, 'data-updated', 300);
            }
        }

        // Threshold warnings (only when value exists)
        if (cardEl && value !== null && value !== undefined) {
            cardEl.classList.remove('warning');
            if (key === 'airtemperature' && value >= 32) cardEl.classList.add('warning');
            if (key === 'airhumidity' && (value < 40 || value > 70)) cardEl.classList.add('warning');
            if (key === 'windspeed' && value >= 5) cardEl.classList.add('warning');
        }
    });

    // Update SVG animations
    updateAnimations(summary);

    // Wind direction name
    if (summary.winddirection !== null && summary.winddirection !== undefined) {
        const dirEl = document.getElementById('directionName');
        if (dirEl) dirEl.textContent = getDirectionName(summary.winddirection);
    }
}

// ============================================
// Update SVG animations
// ============================================
function updateAnimations(data) {
    if (elements.pressureNeedle && data.airpressure != null) {
        const angle = ((data.airpressure - 980) / 40) * 180 - 90;
        elements.pressureNeedle.style.transform = `rotate(${angle}deg)`;
    }
    if (elements.thermoFill && data.airtemperature != null) {
        const pct = Math.max(0, Math.min(100, ((data.airtemperature - 15) / 30) * 100));
        elements.thermoFill.style.height = `${pct}%`;
    }
    if (elements.humidityDrop && data.airhumidity != null) {
        const inv = 100 - data.airhumidity;
        elements.humidityDrop.style.clipPath = `inset(${inv}% 0 0 0)`;
    }
    if (elements.compassArrow && data.winddirection != null) {
        elements.compassArrow.style.transform = `rotate(${data.winddirection}deg)`;
    }
}

// ============================================
// Update data table — show null/empty as "--"
// ============================================
function updateTable(data, pagination) {
    const tbody = elements.dataTableBody;
    if (!data || data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="empty-state">
            <div class="empty-icon">📂</div>
            <div class="empty-text">Upload file Excel atau CSV untuk menampilkan data</div>
            <div class="empty-hint">Drag & drop file ke halaman ini, atau klik tombol "Upload File"</div>
        </td></tr>`;
        elements.pagination.style.display = 'none';
        return;
    }

    const startNum = (pagination.page - 1) * pagination.limit;
    tbody.innerHTML = data.map((row, i) => {
        const fmtVal = (val, decimals) => {
            if (val === null || val === undefined || val === '') return `<td class="empty-cell">--</td>`;
            return `<td>${parseFloat(val).toFixed(decimals)}</td>`;
        };

        return `<tr>
            <td class="row-num">${startNum + i + 1}</td>
            <td>${row.tlocal || '<span class="empty-cell">--</span>'}</td>
            ${fmtVal(row.airpressure, 2)}
            ${fmtVal(row.airtemperature, 1)}
            ${fmtVal(row.airhumidity, 1)}
            ${fmtVal(row.windspeed, 2)}
            ${fmtVal(row.winddirection, 0)}
        </tr>`;
    }).join('');

    // Pagination
    elements.pagination.style.display = 'flex';
    elements.pageInfo.textContent = `Halaman ${pagination.page} dari ${pagination.totalPages} (${pagination.totalRows} data)`;
    elements.prevBtn.disabled = !pagination.hasPrev;
    elements.nextBtn.disabled = !pagination.hasNext;
}

// ============================================
// QoS Display
// ============================================
function updateQoS(qos) {
    if (!qos) {
        elements.qosEmpty.style.display = 'block';
        elements.qosContent.style.display = 'none';
        return;
    }

    elements.qosEmpty.style.display = 'none';
    elements.qosContent.style.display = 'block';

    // --- LATENCY ---
    if (qos.latency) {
        elements.latencyNoData.style.display = 'none';
        elements.latencyResult.style.display = 'flex';

        elements.latencyValue.textContent = qos.latency.avgMs.toFixed(2);
        elements.latencyBadge.textContent = qos.latency.category;
        elements.latencyBadge.style.background = qos.latency.color + '22';
        elements.latencyBadge.style.color = qos.latency.color;
        elements.latencyBadge.style.border = `1px solid ${qos.latency.color}`;
        elements.latencyFormula.textContent = qos.latency.formula;
        elements.latencyMin.textContent = qos.latency.minMs + ' ms';
        elements.latencyMax.textContent = qos.latency.maxMs + ' ms';
        elements.latencyPackets.textContent = qos.latency.totalPackets;

        // Highlight active tier in reference table
        highlightTier('qosLatencyCard', qos.latency.index);
    } else {
        elements.latencyNoData.style.display = 'block';
        elements.latencyResult.style.display = 'none';
        clearTierHighlight('qosLatencyCard');
    }

    // --- PACKET LOSS ---
    if (qos.packetLoss) {
        elements.packetLossNoData.style.display = 'none';
        elements.packetLossResult.style.display = 'flex';

        elements.packetLossValue.textContent = qos.packetLoss.percentage.toFixed(2);
        elements.packetLossBadge.textContent = `${qos.packetLoss.category} — ${qos.packetLoss.description}`;
        elements.packetLossBadge.style.background = qos.packetLoss.color + '22';
        elements.packetLossBadge.style.color = qos.packetLoss.color;
        elements.packetLossBadge.style.border = `1px solid ${qos.packetLoss.color}`;
        elements.packetLossFormula.textContent = qos.packetLoss.formula;
        elements.pSend.textContent = qos.packetLoss.pSend;
        elements.pReceive.textContent = qos.packetLoss.pReceive;
        elements.pLost.textContent = qos.packetLoss.lost;

        highlightTier('qosPacketLossCard', qos.packetLoss.index);
    } else {
        elements.packetLossNoData.style.display = 'block';
        elements.packetLossResult.style.display = 'none';
        clearTierHighlight('qosPacketLossCard');
    }
}

function highlightTier(cardId, tierIndex) {
    const card = document.getElementById(cardId);
    if (!card) return;
    card.querySelectorAll('.ref-row').forEach(row => {
        row.classList.remove('active-tier');
        if (row.dataset.tier === String(tierIndex)) {
            row.classList.add('active-tier');
        }
    });
}

function clearTierHighlight(cardId) {
    const card = document.getElementById(cardId);
    if (!card) return;
    card.querySelectorAll('.ref-row').forEach(row => row.classList.remove('active-tier'));
}

// ============================================
// QoS Recalculate
// ============================================
async function recalculateQoS() {
    const interval = elements.intervalSelect.value;

    try {
        const response = await fetch(`${API_BASE}/api/qos/recalculate`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ interval: parseInt(interval) })
        });

        const result = await response.json();
        if (result.success) {
            updateQoS(result.qos);
        } else {
            alert(result.message || 'Gagal menghitung ulang QoS');
        }
    } catch (error) {
        console.error('QoS recalculate error:', error);
        alert('Error: Pastikan server berjalan');
    }
}

// ============================================
// Clear data
// ============================================
async function clearData() {
    try {
        await fetch(`${API_BASE}/api/data`, { method: 'DELETE' });

        // Reset UI
        elements.fileInfo.classList.remove('has-file');
        elements.fileName.textContent = 'Belum ada file';
        elements.fileRows.textContent = '';
        elements.clearBtn.style.display = 'none';
        elements.exportBtn.disabled = true;
        elements.location.textContent = '📍 Belum ada data';
        elements.dataRange.textContent = '';
        hasData = false;

        // Reset cards
        elements.cards.forEach(card => {
            card.classList.remove('loaded', 'warning');
        });
        ['airpressure', 'airtemperature', 'airhumidity', 'windspeed', 'winddirection'].forEach(key => {
            const el = document.getElementById(key);
            if (el) el.textContent = '--';
        });
        const dirEl = document.getElementById('directionName');
        if (dirEl) dirEl.textContent = '--';

        // Reset table
        elements.dataTableBody.innerHTML = `<tr><td colspan="7" class="empty-state">
            <div class="empty-icon">📂</div>
            <div class="empty-text">Upload file Excel atau CSV untuk menampilkan data</div>
            <div class="empty-hint">Drag & drop file ke halaman ini, atau klik tombol "Upload File"</div>
        </td></tr>`;
        elements.pagination.style.display = 'none';

        // Reset QoS
        elements.qosEmpty.style.display = 'block';
        elements.qosContent.style.display = 'none';
        clearTierHighlight('qosLatencyCard');
        clearTierHighlight('qosPacketLossCard');

    } catch (error) {
        console.error('Clear error:', error);
    }
}

// ============================================
// Export CSV
// ============================================
function exportCSV() {
    window.open(`${API_BASE}/api/export`, '_blank');
}

// ============================================
// Drag & Drop
// ============================================
function setupDragDrop() {
    let dragCounter = 0;

    document.addEventListener('dragenter', (e) => {
        e.preventDefault();
        dragCounter++;
        elements.dropOverlay.classList.add('active');
    });

    document.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dragCounter--;
        if (dragCounter <= 0) {
            dragCounter = 0;
            elements.dropOverlay.classList.remove('active');
        }
    });

    document.addEventListener('dragover', (e) => e.preventDefault());

    document.addEventListener('drop', (e) => {
        e.preventDefault();
        dragCounter = 0;
        elements.dropOverlay.classList.remove('active');

        const file = e.dataTransfer.files[0];
        if (file) uploadFile(file);
    });
}

// ============================================
// Helpers
// ============================================
function triggerAnimation(el, cls, duration) {
    el.classList.remove(cls);
    void el.offsetWidth; // force reflow
    el.classList.add(cls);
    setTimeout(() => el.classList.remove(cls), duration);
}

function getDirectionName(deg) {
    if (deg == null) return '--';
    const dirs = ['U', 'TL', 'T', 'TG', 'S', 'BD', 'B', 'BL'];
    const idx = Math.round(deg / 45) % 8;
    return dirs[idx];
}

// ============================================
// Check for existing data on page load
// ============================================
async function checkExistingData() {
    try {
        const response = await fetch(`${API_BASE}/api/summary`);
        const result = await response.json();

        if (result.success && result.filename && result.summary) {
            // Data exists from a previous upload
            elements.fileInfo.classList.add('has-file');
            elements.fileName.textContent = result.filename;
            elements.fileRows.textContent = `${result.summary.totalRows} baris`;
            elements.clearBtn.style.display = 'flex';
            elements.exportBtn.disabled = false;
            hasData = true;

            if (result.location) {
                elements.location.textContent = `📍 ${result.location}`;
            }

            updateCards(result.summary);

            if (result.summary.firstTime && result.summary.lastTime) {
                elements.dataRange.textContent = `${result.summary.firstTime} — ${result.summary.lastTime}`;
            }

            await fetchData();

            // Also load QoS
            const qosResponse = await fetch(`${API_BASE}/api/qos`);
            const qosResult = await qosResponse.json();
            if (qosResult.success) {
                updateQoS(qosResult.qos);
            }
        }
    } catch (error) {
        console.log('No existing data or server not running.');
    }
}

// ============================================
// Init
// ============================================
function init() {
    console.log('🚀 Web Monitoring v2.0 — Data Visualizer + QoS');

    // Tab navigation
    initTabs();

    // File input
    elements.fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) uploadFile(file);
        e.target.value = '';
    });

    // Clear button
    elements.clearBtn.addEventListener('click', clearData);

    // Export
    elements.exportBtn.addEventListener('click', exportCSV);

    // Pagination
    elements.prevBtn.addEventListener('click', () => {
        if (currentPage > 1) { currentPage--; fetchData(); }
    });
    elements.nextBtn.addEventListener('click', () => {
        currentPage++;
        fetchData();
    });
    elements.pageSizeSelect.addEventListener('change', (e) => {
        pageSize = parseInt(e.target.value);
        currentPage = 1;
        if (hasData) fetchData();
    });

    // QoS recalculate
    elements.recalcBtn.addEventListener('click', recalculateQoS);

    // Drag & drop
    setupDragDrop();

    // Check for existing data
    checkExistingData();
}

// Start
document.addEventListener('DOMContentLoaded', init);
