/**
 * WebMon QoS v1.0 — QoS Analysis Only (Port 1212)
 * TA Maula UGM
 *
 * Derived from WebMon v5.3 — semua fitur Mode Live dihapus.
 * Hanya menyisakan:
 *   - Upload & parse file Excel/CSV
 *   - Grafik tren parameter (single chart + pill tab switcher)
 *   - Statistik ringkasan (Min/Avg/Max)
 *   - QoS Analysis: Latency & Packet Loss (Standar TIPHON)
 *   - Historical Data Table
 **/
'use strict';

/* ============================================================
   COLUMN KEYWORDS (sync with backend sensor.js COLUMN_MAP)
   ============================================================ */
const COLUMN_MAP = {
    tlocal:         ['tlocal','waktu','waktu_kirim','time','datetime','tanggal','timestamp','date_time','date','t_send','tsend','send_time','waktu kirim'],
    received_at:    ['received_at','waktu_terima','treceive','t_receive','receive_time','received','waktu terima','waktu_diterima','server_time'],
    airpressure:    ['airpressure','tekanan','tekanan_udara','pressure','air_pressure','tekanan udara'],
    airtemperature: ['airtemperature','suhu','suhu_udara','temperature','temp','air_temperature','suhu udara'],
    airhumidity:    ['airhumidity','kelembapan','kelembapan_udara','humidity','air_humidity','kelembapan udara','rh'],
    windspeed:      ['windspeed','kecepatan_angin','kecepatan angin','wind_speed','ws','angin_kecepatan'],
    winddirection:  ['winddirection','arah_angin','arah angin','wind_direction','wd','angin_arah'],
    latitude:       ['latitude','lat','lintang'],
    longitude:      ['longitude','lng','lon','long','bujur']
};

function normalizeStr(s) {
    return String(s || '')
        .toLowerCase()
        .trim()
        .replace(/\s*\(.*?\)\s*/g, '')
        .replace(/[_\-]/g, ' ')
        .trim();
}

function matchHeader(raw) {
    const cleaned = normalizeStr(raw);
    for (const [field, aliases] of Object.entries(COLUMN_MAP)) {
        if (aliases.some(a => normalizeStr(a) === cleaned)) return field;
    }
    return null;
}

/* ============================================================
   HELPERS — DATE PARSING
   ============================================================ */
function fmtDate(d) {
    return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}.${String(d.getMilliseconds()).padStart(3, '0')}`;
}

function displayDate(isoStr) {
    if (!isoStr) return '--';
    return String(isoStr).replace('T', ' ');
}

function pad(n) { return String(n).padStart(2, '0'); }

function parseDateCell(val) {
    if (val === null || val === undefined || val === '') return null;

    const numVal = parseFloat(val);

    if (!isNaN(numVal)) {
        if (numVal > 1_000_000_000_000) {
            const d = new Date(numVal);
            if (!isNaN(d)) return fmtDate(d);
        }
        if (numVal > 100_000_000 && numVal < 10_000_000_000) {
            const d = new Date(numVal * 1000);
            if (!isNaN(d)) return fmtDate(d);
        }
        if (numVal < 100_000 && typeof val === 'number') {
            try {
                const p = XLSX.SSF.parse_date_code(numVal);
                if (p) return `${p.y}-${pad(p.m)}-${pad(p.d)}T${pad(p.H)}:${pad(p.M)}:${pad(p.S)}`;
            } catch (_) { /* fallthrough */ }
        }
    }

    const s = String(val).trim();
    if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}/.test(s)) return s.replace(' ', 'T');
    return s || null;
}

/* ============================================================
   APP OBJECT
   ============================================================ */
const App = {
    localData: [],
    charts:    {},

    el: {
        inputLocalFile:      document.getElementById('inputLocalFile'),
        selectLocalInterval: document.getElementById('selectLocalInterval'),
        selectTableLimit:    document.getElementById('selectTableLimit'),
        localEmptyState:     document.getElementById('localEmptyState'),
        localContent:        document.getElementById('localContent'),
        toolbarFileName:     document.getElementById('toolbarFileName'),
        toolbarRowCount:     document.getElementById('toolbarRowCount'),
        localRangeDisplay:   document.getElementById('localRangeDisplay'),
        localTableInfo:      document.getElementById('localTableInfo'),
        localTbody:          document.getElementById('localTbody'),
        localLatVal:         document.getElementById('localLatVal'),
        localLatBadge:       document.getElementById('localLatBadge'),
        localLossVal:        document.getElementById('localLossVal'),
        localLossBadge:      document.getElementById('localLossBadge'),
        localLossCalcInfo:   document.getElementById('localLossCalcInfo'),
        stressTestBanner:    document.getElementById('stressTestBanner'),
        btnClearLocal:       document.getElementById('btnClearLocal'),
        localFileDisplay:    document.getElementById('localFileDisplay'),
        activeFileName:      document.getElementById('activeFileName'),
        mainChartLabel:      document.getElementById('mainChartLabel'),
        mainChartUnit:       document.getElementById('mainChartUnit'),
        dropZone:            document.getElementById('dropZone'),
    },

    TIPHON: {
        LATENCY: [
            { cat:'Sangat Baik', max:150,      color:'#10b981' },
            { cat:'Baik',        max:300,      color:'#3b82f6' },
            { cat:'Sedang',      max:350,      color:'#f59e0b' },
            { cat:'Buruk',       max:Infinity, color:'#ef4444' }
        ],
        LOSS: [
            { cat:'Sangat Baik', max:3,        color:'#10b981' },
            { cat:'Baik',        max:15,       color:'#3b82f6' },
            { cat:'Sedang',      max:25,       color:'#f59e0b' },
            { cat:'Buruk',       max:Infinity, color:'#ef4444' }
        ]
    },

    /* ==========================================================
       INIT
       ========================================================== */
    init() {
        console.log('📊 WebMon QoS v1.0 Initializing...');
        this.setupEventListeners();
    },

    setupEventListeners() {
        this.el.inputLocalFile.addEventListener('change', e => {
            if (e.target.files[0]) this.processLocalFile(e.target.files[0]);
        });

        this.el.selectLocalInterval.addEventListener('change', () => {
            if (this.localData.length > 0) this.analyzeLocalData();
        });

        // Pill Tab Switcher
        document.getElementById('chartParamTabs').addEventListener('click', e => {
            const tab = e.target.closest('.param-tab');
            if (!tab) return;
            document.querySelectorAll('#chartParamTabs .param-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            if (this.localData.length > 0) this.updateChartParam();
        });

        // Table limit switcher
        this.el.selectTableLimit.addEventListener('change', () => {
            if (this.localData.length > 0) {
                const limit = parseInt(this.el.selectTableLimit.value);
                this.el.localTableInfo.textContent = `${this.localData.length} baris (tampil ${limit} terbaru)`;
                this.renderTable(this.el.localTbody, this.localData.slice(0, limit));
            }
        });

        this.el.btnClearLocal.addEventListener('click', () => this.resetLocal());

        // Drag & Drop
        window.addEventListener('dragenter', () => this.el.dropZone.classList.add('active'));
        this.el.dropZone.addEventListener('dragleave', () => this.el.dropZone.classList.remove('active'));
        this.el.dropZone.addEventListener('dragover', e => e.preventDefault());
        this.el.dropZone.addEventListener('drop', e => {
            e.preventDefault();
            this.el.dropZone.classList.remove('active');
            const file = e.dataTransfer.files[0];
            if (file) this.processLocalFile(file);
        });
    },

    /* ==========================================================
       FILE PROCESSING — SMART HEADER SCANNER
       ========================================================== */
    processLocalFile(file) {
        const reader = new FileReader();
        reader.onload = e => {
            try {
                const data     = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: false });
                const sheet    = workbook.Sheets[workbook.SheetNames[0]];

                const raw = XLSX.utils.sheet_to_json(sheet, {
                    header: 1, defval: '', blankrows: false
                });

                let headerRowIdx  = -1;
                let headerMapping = {};

                for (let r = 0; r < Math.min(25, raw.length); r++) {
                    const row = raw[r];
                    let matches = 0;
                    const tmpMap = {};
                    for (let c = 0; c < row.length; c++) {
                        const field = matchHeader(row[c]);
                        if (field) { matches++; tmpMap[c] = field; }
                    }
                    if (matches >= 2) { headerRowIdx = r; headerMapping = tmpMap; break; }
                }

                console.log(`📋 Header row idx: ${headerRowIdx}`, headerMapping);

                if (headerRowIdx === -1) {
                    alert('❌ Header kolom tidak ditemukan.\nPastikan file mengandung kolom seperti "Waktu", "Suhu", "Tekanan", dll.');
                    return;
                }

                const parsedRows = [];
                for (let r = headerRowIdx + 1; r < raw.length; r++) {
                    const rawRow = raw[r];
                    const row    = {};
                    let hasData  = false;

                    for (const [colIdx, field] of Object.entries(headerMapping)) {
                        const val = rawRow[parseInt(colIdx)];
                        if (field === 'tlocal' || field === 'received_at') {
                            row[field] = parseDateCell(val);
                            if (row[field]) hasData = true;
                        } else {
                            const num = parseFloat(val);
                            row[field] = isNaN(num) ? null : num;
                            if (row[field] !== null) hasData = true;
                        }
                    }

                    if (hasData && row.tlocal) parsedRows.push(row);
                }

                if (parsedRows.length === 0) {
                    alert('❌ Tidak ada baris data valid.\nPastikan kolom "Waktu/tlocal" terisi.');
                    return;
                }

                this.localData = parsedRows;
                this.el.activeFileName.textContent  = file.name;
                this.el.toolbarFileName.textContent = file.name;
                this.el.toolbarRowCount.textContent = `${parsedRows.length} baris data`;

                this.el.localEmptyState.style.display  = 'none';
                this.el.localContent.style.display     = 'block';
                this.el.localFileDisplay.style.display = 'flex';

                this.analyzeLocalData();
                console.log(`✅ Loaded ${parsedRows.length} records from "${file.name}"`);

            } catch (err) {
                console.error('File parse error:', err);
                alert(`❌ Gagal membaca file: ${err.message}`);
            }
        };
        reader.readAsArrayBuffer(file);
    },

    /* ==========================================================
       LOCAL ANALYSIS
       ========================================================== */
    analyzeLocalData() {
        const rows        = this.localData;
        const intervalSec = parseInt(this.el.selectLocalInterval.value);
        const limit       = parseInt(this.el.selectTableLimit.value) || 100;

        const sortedForRange = [...rows].sort(
            (a, b) => new Date(a.tlocal).getTime() - new Date(b.tlocal).getTime()
        );
        const first = displayDate(sortedForRange[0]?.tlocal) ?? '--';
        const last  = displayDate(sortedForRange[sortedForRange.length - 1]?.tlocal) ?? '--';
        this.el.localRangeDisplay.textContent = `${first} → ${last}`;
        this.el.localTableInfo.textContent    = `${rows.length} baris (tampil ${limit} terbaru)`;

        this.renderLocalCharts(rows);
        this.renderStats(rows);
        this.renderTable(this.el.localTbody, rows.slice(0, limit));
        this.calculateLocalQoS(rows, intervalSec);
    },

    /* ==========================================================
       STATISTIK — Min / Avg / Max per parameter
       ========================================================== */
    renderStats(rows) {
        const params = [
            { key: 'airpressure',    dec: 2, avg: 'statAvgPressure',  min: 'statMinPressure',  max: 'statMaxPressure'  },
            { key: 'airtemperature', dec: 1, avg: 'statAvgTemp',       min: 'statMinTemp',       max: 'statMaxTemp'       },
            { key: 'airhumidity',    dec: 1, avg: 'statAvgHumidity',   min: 'statMinHumidity',   max: 'statMaxHumidity'   },
            { key: 'windspeed',      dec: 2, avg: 'statAvgWindSpeed',  min: 'statMinWindSpeed',  max: 'statMaxWindSpeed'  },
            { key: 'winddirection',  dec: 0, avg: 'statAvgWindDir',    min: 'statMinWindDir',    max: 'statMaxWindDir'    },
        ];

        const subEl = document.getElementById('statsSubInfo');
        if (subEl) subEl.textContent = `${rows.length} data`;

        params.forEach(p => {
            const vals  = rows.map(r => r[p.key]).filter(v => v != null && !isNaN(v));
            const avgEl = document.getElementById(p.avg);
            const minEl = document.getElementById(p.min);
            const maxEl = document.getElementById(p.max);
            if (!avgEl || !minEl || !maxEl) return;

            if (vals.length === 0) {
                avgEl.textContent = '--';
                minEl.textContent = '--';
                maxEl.textContent = '--';
                return;
            }

            const sum = vals.reduce((a, b) => a + b, 0);
            avgEl.textContent = (sum / vals.length).toFixed(p.dec);
            minEl.textContent = Math.min(...vals).toFixed(p.dec);
            maxEl.textContent = Math.max(...vals).toFixed(p.dec);
        });
    },

    /* ==========================================================
       QoS CALCULATION — Standar TIPHON
       ========================================================== */
    calculateLocalQoS(rows, interval) {

        // --- LATENCY ---
        let avgLat = null;
        const rowsWithReceive = rows.filter(r => r.tlocal && r.received_at);

        if (rowsWithReceive.length > 0) {
            let latSum   = 0;
            let latCount = 0;
            let negCount = 0;

            rowsWithReceive.forEach(r => {
                const ts = new Date(r.tlocal).getTime();
                const tr = new Date(r.received_at).getTime();
                if (!isNaN(ts) && !isNaN(tr)) {
                    const diff = tr - ts;
                    latSum  += Math.abs(diff);
                    latCount++;
                    if (diff < 0) negCount++;
                }
            });

            if (latCount > 0) {
                avgLat = Math.round(latSum / latCount);
                const desyncPct  = Math.round((negCount / latCount) * 100);
                this._desyncNote = negCount > 0
                    ? ` | ⚠ Desync jam alat: ${negCount}/${latCount} paket (${desyncPct}%)`
                    : '';
            }
        } else {
            this._desyncNote = '';
        }

        this.setBadge(this.el.localLatVal, this.el.localLatBadge, avgLat, 'LATENCY');

        // --- PACKET LOSS ---
        const timestamps = rows
            .map(r => new Date(r.tlocal).getTime())
            .filter(t => !isNaN(t));

        if (timestamps.length >= 2) {
            const tMin     = Math.min(...timestamps);
            const tMax     = Math.max(...timestamps);
            const rangeMs  = tMax - tMin;
            const expected = Math.round(rangeMs / (interval * 1000)) + 1;
            const actual   = rows.length;

            let lossPct        = 0;
            let stressDetected = false;
            if (actual >= expected) {
                lossPct = 0; stressDetected = true;
            } else {
                lossPct = ((expected - actual) / expected) * 100;
            }

            this.setBadge(this.el.localLossVal, this.el.localLossBadge, parseFloat(lossPct.toFixed(1)), 'LOSS');

            this.el.localLossCalcInfo.textContent =
                `Target: ${expected} paket | Aktual: ${actual} paket | Rentang: ${Math.round(rangeMs / 60000)} menit` +
                (this._desyncNote || '');

            this.el.stressTestBanner.style.display = stressDetected ? 'block' : 'none';

        } else {
            this.setBadge(this.el.localLossVal, this.el.localLossBadge, null, 'LOSS');
            this.el.localLossCalcInfo.textContent = 'Tidak cukup data untuk menghitung Packet Loss.' + (this._desyncNote || '');
            this.el.stressTestBanner.style.display = 'none';
        }
    },

    setBadge(valEl, badgeEl, value, type) {
        if (value === null || value === undefined) {
            valEl.textContent         = '--';
            badgeEl.textContent       = 'NO DATA';
            badgeEl.style.color       = 'var(--text-lo)';
            badgeEl.style.borderColor = 'var(--text-lo)';
            badgeEl.style.background  = 'transparent';
            return;
        }
        valEl.textContent         = value;
        const tiers = this.TIPHON[type];
        const tier  = tiers.find(t => value <= t.max) || tiers[tiers.length - 1];
        badgeEl.textContent       = tier.cat;
        badgeEl.style.color       = tier.color;
        badgeEl.style.borderColor = tier.color;
        badgeEl.style.background  = tier.color + '18';
    },

    /* ==========================================================
       CHARTS
       ========================================================== */
    _makeXLabels(data) {
        const dates = data.map(r => new Date(r.tlocal));
        const firstDate = dates[0];
        const isMultiDay = dates.some(d =>
            d.getFullYear() !== firstDate.getFullYear() ||
            d.getMonth()    !== firstDate.getMonth()    ||
            d.getDate()     !== firstDate.getDate()
        );

        return dates.map(d => {
            const HH = pad(d.getHours());
            const mm = pad(d.getMinutes());
            if (isMultiDay) {
                const DD = pad(d.getDate());
                const MM = pad(d.getMonth() + 1);
                return `${DD}/${MM} ${HH}:${mm}`;
            }
            return `${HH}:${mm}`;
        });
    },

    renderLocalCharts(rows) {
        const data = [...rows].sort(
            (a, b) => new Date(a.tlocal).getTime() - new Date(b.tlocal).getTime()
        );

        this._chartSortedData = data;
        this._chartLabels     = this._makeXLabels(data);

        const activeTab = document.querySelector('#chartParamTabs .param-tab.active');
        const key   = activeTab?.dataset.param || 'airpressure';
        const label = activeTab?.dataset.label || 'Tekanan Udara';
        const unit  = activeTab?.dataset.unit  || 'mBar';
        const color = activeTab?.dataset.color || '#06b6d4';

        if (this.el.mainChartLabel) this.el.mainChartLabel.textContent = label;
        if (this.el.mainChartUnit)  this.el.mainChartUnit.textContent  = unit;

        const values = data.map(r => r[key]);
        this.buildChart('mainChart', `${label} (${unit})`, this._chartLabels, values, color);
    },

    updateChartParam() {
        const chart = this.charts['mainChart'];
        const data  = this._chartSortedData;
        if (!chart || !data) return;

        const activeTab = document.querySelector('#chartParamTabs .param-tab.active');
        const key   = activeTab?.dataset.param || 'airpressure';
        const label = activeTab?.dataset.label || 'Tekanan Udara';
        const unit  = activeTab?.dataset.unit  || 'mBar';
        const color = activeTab?.dataset.color || '#06b6d4';

        if (this.el.mainChartLabel) this.el.mainChartLabel.textContent = label;
        if (this.el.mainChartUnit)  this.el.mainChartUnit.textContent  = unit;

        const values = data.map(r => r[key]);
        const ds = chart.data.datasets[0];

        ds.data            = values;
        ds.label           = `${label} (${unit})`;
        ds.borderColor     = color;
        ds.backgroundColor = color + '28';
        ds.pointRadius     = values.length > 200 ? 0 : 2;

        chart.options.scales.x.grid.color         = color + '12';
        chart.options.scales.y.grid.color         = color + '20';
        chart.options.plugins.tooltip.borderColor = color + '50';

        const duration = data.length > 500 ? 0 : 800;
        chart.update({ duration, easing: 'easeInOutQuart' });
    },

    buildChart(id, label, labels, values, color) {
        if (this.charts[id]) this.charts[id].destroy();
        const ctx = document.getElementById(id)?.getContext('2d');
        if (!ctx) return;

        const isLarge = values.length > 500;

        this.charts[id] = new Chart(ctx, {
            type: 'line',
            data: {
                labels,
                datasets: [{
                    label,
                    data: values,
                    borderColor: color,
                    backgroundColor: color + '18',
                    borderWidth: 2,
                    pointRadius: values.length > 200 ? 0 : 2,
                    pointHoverRadius: 5,
                    tension: 0.35,
                    fill: true,
                    spanGaps: false
                }]
            },
            options: {
                animation: {
                    duration: isLarge ? 0 : 1100,
                    easing: 'easeOutQuart'
                },
                animations: isLarge ? {} : {
                    y: {
                        duration: 1100,
                        easing: 'easeOutCubic',
                        from: (ctx) => ctx.chart.chartArea
                            ? ctx.chart.chartArea.bottom
                            : ctx.chart.height
                    }
                },
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: 'rgba(15,23,42,0.95)',
                        borderColor: color + '50',
                        borderWidth: 1,
                        titleColor: '#94a3b8', bodyColor: '#f1f5f9',
                        padding: 10, cornerRadius: 8
                    }
                },
                scales: {
                    x: {
                        ticks: {
                            color: '#64748b',
                            maxRotation: 0,
                            autoSkip: true,
                            maxTicksLimit: labels.length > 500 ? 8 : 10,
                            font: { size: 10 }
                        },
                        grid: { color: color + '12' }
                    },
                    y: {
                        ticks: { color: '#64748b', font: { size: 10 } },
                        grid: { color: color + '20' }
                    }
                }
            }
        });
    },

    /* ==========================================================
       TABLE RENDER
       ========================================================== */
    renderTable(tbody, rows) {
        if (!rows || !rows.length) {
            tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">
                <span class="spinner"></span> Tidak ada data
            </td></tr>`;
            return;
        }
        const safe = (v, dec) => (v != null && !isNaN(v))
            ? Number(v).toFixed(dec)
            : '<span style="color:var(--text-lo)">--</span>';

        tbody.innerHTML = rows.map((r, i) => `
            <tr>
                <td>${i + 1}</td>
                <td>${displayDate(r.tlocal)}</td>
                <td>${safe(r.airpressure, 2)}</td>
                <td>${safe(r.airtemperature, 1)}</td>
                <td>${safe(r.airhumidity, 1)}</td>
                <td>${safe(r.windspeed, 2)}</td>
                <td>${safe(r.winddirection, 0)}${r.winddirection != null ? '°' : ''}</td>
            </tr>
        `).join('');
    },

    /* ==========================================================
       RESET LOCAL
       ========================================================== */
    resetLocal() {
        this.localData    = [];
        this._desyncNote  = '';
        this.el.localContent.style.display      = 'none';
        this.el.localEmptyState.style.display   = 'flex';
        this.el.localFileDisplay.style.display  = 'none';
        this.el.inputLocalFile.value            = '';
        Object.values(this.charts).forEach(c => c && c.destroy());
        this.charts = {};
        if (this.el.mainChartLabel) this.el.mainChartLabel.textContent = 'Tekanan Udara';
        if (this.el.mainChartUnit)  this.el.mainChartUnit.textContent  = 'mBar';
        if (this.el.selectTableLimit) this.el.selectTableLimit.value   = '100';
        document.querySelectorAll('#chartParamTabs .param-tab').forEach((t, i) =>
            t.classList.toggle('active', i === 0)
        );
        ['statAvgPressure','statMinPressure','statMaxPressure',
         'statAvgTemp','statMinTemp','statMaxTemp',
         'statAvgHumidity','statMinHumidity','statMaxHumidity',
         'statAvgWindSpeed','statMinWindSpeed','statMaxWindSpeed',
         'statAvgWindDir','statMinWindDir','statMaxWindDir'
        ].forEach(id => { const el = document.getElementById(id); if (el) el.textContent = '--'; });
        const subEl = document.getElementById('statsSubInfo');
        if (subEl) subEl.textContent = '--';
    }
};

document.addEventListener('DOMContentLoaded', () => App.init());
