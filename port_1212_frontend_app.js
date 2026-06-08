/**
 * WebMon v5.3 — Dual Mode (Premium Edition)
 * TA Maula UGM
 *
 * FIXES v5.2:
 *  #1 — matchHeader: normalize aliases juga (fix root cause received_at tidak ter-mapping)
 *  #2 — parseDateCell: pakai parseFloat agar lebih robust untuk Unix timestamp dari CSV
 *  #3 — calculateLocalQoS: Math.abs() + deteksi desync jam RTC vs Server
 *
 * UPDATES v5.3:
 *  #4 — Single Chart dengan Parameter Dropdown Switcher (ganti 5 chart terpisah)
 *  #5 — Dynamic Table Limit (50/100/250/500/1000/3000 baris)
 *  #6 — Hapus fitur Export CSV (Mode Lokal)
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

/**
 * FIX #1: Normalize BOTH input header dan setiap alias sebelum dibandingkan.
 * Sebelumnya: "received_at" → "received at" tapi alias masih "received_at"
 *             → tidak cocok → kolom diabaikan → NO DATA latency
 * Sesudahnya: keduanya dinormalisasi → selalu cocok
 */
function normalizeStr(s) {
    return String(s || '')
        .toLowerCase()
        .trim()
        .replace(/\s*\(.*?\)\s*/g, '')  // buang "(mBar)", "(°C)", dst.
        .replace(/[_\-]/g, ' ')          // normalise underscore/dash → spasi
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

/** Format Date object → ISO 8601 dengan 'T' (valid di semua browser termasuk Safari) */
function fmtDate(d) {
    return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}.${String(d.getMilliseconds()).padStart(3, '0')}`;
}

/** Tampilan UI: ganti 'T' → spasi agar terbaca manusia */
function displayDate(isoStr) {
    if (!isoStr) return '--';
    return String(isoStr).replace('T', ' ');
}

function pad(n) { return String(n).padStart(2, '0'); }

/**
 * FIX #2: Gunakan parseFloat agar Unix timestamp dari CSV (format string/float)
 * bisa terbaca dengan benar. Contoh: "1773594000" atau "1773594000.0"
 */
function parseDateCell(val) {
    if (val === null || val === undefined || val === '') return null;

    const numVal = parseFloat(val);

    if (!isNaN(numVal)) {
        // 13-digit Unix ms (e.g. Date.now())
        if (numVal > 1_000_000_000_000) {
            const d = new Date(numVal);
            if (!isNaN(d)) return fmtDate(d);
        }
        // 9-11-digit Unix seconds (e.g. 1773594000)
        if (numVal > 100_000_000 && numVal < 10_000_000_000) {
            const d = new Date(numVal * 1000);
            if (!isNaN(d)) return fmtDate(d);
        }
        // Excel serial date — hanya jika val aslinya memang number (bukan string dari CSV)
        if (numVal < 100_000 && typeof val === 'number') {
            try {
                const p = XLSX.SSF.parse_date_code(numVal);
                if (p) return `${p.y}-${pad(p.m)}-${pad(p.d)}T${pad(p.H)}:${pad(p.M)}:${pad(p.S)}`;
            } catch (_) { /* fallthrough */ }
        }
    }

    const s = String(val).trim();

    // "YYYY-MM-DD HH:MM:SS" (space separator) → normalise ke ISO 'T'
    if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}/.test(s)) return s.replace(' ', 'T');

    // Sudah ISO (ada 'T') atau format lain → kembalikan apa adanya
    return s || null;
}

/* ============================================================
   APP OBJECT
   ============================================================ */
const App = {
    currentTab: 'live',
    liveData:   [],
    localData:  [],
    charts:     {},
    liveInterval: null,

    el: {
        tabBtns:       document.querySelectorAll('.tab-btn'),
        views:         document.querySelectorAll('.view-container'),
        location:      document.getElementById('location'),
        liveIndicator: document.getElementById('liveIndicator'),
        livePacketCount: document.getElementById('livePacketCount'),
        lastUpdateWrap:  document.getElementById('lastUpdateWrap'),
        lastUpdateTime:  document.getElementById('lastUpdateTime'),

        // Live
        liveTbody:        document.getElementById('liveTbody'),
        liveRangeDisplay: document.getElementById('liveRangeDisplay'),
        liveLatVal:       document.getElementById('liveLatVal'),
        liveLatBadge:     document.getElementById('liveLatBadge'),
        liveLossVal:      document.getElementById('liveLossVal'),
        liveLossBadge:    document.getElementById('liveLossBadge'),

        // Local
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
        console.log('💎 WebMon v5.3 Initializing...');
        this.setupEventListeners();
        this.startLivePolling();
        this.switchTab('live');
    },

    setupEventListeners() {
        this.el.tabBtns.forEach(btn =>
            btn.addEventListener('click', () => this.switchTab(btn.dataset.tab))
        );
        this.el.inputLocalFile.addEventListener('change', e => {
            if (e.target.files[0]) this.processLocalFile(e.target.files[0]);
        });
        this.el.selectLocalInterval.addEventListener('change', () => {
            if (this.localData.length > 0) this.analyzeLocalData();
        });

        // #4 (v5.4) — Pill Tab Switcher: smooth update tanpa destroy chart
        document.getElementById('chartParamTabs').addEventListener('click', e => {
            const tab = e.target.closest('.param-tab');
            if (!tab) return;
            document.querySelectorAll('#chartParamTabs .param-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            if (this.localData.length > 0) this.updateChartParam();
        });

        // #5 — Table limit switcher: hanya re-render tabel, tanpa hitung ulang QoS
        this.el.selectTableLimit.addEventListener('change', () => {
            if (this.localData.length > 0) {
                const limit = parseInt(this.el.selectTableLimit.value);
                this.el.localTableInfo.textContent = `${this.localData.length} baris (tampil ${limit} terbaru)`;
                this.renderTable(this.el.localTbody, this.localData.slice(0, limit));
            }
        });

        this.el.btnClearLocal.addEventListener('click', () => this.resetLocal());

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
       TAB MANAGEMENT
       ========================================================== */
    switchTab(tab) {
        this.currentTab = tab;
        this.el.tabBtns.forEach(btn => btn.classList.toggle('active', btn.dataset.tab === tab));
        this.el.views.forEach(v => v.classList.toggle('active',
            v.id === `view${tab.charAt(0).toUpperCase() + tab.slice(1)}`
        ));
        const hasLocal = this.localData.length > 0;
        this.el.localFileDisplay.style.display = (tab === 'local' && hasLocal) ? 'flex' : 'none';
        if (tab === 'local') {
            Object.values(this.charts).forEach(c => c && c.resize());
        }
    },

    /* ==========================================================
       LIVE MODE — POLLING
       ========================================================== */
    startLivePolling() {
        this.fetchLive();
        this.liveInterval = setInterval(() => this.fetchLive(), 5000);
    },

    async fetchLive() {
        try {
            const [dataRes, qosRes] = await Promise.all([
                fetch('/api/data?page=1&limit=50'),
                fetch('/api/qos')
            ]);
            const dataJson = await dataRes.json();
            const qosJson  = await qosRes.json();
            if (dataJson.success && dataJson.data) {
                this.liveData = dataJson.data;
                this.updateLiveUI(dataJson);
            }
            if (qosJson.success && qosJson.qos) {
                this.updateLiveQoS(qosJson.qos);
            }
        } catch (_) { /* Backend standby — silent */ }
    },

    updateLiveUI(res) {
        if (!this.liveData.length) return;
        const latest = this.liveData[0];
        this.el.location.textContent        = `📍 ${res.location || 'Stasiun Cuaca'}`;
        this.el.liveIndicator.style.display = 'flex';
        this.el.livePacketCount.textContent = `${res.pagination.totalRows} pkt`;
        this.el.lastUpdateWrap.style.display = 'flex';
        this.el.lastUpdateTime.textContent   = new Date().toLocaleTimeString('id-ID');
        this.updateGauges(latest);
        this.renderTable(this.el.liveTbody, this.liveData.slice(0, 15));
        if (this.liveData.length > 1) {
            const last = this.liveData[this.liveData.length - 1];
            this.el.liveRangeDisplay.textContent = `${displayDate(last.tlocal)} — ${displayDate(latest.tlocal)}`;
        }
    },

    updateLiveQoS(qos) {
        if (qos.latency)    this.setBadge(this.el.liveLatVal,  this.el.liveLatBadge,  qos.latency.avgMs,         'LATENCY');
        if (qos.packetLoss) this.setBadge(this.el.liveLossVal, this.el.liveLossBadge, qos.packetLoss.percentage, 'LOSS');
    },

    /* ==========================================================
       LIVE SVG GAUGES
       ========================================================== */
    updateGauges(data) {
        // Pressure
        const gP = document.getElementById('gaugePressureLive');
        const nP = document.getElementById('needlePressureLive');
        const vP = document.getElementById('valPressureLive');
        if (gP && data.airpressure != null) {
            const pct = Math.max(0, Math.min(1, (data.airpressure - 960) / 60));
            gP.style.strokeDashoffset = 251 - pct * 251;
            nP.style.transform = `rotate(${pct * 270 - 135}deg)`;
            vP.textContent = data.airpressure.toFixed(1);
        }
        // Temperature
        const lT = document.getElementById('levelTempLive');
        const vT = document.getElementById('valTempLive');
        if (lT && data.airtemperature != null) {
            const h = Math.max(0, Math.min(55, ((data.airtemperature - 0) / 45) * 55));
            lT.setAttribute('height', h.toFixed(1));
            lT.setAttribute('y', (65 - h).toFixed(1));
            vT.textContent = data.airtemperature.toFixed(1);
        }
        // Humidity
        const lH = document.getElementById('levelHumLive');
        const vH = document.getElementById('valHumLive');
        if (lH && data.airhumidity != null) {
            const fillH = Math.max(0, Math.min(1, data.airhumidity / 100)) * 90;
            lH.setAttribute('y',      (95 - fillH).toFixed(1));
            lH.setAttribute('height', fillH.toFixed(1));
            vH.textContent = data.airhumidity.toFixed(1);
        }
        // Wind speed
        const vW = document.getElementById('valWindLive');
        if (vW && data.windspeed != null) vW.textContent = data.windspeed.toFixed(2);
        // Wind direction
        const nD = document.getElementById('needleDirLive');
        const vD = document.getElementById('valDirLive');
        if (nD && data.winddirection != null) {
            nD.style.transform = `rotate(${data.winddirection.toFixed(0)}deg)`;
            vD.textContent = data.winddirection.toFixed(0);
        }
    },

    /* ==========================================================
       LOCAL MODE — SMART HEADER SCANNER
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

                // Scan baris header
                let headerRowIdx  = -1;
                let headerMapping = {};

                for (let r = 0; r < Math.min(25, raw.length); r++) {
                    const row = raw[r];
                    let matches = 0;
                    const tmpMap = {};
                    for (let c = 0; c < row.length; c++) {
                        const field = matchHeader(row[c]);  // FIX #1: aliases juga dinormalisasi
                        if (field) { matches++; tmpMap[c] = field; }
                    }
                    if (matches >= 2) { headerRowIdx = r; headerMapping = tmpMap; break; }
                }

                console.log(`📋 Header row idx: ${headerRowIdx}`, headerMapping);

                if (headerRowIdx === -1) {
                    alert('❌ Header kolom tidak ditemukan.\nPastikan file mengandung kolom seperti "Waktu", "Suhu", "Tekanan", dll.');
                    return;
                }

                // Parse baris data
                const parsedRows = [];
                for (let r = headerRowIdx + 1; r < raw.length; r++) {
                    const rawRow = raw[r];
                    const row    = {};
                    let hasData  = false;

                    for (const [colIdx, field] of Object.entries(headerMapping)) {
                        const val = rawRow[parseInt(colIdx)];
                        if (field === 'tlocal' || field === 'received_at') {
                            row[field] = parseDateCell(val);  // FIX #2: parseFloat-based
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
                this.el.localFileDisplay.style.display = (this.currentTab === 'local') ? 'flex' : 'none';

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
       STATISTIK — Min / Avg / Max per parameter (v5.3)
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
       QoS CALCULATION
       FIX #3: Math.abs() + deteksi desinkronisasi jam RTC
       ========================================================== */
    calculateLocalQoS(rows, interval) {

        // --- LATENCY ---
        let avgLat = null;
        const rowsWithReceive = rows.filter(r => r.tlocal && r.received_at);

        if (rowsWithReceive.length > 0) {
            let latSum      = 0;
            let latCount    = 0;
            let negCount    = 0;  // hitung diff negatif (indikator desync)

            rowsWithReceive.forEach(r => {
                const ts = new Date(r.tlocal).getTime();
                const tr = new Date(r.received_at).getTime();
                if (!isNaN(ts) && !isNaN(tr)) {
                    const diff = tr - ts;
                    latSum  += Math.abs(diff);  // FIX #3: abs agar desync tidak buang data
                    latCount++;
                    if (diff < 0) negCount++;
                }
            });

            if (latCount > 0) {
                avgLat = Math.round(latSum / latCount);

                // Info desync untuk laporan sidang
                const desyncPct   = Math.round((negCount / latCount) * 100);
                const desyncNote  = negCount > 0
                    ? ` | ⚠ Desync jam alat: ${negCount}/${latCount} paket (${desyncPct}%)`
                    : '';

                // Simpan note desync untuk ditampilkan di bawah QoS
                this._desyncNote = desyncNote;
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

            // Tampilkan info packet loss + desync note (jika ada)
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
       CHARTS — v5.4: Auto-detect X label, smooth tab transition
       ========================================================== */

    /**
     * Buat label X-axis dengan auto-detect format:
     *  - Same-day  → "HH:mm"
     *  - Multi-day → "DD/MM HH:mm"
     */
    _makeXLabels(data) {
        // Parsing tanggal dari setiap row
        const dates = data.map(r => new Date(r.tlocal));

        // Deteksi apakah semua data berada di hari yang sama
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

    /**
     * renderLocalCharts — dipanggil saat file baru di-load atau interval berganti.
     * Selalu rebuild chart (destroy + buat baru) agar data & labels ter-mapping ulang.
     */
    renderLocalCharts(rows) {
        const data = [...rows].sort(
            (a, b) => new Date(a.tlocal).getTime() - new Date(b.tlocal).getTime()
        );

        // Simpan sorted data + labels ke state agar updateChartParam bisa pakai ulang
        this._chartSortedData = data;
        this._chartLabels     = this._makeXLabels(data);

        // Baca tab aktif
        const activeTab = document.querySelector('#chartParamTabs .param-tab.active');
        const key   = activeTab?.dataset.param || 'airpressure';
        const label = activeTab?.dataset.label || 'Tekanan Udara';
        const unit  = activeTab?.dataset.unit  || 'mBar';
        const color = activeTab?.dataset.color || '#06b6d4';

        // Update info label header
        if (this.el.mainChartLabel) this.el.mainChartLabel.textContent = label;
        if (this.el.mainChartUnit)  this.el.mainChartUnit.textContent  = unit;

        const values = data.map(r => r[key]);
        this.buildChart('mainChart', `${label} (${unit})`, this._chartLabels, values, color);
    },

    /**
     * updateChartParam — dipanggil HANYA saat klik Pill Tab.
     * Update data existing chart instance → smooth transition tanpa destroy.
     * v5.4+: Grid & tooltip border color juga diupdate secara dinamis.
     */
    updateChartParam() {
        const chart = this.charts['mainChart'];
        const data  = this._chartSortedData;
        if (!chart || !data) return;

        const activeTab = document.querySelector('#chartParamTabs .param-tab.active');
        const key   = activeTab?.dataset.param || 'airpressure';
        const label = activeTab?.dataset.label || 'Tekanan Udara';
        const unit  = activeTab?.dataset.unit  || 'mBar';
        const color = activeTab?.dataset.color || '#06b6d4';

        // Update header label
        if (this.el.mainChartLabel) this.el.mainChartLabel.textContent = label;
        if (this.el.mainChartUnit)  this.el.mainChartUnit.textContent  = unit;

        const values = data.map(r => r[key]);
        const ds = chart.data.datasets[0];

        // Update dataset properties in-place
        ds.data            = values;
        ds.label           = `${label} (${unit})`;
        ds.borderColor     = color;
        ds.backgroundColor = color + '28';
        ds.pointRadius     = values.length > 200 ? 0 : 2;

        // ✨ Dynamic grid & tooltip color — menyesuaikan warna Pill Tab aktif
        chart.options.scales.x.grid.color         = color + '12'; // sangat subtle
        chart.options.scales.y.grid.color         = color + '20'; // sedikit lebih tebal
        chart.options.plugins.tooltip.borderColor = color + '50'; // border tooltip

        // Performa guard: disable animasi jika data > 500 baris
        const duration = data.length > 500 ? 0 : 800;
        chart.update({ duration, easing: 'easeInOutQuart' });
    },

    /**
     * buildChart — hanya dipanggil untuk INITIAL creation atau REBUILD (file baru/interval).
     * v5.4+: Entrance animation "rise from bottom" + dynamic grid/tooltip colors.
     */
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
                // ✨ Entrance animation: titik-titik "naik" dari bawah grafik
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
                        borderColor: color + '50', // ✨ Dynamic tooltip border
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
                        grid: { color: color + '12' } // ✨ Dynamic grid color X
                    },
                    y: {
                        ticks: { color: '#64748b', font: { size: 10 } },
                        grid: { color: color + '20' } // ✨ Dynamic grid color Y
                    }
                }
            }
        });
    },

    /* ==========================================================
       TABLE RENDER — displayDate() strips 'T', null-safe values
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

    /* exportLocalCSV() — dihapus di v5.3 (fitur export CSV Mode Lokal dihilangkan) */

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
        // Reset label chart ke default
        if (this.el.mainChartLabel) this.el.mainChartLabel.textContent = 'Tekanan Udara';
        if (this.el.mainChartUnit)  this.el.mainChartUnit.textContent  = 'mBar';
        if (this.el.selectTableLimit) this.el.selectTableLimit.value   = '100';
        // Reset pill tabs ke Tekanan (index 0)
        document.querySelectorAll('#chartParamTabs .param-tab').forEach((t, i) =>
            t.classList.toggle('active', i === 0)
        );
        // Clear statistik
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
