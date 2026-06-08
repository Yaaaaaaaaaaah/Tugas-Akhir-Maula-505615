const express = require('express');
const router = express.Router();
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// ============================================
// Multer config — store uploaded files in /uploads
// ============================================
const uploadDir = path.join(__dirname, '../uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });

const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, uploadDir),
    filename: (req, file, cb) => cb(null, 'current_data' + path.extname(file.originalname))
});

const upload = multer({
    storage,
    fileFilter: (req, file, cb) => {
        const ext = path.extname(file.originalname).toLowerCase();
        if (['.xlsx', '.xls', '.csv'].includes(ext)) {
            cb(null, true);
        } else {
            cb(new Error('Only .xlsx, .xls, and .csv files are allowed'));
        }
    },
    limits: { fileSize: 50 * 1024 * 1024 }
});

// ============================================
// In-memory data store
// ============================================
let activeData = {
    filename: null,
    location: null,
    rows: [],
    summary: null,
    qos: null
};

// ============================================
// Column name mapping — flexible header recognition
// ============================================
const COLUMN_MAP = {
    // Waktu kirim (T_send)
    tlocal: ['tlocal', 'waktu', 'waktu_kirim', 'time', 'datetime', 'tanggal', 'timestamp', 'date_time', 'date', 't_send', 'tsend', 'send_time', 'waktu kirim'],
    // Waktu terima (T_receive)
    received_at: ['received_at', 'waktu_terima', 'treceive', 't_receive', 'receive_time', 'received', 'waktu terima', 'waktu_diterima', 'server_time'],
    // Tekanan Udara
    airpressure: ['airpressure', 'tekanan', 'tekanan_udara', 'pressure', 'air_pressure', 'tekanan udara'],
    // Suhu
    airtemperature: ['airtemperature', 'suhu', 'suhu_udara', 'temperature', 'temp', 'air_temperature', 'suhu udara'],
    // Kelembapan
    airhumidity: ['airhumidity', 'kelembapan', 'kelembapan_udara', 'humidity', 'air_humidity', 'kelembapan udara', 'rh'],
    // Kecepatan Angin
    windspeed: ['windspeed', 'kecepatan_angin', 'kecepatan angin', 'wind_speed', 'ws', 'angin_kecepatan'],
    // Arah Angin
    winddirection: ['winddirection', 'arah_angin', 'arah angin', 'wind_direction', 'wd', 'angin_arah'],
    // Koordinat
    latitude: ['latitude', 'lat', 'lintang'],
    longitude: ['longitude', 'lng', 'lon', 'long', 'bujur']
};

function mapColumnName(header) {
    const lower = header.toLowerCase().trim();
    for (const [key, aliases] of Object.entries(COLUMN_MAP)) {
        if (aliases.includes(lower)) return key;
    }
    // Strip content in parentheses, e.g. "Tekanan Udara (mBar)" → "tekanan udara"
    const stripped = lower.replace(/\s*\(.*?\)\s*/g, '').trim();
    if (stripped !== lower) {
        for (const [key, aliases] of Object.entries(COLUMN_MAP)) {
            if (aliases.includes(stripped)) return key;
        }
    }
    return null;
}

// ============================================
// Parse datetime from Excel cell
// ============================================
function parseDateTimeCell(cell) {
    if (!cell || cell.v === undefined || cell.v === null || cell.v === '') return null;

    if (typeof cell.v === 'number') {
        // UNIX timestamp (seconds since epoch) — typically > 1,000,000,000
        if (cell.v > 100000000) {
            // Check if it's in milliseconds (13 digits) or seconds (10 digits)
            const isMs = cell.v > 1000000000000;
            const d = new Date(isMs ? cell.v : cell.v * 1000);
            if (!isNaN(d.getTime())) {
                return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}:${String(d.getSeconds()).padStart(2, '0')}.${String(d.getMilliseconds()).padStart(3, '0')}`;
            }
        }

        // Excel date serial number (typically < 100000)
        const date = XLSX.SSF.parse_date_code(cell.v);
        if (date) {
            return `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')} ${String(date.H).padStart(2, '0')}:${String(date.M).padStart(2, '0')}:${String(date.S).padStart(2, '0')}`;
        }

        // Fallback: return as string
        return String(cell.v);
    }

    // Handle string UNIX timestamps (e.g. from re-uploaded CSV exports)
    const strVal = String(cell.v).trim();
    if (/^\d{9,13}$/.test(strVal)) {
        const unixVal = parseInt(strVal);
        const isMs = unixVal > 1000000000000;
        const d = new Date(isMs ? unixVal : unixVal * 1000);
        if (!isNaN(d.getTime())) {
            return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}:${String(d.getSeconds()).padStart(2, '0')}.${String(d.getMilliseconds()).padStart(3, '0')}`;
        }
    }

    return strVal;
}

// ============================================
// Parse numeric value — keep null if empty
// ============================================
function parseNumericCell(cell) {
    if (!cell || cell.v === undefined || cell.v === null || cell.v === '') return null;
    const num = parseFloat(cell.v);
    return isNaN(num) ? null : num;
}

// ============================================
// Parse Excel/CSV file
// ============================================
function parseFile(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Try to detect location + coordinates from the first few rows
    let locationName = null;
    let coordinates = null;
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');

    for (let r = range.s.r; r <= Math.min(range.s.r + 10, range.e.r); r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            const cell = sheet[cellRef];
            if (cell && cell.v) {
                const val = String(cell.v).toLowerCase();
                const raw = String(cell.v);

                // Detect location name
                if (!locationName && (val.includes('lokasi') || val.includes('location') || val.includes('stasiun') || val.includes('station'))) {
                    const colonIdx = raw.indexOf(':');
                    if (colonIdx !== -1) {
                        locationName = raw.substring(colonIdx + 1).trim();
                    } else {
                        const nextCellRef = XLSX.utils.encode_cell({ r, c: c + 1 });
                        const nextCell = sheet[nextCellRef];
                        if (nextCell && nextCell.v) {
                            locationName = String(nextCell.v).trim();
                        }
                    }
                }

                // Detect coordinates
                if (!coordinates && (val.includes('koordinat') || val.includes('coordinate') || val.includes('coord') || val.includes('gps'))) {
                    const colonIdx = raw.indexOf(':');
                    if (colonIdx !== -1) {
                        coordinates = raw.substring(colonIdx + 1).trim();
                    } else {
                        const nextCellRef = XLSX.utils.encode_cell({ r, c: c + 1 });
                        const nextCell = sheet[nextCellRef];
                        if (nextCell && nextCell.v) {
                            coordinates = String(nextCell.v).trim();
                        }
                    }
                }
            }
        }
    }

    // Combine location name + coordinates
    let location = null;
    if (locationName && coordinates) {
        location = `${locationName} (${coordinates})`;
    } else if (locationName) {
        location = locationName;
    } else if (coordinates) {
        location = coordinates;
    }

    // Find the header row (first row with >=2 recognized columns to allow simpler data)
    let headerRow = -1;
    for (let r = range.s.r; r <= Math.min(range.s.r + 15, range.e.r); r++) {
        let matchCount = 0;
        for (let c = range.s.c; c <= range.e.c; c++) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            const cell = sheet[cellRef];
            if (cell && cell.v) {
                const mapped = mapColumnName(String(cell.v));
                if (mapped) matchCount++;
            }
        }
        if (matchCount >= 2) {
            headerRow = r;
            break;
        }
    }

    if (headerRow === -1) headerRow = range.s.r;

    // Read headers
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
        const cellRef = XLSX.utils.encode_cell({ r: headerRow, c });
        const cell = sheet[cellRef];
        headers.push(cell ? String(cell.v).trim() : `col_${c}`);
    }

    // Map headers to field names
    const columnMapping = {};
    headers.forEach((h, idx) => {
        const mapped = mapColumnName(h);
        if (mapped) columnMapping[idx] = mapped;
    });

    // Check which fields we have
    const mappedFields = new Set(Object.values(columnMapping));
    const hasReceivedAt = mappedFields.has('received_at');

    // Read data rows
    const rows = [];
    for (let r = headerRow + 1; r <= range.e.r; r++) {
        const row = {};
        let hasData = false;

        for (let c = range.s.c; c <= range.e.c; c++) {
            const fieldName = columnMapping[c - range.s.c];
            if (!fieldName) continue;

            const cellRef = XLSX.utils.encode_cell({ r, c });
            const cell = sheet[cellRef];

            if (fieldName === 'tlocal' || fieldName === 'received_at') {
                row[fieldName] = parseDateTimeCell(cell);
                if (row[fieldName]) hasData = true;
            } else {
                row[fieldName] = parseNumericCell(cell);
                if (row[fieldName] !== null) hasData = true;
            }
        }

        if (hasData && row.tlocal) {
            // Keep nulls as null — don't fill with 0
            rows.push(row);
        }
    }

    // Fallback: if no location from header, try coordinates from first data row
    if (!location && rows.length > 0) {
        const firstRow = rows[0];
        if (firstRow.latitude != null && firstRow.longitude != null) {
            location = `${firstRow.latitude}, ${firstRow.longitude}`;
        }
    }

    return { location, rows, hasReceivedAt };
}

// ============================================
// Compute summary (averages) — only from non-null values
// ============================================
function computeSummary(rows) {
    if (!rows.length) return null;

    const fields = ['airpressure', 'airhumidity', 'airtemperature', 'windspeed', 'winddirection'];
    const sums = {};
    const counts = {};

    fields.forEach(f => { sums[f] = 0; counts[f] = 0; });

    rows.forEach(row => {
        fields.forEach(f => {
            if (row[f] !== null && row[f] !== undefined && !isNaN(row[f])) {
                sums[f] += row[f];
                counts[f]++;
            }
        });
    });

    const averages = {};
    fields.forEach(f => {
        averages[f] = counts[f] > 0 ? sums[f] / counts[f] : null;
    });

    return {
        ...averages,
        totalRows: rows.length,
        firstTime: rows[0]?.tlocal || null,
        lastTime: rows[rows.length - 1]?.tlocal || null
    };
}

// ============================================
// TIPHON QoS Standards
// ============================================
const TIPHON_LATENCY = [
    { category: 'Sangat Baik', max: 150, index: 4, color: '#10b981' },
    { category: 'Baik', max: 300, index: 3, color: '#3b82f6' },
    { category: 'Sedang', max: 350, index: 2, color: '#f59e0b' },
    { category: 'Buruk', max: Infinity, index: 1, color: '#ef4444' }
];

const TIPHON_PACKETLOSS = [
    { category: 'Sangat Baik', max: 3, index: 4, color: '#10b981', desc: 'Tidak ada degradasi' },
    { category: 'Baik', max: 15, index: 3, color: '#3b82f6', desc: 'Degradasi kecil' },
    { category: 'Sedang', max: 25, index: 2, color: '#f59e0b', desc: 'Degradasi signifikan' },
    { category: 'Buruk', max: Infinity, index: 1, color: '#ef4444', desc: 'Tidak layak' }
];

function getLatencyCategory(latencyMs) {
    for (const tier of TIPHON_LATENCY) {
        if (latencyMs < tier.max || (tier.max === Infinity && latencyMs >= 350)) {
            return tier;
        }
    }
    return TIPHON_LATENCY[3];
}

function getPacketLossCategory(pctLoss) {
    for (const tier of TIPHON_PACKETLOSS) {
        if (pctLoss <= tier.max || (tier.max === Infinity && pctLoss > 25)) {
            return tier;
        }
    }
    return TIPHON_PACKETLOSS[3];
}

// ============================================
// Compute QoS metrics
// ============================================
function computeQoS(rows, hasReceivedAt, sendIntervalSec = 300) {
    const result = {
        latency: null,
        packetLoss: null
    };

    // --- LATENCY ---
    // Requires both tlocal (T_send) and received_at (T_receive)
    if (hasReceivedAt) {
        const delays = [];

        rows.forEach(row => {
            if (row.tlocal && row.received_at) {
                const tSend = new Date(row.tlocal).getTime();
                const tReceive = new Date(row.received_at).getTime();

                if (!isNaN(tSend) && !isNaN(tReceive)) {
                    const delayMs = tReceive - tSend;
                    if (delayMs >= 0) { // valid delay
                        delays.push(delayMs);
                    }
                }
            }
        });

        if (delays.length > 0) {
            const totalDelay = delays.reduce((sum, d) => sum + d, 0);
            const avgLatency = totalDelay / delays.length;
            const minLatency = Math.min(...delays);
            const maxLatency = Math.max(...delays);
            const category = getLatencyCategory(avgLatency);

            // Per-packet latency details
            const perPacket = delays.map((d, i) => ({ packet: i + 1, latencyMs: d }));

            result.latency = {
                avgMs: Math.round(avgLatency * 100) / 100,
                minMs: minLatency,
                maxMs: maxLatency,
                totalPackets: delays.length,
                category: category.category,
                index: category.index,
                color: category.color,
                formula: `Latency = (1/${delays.length}) × Σ(T_receive - T_send) = ${avgLatency.toFixed(2)} ms`,
                perPacket
            };
        }
    }

    // --- PACKET LOSS ---
    // P_send = expected packets based on time range and interval
    // P_receive = actual received packets (rows in data)
    if (rows.length >= 2) {
        // Use min/max to handle both ascending (Local) and descending (Live/unshift) order
        const timestamps = rows
            .map(r => new Date(r.tlocal).getTime())
            .filter(t => !isNaN(t));

        if (timestamps.length >= 2) {
            const firstTime = Math.min(...timestamps);
            const lastTime = Math.max(...timestamps);
            const timeRangeMs = lastTime - firstTime;
            const intervalMs = sendIntervalSec * 1000;

            // Expected packets = (time range / interval) + 1 (inclusive of first)
            const pSend = Math.round(timeRangeMs / intervalMs) + 1;
            const pReceive = rows.length;
            const lost = Math.max(0, pSend - pReceive);
            const packetLossPct = pSend > 0 ? (lost / pSend) * 100 : 0;
            const category = getPacketLossCategory(packetLossPct);

            result.packetLoss = {
                pSend,
                pReceive,
                lost,
                percentage: Math.round(packetLossPct * 100) / 100,
                intervalSec: sendIntervalSec,
                category: category.category,
                index: category.index,
                color: category.color,
                description: category.desc,
                formula: `Packet Loss = ((${pSend} - ${pReceive}) / ${pSend}) × 100% = ${packetLossPct.toFixed(2)}%`
            };
        }
    }

    return result;
}

// ============================================
// POST /api/upload — Upload Excel/CSV file
// ============================================
router.post('/upload', upload.single('file'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ success: false, message: 'No file uploaded' });
        }

        // Get optional interval parameter (default 300 sec = 5 min)
        const interval = parseInt(req.body?.interval) || 300;

        const filePath = req.file.path;
        const { location, rows, hasReceivedAt } = parseFile(filePath);

        if (rows.length === 0) {
            return res.status(400).json({
                success: false,
                message: 'No valid data found in the file. Make sure columns include: waktu/tlocal + sensor columns'
            });
        }

        // Store in memory
        activeData.filename = req.file.originalname;
        activeData.location = location;
        activeData.rows = rows;
        activeData.summary = computeSummary(rows);
        activeData.qos = computeQoS(rows, hasReceivedAt, interval);

        console.log(`📂 File uploaded: ${req.file.originalname} (${rows.length} rows, hasReceivedAt: ${hasReceivedAt})`);

        res.json({
            success: true,
            message: `Successfully loaded ${rows.length} data rows`,
            filename: req.file.originalname,
            location,
            totalRows: rows.length,
            hasReceivedAt,
            summary: activeData.summary,
            qos: activeData.qos
        });

    } catch (error) {
        console.error('❌ Error processing file:', error.message);
        res.status(500).json({
            success: false,
            message: 'Failed to process file: ' + error.message
        });
    }
});

// ============================================
// GET /api/summary
// ============================================
router.get('/summary', (req, res) => {
    res.json({
        success: true,
        filename: activeData.filename,
        location: activeData.location,
        summary: activeData.summary
    });
});

// ============================================
// GET /api/qos — Get QoS analysis results
// ============================================
router.get('/qos', (req, res) => {
    res.json({
        success: true,
        filename: activeData.filename,
        qos: activeData.qos
    });
});

// ============================================
// POST /api/qos/recalculate — Recalculate with different interval
// ============================================
router.post('/qos/recalculate', express.json(), (req, res) => {
    const interval = parseInt(req.body?.interval) || 300;

    if (activeData.rows.length === 0) {
        return res.status(400).json({ success: false, message: 'No data loaded' });
    }

    const hasReceivedAt = activeData.rows.some(r => r.received_at);
    activeData.qos = computeQoS(activeData.rows, hasReceivedAt, interval);

    res.json({
        success: true,
        qos: activeData.qos
    });
});

// ============================================
// GET /api/data — Get paginated data rows
// ============================================
router.get('/data', (req, res) => {
    const page = parseInt(req.query.page) || 1;
    const limit = parseInt(req.query.limit) || 20;

    const totalRows = activeData.rows.length;
    const totalPages = Math.ceil(totalRows / limit);
    const startIdx = (page - 1) * limit;
    const endIdx = Math.min(startIdx + limit, totalRows);
    const pageData = activeData.rows.slice(startIdx, endIdx);

    res.json({
        success: true,
        filename: activeData.filename,
        data: pageData,
        pagination: {
            page,
            limit,
            totalRows,
            totalPages,
            hasNext: page < totalPages,
            hasPrev: page > 1
        }
    });
});

// ============================================
// GET /api/export
// ============================================
router.get('/export', (req, res) => {
    if (!activeData.rows.length) {
        return res.status(404).json({ success: false, message: 'No data to export' });
    }

    const hasReceivedAt = activeData.rows.some(r => r.received_at);
    // Add check for coordinates
    const hasLocation = activeData.rows.some(r => r.latitude !== undefined && r.latitude !== null);
    
    let headers = 'Waktu Kirim';
    if (hasReceivedAt) headers += ',Waktu Terima';
    if (hasLocation) headers += ',Latitude,Longitude';
    headers += ',Tekanan Udara (mBar),Suhu (°C),Kelembapan (%),Kecepatan Angin (m/s),Arah Angin (°)\n';

    // Helper: convert UNIX timestamp to datetime string if needed
    function toDateStr(val) {
        if (!val) return '';
        const s = String(val).trim();
        // If it's a UNIX timestamp string (9-13 digits), convert to readable
        if (/^\d{9,13}$/.test(s)) {
            const unixVal = parseInt(s);
            const isMs = unixVal > 1000000000000;
            const d = new Date(isMs ? unixVal : unixVal * 1000);
            if (!isNaN(d.getTime())) {
                return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')} ${String(d.getHours()).padStart(2,'0')}:${String(d.getMinutes()).padStart(2,'0')}:${String(d.getSeconds()).padStart(2,'0')}.${String(d.getMilliseconds()).padStart(3,'0')}`;
            }
        }
        return s;
    }

    const csvData = activeData.rows.map(row => {
        let line = `${toDateStr(row.tlocal)}`;
        if (hasReceivedAt) line += `,${toDateStr(row.received_at)}`;
        if (hasLocation) line += `,${row.latitude ?? ''},${row.longitude ?? ''}`;
        line += `,${row.airpressure ?? ''},${row.airtemperature ?? ''},${row.airhumidity ?? ''},${row.windspeed ?? ''},${row.winddirection ?? ''}`;
        return line;
    }).join('\n');

    let safeName = (activeData.filename || 'data').replace(/\.[^.]+$/, '');
    // Prevent looping export_export_ prefix
    safeName = safeName.replace(/^export_/i, '');
    
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename=export_${safeName}.csv`);
    res.send('\uFEFF' + headers + csvData);
});

// ============================================
// DELETE /api/data
// ============================================
router.delete('/data', (req, res) => {
    activeData = { filename: null, location: null, rows: [], summary: null, qos: null };
    console.log('🗑️ Data cleared');
    res.json({ success: true, message: 'Data cleared' });
});

// ============================================
// POST /api/terima_data — Receive live data from Teltonika router
// ============================================
router.post('/terima_data', (req, res) => {
    const data = req.body;
    const receivedTime = Date.now(); // Generate server reception time (milliseconds)

    // 1. Initialize memory if empty
    if (!activeData.filename) {
        activeData.filename = 'Live_Monitoring.csv';
        activeData.location = 'Stasiun Cuaca Real-Time';
    }

    // CEK PAKET REDUNDAN (Duplikasi Data)
    // Jika tlocal paket baru sama persis dengan tlocal paket terakhir yang diterima, abaikan.
    if (activeData.rows.length > 0 && activeData.rows[0].tlocal === data.tlocal) {
        console.log(`[WARNING] Paket redundan terdeteksi (tlocal: ${data.tlocal}). Diabaikan agar tidak merusak QoS.`);
        return res.status(200).json({ success: true, message: 'Paket redundan diabaikan.' });
    }

    // 2. Store in RAM (unshift = newest data on top)
    activeData.rows.unshift({
        tlocal: data.tlocal, // String time from router
        received_at: receivedTime, // Unix Epoch 13 digit
        latitude: data.latitude,
        longitude: data.longitude,
        airpressure: data.airpressure != null ? parseFloat(data.airpressure) * 0.1 : null,
        airtemperature: data.airtemperature != null ? parseFloat(data.airtemperature) * 0.1 : null,
        airhumidity: data.airhumidity != null ? parseFloat(data.airhumidity) * 0.1 : null,
        windspeed: data.windspeed != null ? parseFloat(data.windspeed) * 0.01 : null,
        winddirection: data.winddirection != null ? parseFloat(data.winddirection) * 22.5 : null
    });

    // Limit memory to 5000 rows max
    if (activeData.rows.length > 5000) activeData.rows.pop();

    // 3. Auto-update summary so dashboard cards refresh
    activeData.summary = computeSummary(activeData.rows);

    // 4. Update location from coordinates if available
    if (data.latitude && data.longitude && activeData.location === 'Stasiun Cuaca Real-Time') {
        activeData.location = `Stasiun Cuaca Real-Time (${data.latitude}, ${data.longitude})`;
    }

    // 5. Update QoS Live automatically
    const hasReceivedAt = activeData.rows.some(r => r.received_at);
    activeData.qos = computeQoS(activeData.rows.slice(0, 100), hasReceivedAt, 300);

    // 6. Backup to CSV file on server
    const backupPath = path.join(__dirname, '../uploads/Backup_Live.csv');

    if (!fs.existsSync(backupPath)) {
        const header = "tlocal,received_at,latitude,longitude,airpressure,airtemperature,airhumidity,windspeed,winddirection\n";
        fs.writeFileSync(backupPath, header);
    }

    const backupLine = `${data.tlocal},${receivedTime},${data.latitude || ''},${data.longitude || ''},${data.airpressure || ''},${data.airtemperature || ''},${data.airhumidity || ''},${data.windspeed || ''},${data.winddirection || ''}\n`;
    fs.appendFileSync(backupPath, backupLine);

    console.log(`[LIVE] Data masuk pada ${data.tlocal}: Suhu ${data.airtemperature}°C, Angin ${data.windspeed} m/s | Latency: ${activeData.qos?.latency?.avgMs || '--'}ms`);
    res.status(200).json({ success: true, message: 'Data berhasil diterima & dibackup!' });
});

module.exports = router;
