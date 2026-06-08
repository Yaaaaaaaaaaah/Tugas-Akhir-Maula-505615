const express = require('express');
const cors = require('cors');
const path = require('path');

const dataRoutes = require('./routes/sensor');

const app = express();
const PORT = 1212;
const HOST = '0.0.0.0'; // Listen on all interfaces for external access

// ============================================
// Middleware
// ============================================
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve static files from frontend folder
app.use(express.static(path.join(__dirname, '../frontend')));

// ============================================
// API Routes
// ============================================
app.use('/api', dataRoutes);

// ============================================
// Root route - serve dashboard
// ============================================
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '../frontend/index.html'));
});

// ============================================
// Error handling middleware
// ============================================
app.use((err, req, res, next) => {
    console.error('❌ Server error:', err.message);
    res.status(500).json({
        success: false,
        message: 'Internal server error'
    });
});

// ============================================
// Start server
// ============================================
app.listen(PORT, HOST, () => {
    console.log('');
    console.log('╔══════════════════════════════════════════════╗');
    console.log('║   📊 Data Visualization Tool Started         ║');
    console.log('╠══════════════════════════════════════════════╣');
    console.log(`║   🌐 Local:  http://localhost:${PORT}          ║`);
    console.log(`║   🌐 Network: http://16.29.10.8:${PORT}       ║`);
    console.log('╚══════════════════════════════════════════════╝');
    console.log('');
});
