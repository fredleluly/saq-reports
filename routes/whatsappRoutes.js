const { Router } = require('express');
const { reportRepo, assistantRepo, globalConfigRepo } = require('../repositories');
const { getStatus, getQRCode, restartClient, getChats, getClient } = require('../services/whatsapp/whatsappClient');
const { generateRemindMessage } = require('../services/messageGeneratorService');
const { asyncHandler } = require('./_helpers');

const router = Router();

// Halaman daftar chat/grup
router.get('/chats', asyncHandler(async (req, res) => {
    const chats = await getChats();
    const assistants = assistantRepo.read();
    
    res.render('whatsapp-chats', {
        title:      'WhatsApp Chats & Groups',
        status:     getStatus(),
        chats:      chats,
        assistants: assistants,
        activePage: 'whatsapp',
    });
}));

// Halaman scan QR Code
router.get('/qr', (req, res) => {
    res.render('whatsapp-qr', {
        title:      'WhatsApp Login - QR Code',
        qrCode:     getQRCode(),
        status:     getStatus(),
        activePage: 'whatsapp',
    });
});

// API endpoint: status real-time (untuk polling di frontend)
router.get('/status', (req, res) => {
    res.json({
        status: getStatus(),
        qrCode: getQRCode(),
        hasQR:  !!getQRCode(),
    });
});

// Restart client
router.post('/restart', asyncHandler(async (req, res) => {
    await restartClient();
    res.json({ success: true, message: 'WhatsApp client restarting...' });
}));

// Dashboard manajemen WhatsApp
router.get('/', asyncHandler(async (req, res) => {
    const reports         = reportRepo.read();
    const whatsappReports = reports.filter(r => r.author?.includes('@'));
    const today           = new Date().toDateString();
    const todayReports    = whatsappReports.filter(r =>
        new Date(r.date).toDateString() === today
    );
    
    res.render('whatsapp-dashboard', {
        title:               'WhatsApp Bot Management',
        status:              getStatus(),
        totalWhatsAppReports: whatsappReports.length,
        todayReports:        todayReports.length,
        recentReports:       whatsappReports.slice(-10).reverse(),
        activePage:          'whatsapp',
    });
}));

// API: Send WA Reminder to PJ in group
router.post('/api/remind-pj', asyncHandler(async (req, res) => {
    const { pjName, pjPhone, courseName, courseCode, className, day, time, date } = req.body;
    
    // Ambil group_id dari config_global.json (dinamis, tidak hardcoded)
    const globalConfig = globalConfigRepo.read();
    const GROUP_ID = globalConfig?.system?.group_id;
    
    if (!GROUP_ID) {
        return res.status(500).json({
            success: false,
            message: 'group_id belum dikonfigurasi di config_global.json'
        });
    }
    
    const client = getClient();
    if (!client || getStatus() !== 'ready') {
        return res.status(503).json({
            success: false,
            message: 'WhatsApp bot belum terhubung. Silakan scan QR terlebih dahulu.'
        });
    }
    
    if (!pjPhone) {
        return res.status(400).json({
            success: false,
            message: 'Nomor telepon PJ tidak ditemukan.'
        });
    }
    
    // Format nomor WA: 08xxx → 628xxx
    const waNumber = pjPhone.startsWith('0')
    ? '62' + pjPhone.slice(1)
    : pjPhone;
    const mention  = `@${waNumber}`;
    
    // Generate pesan dari Gemini dengan fallback ke template default
    const message = await generateRemindMessage(pjName, pjPhone, courseName, courseCode, className, day, time, date);
    
    try {
        // Kirim ke grup dengan mention
        const chat = await client.getChatById(GROUP_ID);
        
        // Jika pesan dari Gemini belum termasuk mention, tambahkan di awal
        let finalMessage = message;
        if (!message.includes('@')) {
            const mention = `@${waNumber}`;
            finalMessage = message.startsWith('📋') 
            ? message 
            : `${mention} ${message}`;
        }
        
        await chat.sendMessage(finalMessage, {
            mentions: [`${waNumber}@c.us`]
        });
        
        console.log(`[WA-REMIND] Berhasil kirim reminder ke grup untuk PJ: ${pjName} (${pjPhone})`);
        return res.json({
            success: true,
            message: `Reminder berhasil dikirim ke grup untuk @${pjName}`
        });
    } catch (err) {
        console.error('[WA-REMIND] Gagal kirim pesan:', err);
        return res.status(500).json({
            success: false,
            message: `Gagal kirim pesan: ${err.message}`
        });
    }
}));

module.exports = router;
