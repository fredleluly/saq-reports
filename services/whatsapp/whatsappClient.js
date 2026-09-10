const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode  = require('qrcode-terminal');
const QRCode  = require('qrcode');

const { isTargetClosedError } = require('./whatsappHelpers');
const { handleIncomingMessage } = require('./messageHandler');

// ─── State ─────────────────────────────────────────────────────────────────────

const state = {
    client:             null,
    currentQRCode:      null,
    status:             'disconnected',
    reconnectTimer:     null,
    reconnectAttempts:  0,
    isInitializing:     false,
    isRestarting:       false,
};

const MAX_RECONNECT_DELAY_MS = 60_000;

// ─── Getters (ekspor state buat routes) ───────────────────────────────────────

const getStatus   = () => state.status;
const getQRCode   = () => state.currentQRCode;
const getClient   = () => state.client;

// ─── Internal Helpers ──────────────────────────────────────────────────────────

function clearReconnectTimer() {
    if (state.reconnectTimer) {
        clearTimeout(state.reconnectTimer);
        state.reconnectTimer = null;
    }
}

function scheduleReconnect(trigger) {
    if (state.reconnectTimer || state.isInitializing) return;
    
    state.reconnectAttempts += 1;
    const delay = Math.min(
        2000 * Math.pow(2, state.reconnectAttempts - 1),
        MAX_RECONNECT_DELAY_MS
    );
    state.status = 'reconnecting';
    
    console.warn(`[WHATSAPP] Menjadwalkan reconnect (${trigger}) dalam ${delay}ms (percobaan ke-${state.reconnectAttempts})`);
    
    state.reconnectTimer = setTimeout(() => {
        state.reconnectTimer = null;
        initializeClient({ reason: `reconnect_${trigger}` });
    }, delay);
}

async function destroyClientSafe() {
    if (!state.client) return;
    try {
        await state.client.destroy();
    } catch (err) {
        if (!isTargetClosedError(err)) {
            console.error('[WHATSAPP] Gagal destroy client:', err);
        }
    } finally {
        state.client        = null;
        state.currentQRCode = null;
    }
}

// ─── Initialization ────────────────────────────────────────────────────────────

function initializeClient(options = {}) {
    if (state.isInitializing) {
        console.log('[WHATSAPP] Sudah sedang inisialisasi, permintaan baru diabaikan.');
        return;
    }
    
    state.isInitializing = true;
    clearReconnectTimer();
    state.status = 'initializing';
    console.log(`[WHATSAPP] Memulai inisialisasi (${options.reason || 'startup'})...`);
    
    const client = new Client({
        authStrategy: new LocalAuth(),
        puppeteer: {
            executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
            headless: true,
            protocolTimeout: 300000, // Increase protocol timeout to 5 minutes to prevent Runtime.callFunctionOn timeouts
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--no-zygote',
                '--disable-gpu',
            ],
        },
    });
    
    state.client = client;
    
    // ── Event: QR ──────────────────────────────────────────────────────────────
    client.on('qr', async (qr) => {
        qrcode.generate(qr, { small: true });
        try {
            state.currentQRCode = await QRCode.toDataURL(qr);
            state.status        = 'qr_ready';
            console.log('[WHATSAPP] QR Code siap di /whatsapp/qr');
        } catch (err) {
            console.error('[WHATSAPP] Gagal generate QR untuk web:', err);
        }
    });
    
    // ── Event: Ready ───────────────────────────────────────────────────────────
    client.on('ready', () => {
        state.status            = 'ready';
        state.currentQRCode     = null;
        state.reconnectAttempts = 0;
        console.log('[WHATSAPP] ✅ Client siap. Bot terhubung.');
    });
    
    // ── Event: Authenticated ───────────────────────────────────────────────────
    client.on('authenticated', () => {
        state.status = 'authenticated';
        console.log('[WHATSAPP] ✅ Authenticated.');
    });
    
    // ── Event: Auth Failure ────────────────────────────────────────────────────
    client.on('auth_failure', (msg) => {
        state.status        = 'auth_failed';
        state.currentQRCode = null;
        console.error('[WHATSAPP] ❌ Auth gagal:', msg);
        scheduleReconnect('auth_failure');
    });
    
    // ── Event: Disconnected ────────────────────────────────────────────────────
    client.on('disconnected', (reason) => {
        state.status        = 'disconnected';
        state.currentQRCode = null;
        console.log('[WHATSAPP] ❌ Disconnected:', reason);
        
        if (!state.isRestarting) {
            scheduleReconnect('disconnected');
        }
    });
    
    // ── Event: Message ─────────────────────────────────────────────────────────
    client.on('message_create', async (message) => {
        await handleIncomingMessage(message, state.client);
    });
    
    // ── Initialize ─────────────────────────────────────────────────────────────
    client.initialize()
    .then(() => console.log('[WHATSAPP] initialize() berhasil dipanggil.'))
    .catch((err) => {
        if (isTargetClosedError(err)) {
            console.warn('[WHATSAPP] Browser target tertutup saat initialize(). Akan reconnect.');
        } else {
            console.error('[WHATSAPP] Gagal initialize():', err);
        }
        scheduleReconnect('initialize_failed');
    })
    .finally(() => {
        state.isInitializing = false;
        state.isRestarting   = false;
    });
}

// ─── Restart (untuk route /whatsapp/restart) ──────────────────────────────────

async function restartClient() {
    state.isRestarting  = true;
    state.status        = 'restarting';
    state.currentQRCode = null;
    clearReconnectTimer();
    await destroyClientSafe();
    initializeClient({ reason: 'manual_restart' });
}

const getChats    = async () => {
    if (!state.client || state.status !== 'ready') return [];
    try {
        const chats = await state.client.getChats();
        return chats.map(chat => ({
            id: chat.id._serialized,
            name: chat.name,
            isGroup: chat.isGroup,
            unreadCount: chat.unreadCount,
            timestamp: chat.timestamp,
        }));
    } catch (err) {
        console.error('[WHATSAPP] Gagal mengambil daftar chat:', err);
        return [];
    }
};

module.exports = {
    initializeClient,
    restartClient,
    getStatus,
    getQRCode,
    getClient,
    getChats,
};
