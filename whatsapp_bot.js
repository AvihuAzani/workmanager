const { Client, LocalAuth, List } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const axios = require('axios');
const fs = require('fs');
const https = require('https');
const { execFile } = require('child_process');
const path = require('path');

// ============================================================
// הגדרות
// ============================================================
const PHONE_NUMBER    = '0526845629';
const EMPLOYEE_ID     = '104427';
const CLIENT_ID       = '354193a2-8d29-11ea-bc55-0242ac130004';
const CLIENT_SECRET   = '354193a2-8d29-11ea-bc55-0242ac130003';
const PHONE_DEVICE_ID = 'FF210C6E-5313-4961-846D-229DF3FAC0FC';
const DEVICE_MODEL    = 'ios_Apple_iPhone 16_SysVer_26.3.1_appVer_23.03.26.1P';
const TOKEN_FILE      = 'C:\\Users\\Avihu\\Documents\\cellcom_reports\\token.txt';
const SCRIPTS_DIR     = 'C:\\Users\\Avihu\\Documents\\cellcom_reports';
const PYTHON          = 'python';

const httpsAgent = new https.Agent({ rejectUnauthorized: false });

const BASE_HEADERS = {
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'he-IL,he;q=0.9',
    'Content-Type': 'application/json',
    'format': 'application/json',
    'phonedeviceid': PHONE_DEVICE_ID,
    'devicemodel': DEVICE_MODEL,
    'User-Agent': 'HomeTechAppClient/23.03.26 CFNetwork/3860.400.51 Darwin/25.3.0',
};

// ============================================================
// State
// ============================================================
const state = {};  // keyed by chatId

function getState(chatId) {
    if (!state[chatId]) state[chatId] = { step: 'idle', ticketId: null };
    return state[chatId];
}

// ============================================================
// Cellcom Auth
// ============================================================
async function loginStep1() {
    const res = await axios.post(
        'https://tech-api.cellcom.co.il/api/technician/loginStep1',
        {
            DeviceModel: DEVICE_MODEL, PhoneNumber: PHONE_NUMBER,
            EmployeeId: EMPLOYEE_ID, ClientId: CLIENT_ID,
            LoginType: 'EMPLOYEEPHONE', Scope: 'USERNAME',
            SId: 'ios_Apple_iPhone 16_SysVer_26.3.1',
            PhoneDeviceId: PHONE_DEVICE_ID,
        },
        { headers: { ...BASE_HEADERS, Authorization: 'default' }, httpsAgent }
    );
    const body = res.data?.Body;
    if (!body?.isSuccess) throw new Error('שגיאה בשליחת SMS');
    return body.ticketId;
}

async function loginStep2(otp, ticketId) {
    const res = await axios.post(
        'https://tech-api.cellcom.co.il/api/technician/loginStep2',
        {
            PhoneNumber: PHONE_NUMBER, EmployeeId: EMPLOYEE_ID,
            ClientId: CLIENT_ID, ClientSecret: CLIENT_SECRET,
            LoginType: 'EMPLOYEEPHONE', Scope: 'USERNAME',
            SId: 'ios_Apple_iPhone 16_SysVer_26.3.1',
            OtpCode: otp, OtpGuid: ticketId,
            PhoneDeviceId: PHONE_DEVICE_ID,
        },
        { headers: BASE_HEADERS, httpsAgent }
    );
    const rc = res.data?.Header?.ReturnCode;
    if (rc !== '0') throw new Error(res.data?.Header?.ReturnCodeMessage || 'קוד שגוי');
    return res.data.Body.access_token;
}

// ============================================================
// הרצת סקריפט Python
// ============================================================
function runScript(scriptName, args = []) {
    return new Promise((resolve) => {
        const scriptPath = path.join(SCRIPTS_DIR, scriptName);
        execFile(PYTHON, [scriptPath, ...args], { encoding: 'utf8', timeout: 300000 }, (err, stdout, stderr) => {
            resolve((stdout || stderr || err?.message || '').trim());
        });
    });
}

// ============================================================
// תפריט ראשי
// ============================================================
function buildMainMenu() {
    return new List(
        'ברוך הבא לבוט סלקום 🔧\nבחר פעולה מהרשימה:',
        '☰  לחץ לבחירה',
        [{
            title: 'פעולות',
            rows: [
                { id: 'get_token',   title: '🔑 קבל טוקן',         description: 'התחברות מחדש לסלקום' },
                { id: 'get_pekaot',  title: '📋 בנק פקעות',         description: 'שליפת כל הפקעות' },
                { id: 'get_malai',   title: '📦 מלאי',              description: 'שליפת מלאי ציוד' },
                { id: 'get_history', title: '📜 היסטוריה',           description: 'היסטוריית קריאות' },
            ]
        }],
        'סלקום בוט',
        'שלח "תפריט" בכל עת לחזרה'
    );
}

// ============================================================
// WhatsApp Client
// ============================================================
const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: {
        headless: false,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
    }
});

client.on('qr', qr => {
    console.log('סרוק את ה-QR:');
    qrcode.generate(qr, { small: true });
});

client.on('ready', () => console.log('✅ הבוט מחובר ומוכן'));

client.on('message', async msg => {
    if (!msg.fromMe) return;

    const chatId = msg.from;
    const s = getState(chatId);
    const text = msg.body.trim();

    // תפריט ראשי
    if (text === 'תפריט' || text === 'menu' || text === 'התחל') {
        s.step = 'idle';
        await client.sendMessage(chatId, buildMainMenu());
        return;
    }

    // בחירה מהתפריט
    if (msg.type === 'list_response' || s.step === 'idle') {
        const rowId = msg.selectedRowId || text;

        if (rowId === 'get_token' || text === '🔑 קבל טוקן') {
            try {
                await msg.reply('⏳ שולח SMS...');
                const ticketId = await loginStep1();
                s.ticketId = ticketId;
                s.step = 'awaiting_otp';
                await msg.reply('📱 *נשלח קוד SMS*\nשלח לי את הקוד כאן:');
            } catch (e) {
                await msg.reply(`❌ שגיאה: ${e.message}`);
            }
            return;
        }

        if (rowId === 'get_pekaot' || text === '📋 בנק פקעות') {
            await msg.reply('⏳ שולף פקעות, אנא המתן...');
            const output = await runScript('cellcom_bank_pekaot_3.py');
            await msg.reply(`📊 ${output}\n\nשלח "תפריט" לחזרה`);
            return;
        }

        if (rowId === 'get_malai' || text === '📦 מלאי') {
            await msg.reply('⏳ שולף מלאי...');
            const output = await runScript('cellcom_malai.py');
            await msg.reply(`📦 ${output}\n\nשלח "תפריט" לחזרה`);
            return;
        }

        if (rowId === 'get_history' || text === '📜 היסטוריה') {
            await msg.reply('⏳ שולף היסטוריה...');
            const output = await runScript('cellcom_history.py');
            await msg.reply(`📜 ${output}\n\nשלח "תפריט" לחזרה`);
            return;
        }
    }

    // קבלת OTP
    if (s.step === 'awaiting_otp' && /^\d{4,8}$/.test(text)) {
        try {
            await msg.reply('⏳ מאמת...');
            const token = await loginStep2(text, s.ticketId);
            fs.writeFileSync(TOKEN_FILE, token, 'utf8');
            s.step = 'idle';
            s.ticketId = null;
            await msg.reply('✅ *טוקן נשמר בהצלחה!*\n\nשלח "תפריט" למשך נתונים.');
        } catch (e) {
            await msg.reply(`❌ ${e.message}\nנסה שוב או שלח "טוקן" לקוד חדש.`);
        }
        return;
    }

    // הודעה לא מזוהה — הצג תפריט
    await client.sendMessage(chatId, buildMainMenu());
});

client.initialize();
