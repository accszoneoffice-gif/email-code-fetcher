const express = require('express');
const axios = require('axios');
const path = require('path');
const Imap = require('imap');
const { simpleParser } = require('mailparser');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Helper function to extract exact OTP code, clean verification link, and full mail body
function extractCodeAndLink(text, htmlContent) {
    let foundCode = null;
    let actionLink = null;

    // 1. OTP Code Match (4-8 digits)
    const codeMatch = text.match(/\b\d{4,8}\b/);
    if (codeMatch) foundCode = codeMatch[0];

    // 2. Extract Verification / Action Links from HTML
    if (htmlContent) {
        const hrefRegex = /href=["'](https?:\/\/[^"']+)["']/gi;
        let match;
        const extractedLinks = [];

        while ((match = hrefRegex.exec(htmlContent)) !== null) {
            extractedLinks.push(match[1]);
        }

        // Target primary verification/activation/confirm links (Reddit, Discord, social sites, etc.)
        actionLink = extractedLinks.find(link => {
            const clean = link.toLowerCase();
            const isIgnored = clean.endsWith('.png') || 
                              clean.endsWith('.jpg') || 
                              clean.endsWith('.jpeg') || 
                              clean.endsWith('.gif') || 
                              clean.endsWith('.css') || 
                              clean.includes('schemas.microsoft') || 
                              clean.includes('w3.org') || 
                              clean.includes('schema.org') || 
                              clean.includes('unsubscribe') || 
                              clean.includes('faq') || 
                              clean.includes('contact');
            return !isIgnored;
        }) || null;
    }

    // 3. Fallback: Plain Text Link Search
    if (!actionLink) {
        const urlRegex = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/gi;
        const foundUrls = text.match(urlRegex) || [];
        actionLink = foundUrls.find(link => {
            const clean = link.toLowerCase();
            return !clean.endsWith('.png') && 
                   !clean.endsWith('.jpg') && 
                   !clean.includes('schemas.microsoft') &&
                   !clean.includes('unsubscribe');
        }) || null;
    }

    if (actionLink) {
        actionLink = actionLink.replace(/[.,;)]+$/, '');
    }

    return { foundCode, actionLink };
}

// IMAP Fetcher for Gmail and AT&T
function fetchViaImap(host, email, password) {
    return new Promise((resolve, reject) => {
        const imap = new Imap({
            user: email,
            password: password,
            host: host,
            port: 993,
            tls: true,
            tlsOptions: { rejectUnauthorized: false }
        });

        imap.once('ready', () => {
            imap.openBox('INBOX', true, (err, box) => {
                if (err) {
                    imap.end();
                    return reject(err);
                }

                imap.search(['ALL'], (searchErr, results) => {
                    if (searchErr || !results || !results.length) {
                        imap.end();
                        return resolve({ success: false, error: 'No emails found.' });
                    }

                    const recent = results.slice(-1); // Get latest email
                    const fetch = imap.fetch(recent, { bodies: '' });

                    fetch.on('message', (msg) => {
                        msg.on('body', (stream) => {
                            simpleParser(stream, async (err, parsed) => {
                                const text = parsed.text || '';
                                const html = parsed.html || parsed.textAsHtml || '';
                                const extracted = extractCodeAndLink(`${parsed.subject} ${text}`, html);

                                imap.end();
                                resolve({
                                    success: true,
                                    subject: parsed.subject,
                                    code: extracted.foundCode,
                                    link: extracted.actionLink,
                                    fullHtml: html || text.replace(/\n/g, '<br>')
                                });
                            });
                        });
                    });

                    fetch.once('error', (fErr) => {
                        imap.end();
                        reject(fErr);
                    });
                });
            });
        });

        imap.once('error', (err) => {
            reject(err);
        });

        imap.connect();
    });
}

app.post('/api/get-code', async (req, res) => {
    try {
        const { provider, accountData } = req.body;
        if (!accountData) {
            return res.status(400).json({ success: false, error: 'Account data is required.' });
        }

        // 1. Gmail IMAP
        if (provider === 'gmail') {
            const parts = accountData.split('|');
            if (parts.length < 2) {
                return res.status(400).json({ success: false, error: 'Format must be: email|app_password' });
            }
            const [email, appPassword] = parts.map(p => p.trim());
            const result = await fetchViaImap('imap.gmail.com', email, appPassword);
            return res.json(result);
        }

        // 2. AT&T IMAP
        if (provider === 'att') {
            const parts = accountData.split('|');
            if (parts.length < 2) {
                return res.status(400).json({ success: false, error: 'Format must be: email|app_password' });
            }
            const [email, appPassword] = parts.map(p => p.trim());
            const result = await fetchViaImap('imap.mail.att.net', email, appPassword);
            return res.json(result);
        }

        // 3. Outlook / Hotmail (Microsoft Graph API)
        const parts = accountData.split('|');
        if (parts.length < 4) {
            return res.status(400).json({ success: false, error: 'Invalid format. Use email|password|refresh_token|client_id' });
        }

        const [email, password, refreshToken, clientId] = parts.map(p => p.trim());

        const tokenParams = new URLSearchParams({
            client_id: clientId,
            grant_type: 'refresh_token',
            refresh_token: refreshToken,
            scope: 'https://graph.microsoft.com/Mail.Read'
        });

        const tokenRes = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', tokenParams.toString(), {
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        });

        const accessToken = tokenRes.data.access_token;

        const mailRes = await axios.get('https://graph.microsoft.com/v1.0/me/messages?$top=1&$select=subject,bodyPreview,body', {
            headers: { Authorization: `Bearer ${accessToken}` }
        });

        const messages = mailRes.data.value;
        if (!messages || messages.length === 0) {
            return res.json({ success: false, error: 'No emails found.' });
        }

        const msg = messages[0];
        const htmlBody = msg.body?.content || '';
        const textBody = msg.bodyPreview || '';
        const extracted = extractCodeAndLink(`${msg.subject} ${textBody}`, htmlBody);

        return res.json({ 
            success: true, 
            subject: msg.subject,
            code: extracted.foundCode, 
            link: extracted.actionLink,
            fullHtml: htmlBody || textBody.replace(/\n/g, '<br>')
        });

    } catch (err) {
        return res.status(500).json({
            success: false,
            error: err.response ? JSON.stringify(err.response.data) : err.message
        });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));