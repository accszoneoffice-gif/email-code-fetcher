const express = require('express');
const axios = require('axios');
const path = require('path');
const Imap = require('imap');
const { simpleParser } = require('mailparser');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Helper function to extract exact OTP code and clean full links
function extractCodeAndLink(text, htmlContent) {
    let foundCode = null;
    let actionLink = null;

    // 1. OTP Code Match (4-8 digits)
    const codeMatch = text.match(/\b\d{4,8}\b/);
    if (codeMatch) foundCode = codeMatch[0];

    // 2. HTML href Tag Parsing (Extracts actual hidden activation/verification URLs)
    if (htmlContent) {
        const hrefRegex = /href=["'](https?:\/\/[^"']+)["']/gi;
        let match;
        const extractedLinks = [];

        while ((match = hrefRegex.exec(htmlContent)) !== null) {
            extractedLinks.push(match[1]);
        }

        // Filter out images, schemas, and unsubscribe links
        actionLink = extractedLinks.find(link => {
            const clean = link.toLowerCase();
            return !clean.endsWith('.png') && 
                   !clean.endsWith('.jpg') && 
                   !clean.endsWith('.jpeg') &&
                   !clean.endsWith('.gif') &&
                   !clean.endsWith('.css') && 
                   !clean.includes('schemas.microsoft') && 
                   !clean.includes('w3.org') &&
                   !clean.includes('schema.org') &&
                   !clean.includes('unsubscribe');
        }) || null;
    }

    // 3. Fallback: Plain Text URL Extraction if HTML link not found
    if (!actionLink) {
        const urlRegex = /https?:\/\/[^\s<>"{}|\\^`\[\]]+/gi;
        const foundUrls = text.match(urlRegex) || [];
        actionLink = foundUrls.find(link => {
            const clean = link.toLowerCase();
            return !clean.endsWith('.png') && 
                   !clean.endsWith('.jpg') && 
                   !clean.includes('schemas.microsoft');
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

                    const recent = results.slice(-3);
                    const fetch = imap.fetch(recent, { bodies: '' });
                    let parsedCount = 0;
                    let foundData = { foundCode: null, actionLink: null };

                    fetch.on('message', (msg) => {
                        msg.on('body', (stream) => {
                            simpleParser(stream, async (err, parsed) => {
                                parsedCount++;
                                const text = parsed.text || '';
                                const html = parsed.html || '';
                                const extracted = extractCodeAndLink(`${parsed.subject} ${text}`, html);

                                if (extracted.foundCode && !foundData.foundCode) foundData.foundCode = extracted.foundCode;
                                if (extracted.actionLink && !foundData.actionLink) foundData.actionLink = extracted.actionLink;

                                if (parsedCount === recent.length) {
                                    imap.end();
                                    resolve({
                                        success: true,
                                        code: foundData.foundCode,
                                        link: foundData.actionLink
                                    });
                                }
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

        const mailRes = await axios.get('https://graph.microsoft.com/v1.0/me/messages?$top=5&$select=subject,bodyPreview,body', {
            headers: { Authorization: `Bearer ${accessToken}` }
        });

        const messages = mailRes.data.value;
        if (!messages || messages.length === 0) {
            return res.json({ success: false, error: 'No emails found.' });
        }

        let foundCode = null;
        let actionLink = null;

        for (const msg of messages) {
            const htmlBody = msg.body?.content || '';
            const textBody = msg.bodyPreview || '';
            const extracted = extractCodeAndLink(`${msg.subject} ${textBody}`, htmlBody);

            if (extracted.foundCode && !foundCode) foundCode = extracted.foundCode;
            if (extracted.actionLink && !actionLink) actionLink = extracted.actionLink;
            if (foundCode && actionLink) break;
        }

        return res.json({ success: true, code: foundCode, link: actionLink });

    } catch (err) {
        return res.status(500).json({
            success: false,
            error: err.response ? JSON.stringify(err.response.data) : err.message
        });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));