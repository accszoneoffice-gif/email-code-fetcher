const express = require('express');
const cors = require('cors');
const axios = require('axios');
const Imap = require('node-imap');
const { simpleParser } = require('mailparser');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// --- HELPER FUNCTION TO EXTRACT CODE & LINK ---
function parseVerificationDetails(subject = '', html = '', text = '') {
    let extractedCode = null;
    let extractedLink = null;

    // 1. Priority check: Extract 6-digit code directly from email Subject
    const subjectCodeMatch = subject.match(/\b\d{6}\b/);
    if (subjectCodeMatch) {
        extractedCode = subjectCodeMatch[0];
    } else {
        // Fallback: Extract 4-8 digit code from body (Text/HTML)
        const bodyCodeMatch = text.match(/\b\d{4,8}\b/) || html.match(/\b\d{4,8}\b/);
        extractedCode = bodyCodeMatch ? bodyCodeMatch[0] : null;
    }

    // 2. Extract Direct Verification Link
    const linkMatch = html.match(/href=["'](https?:\/\/[^"']+)["']/i) || text.match(/(https?:\/\/[^\s]+)/i);
    extractedLink = linkMatch ? linkMatch[1] : null;

    return { code: extractedCode, link: extractedLink };
}

// --- EMAIL FETCHER API ---
app.post('/api/get-code', async (req, res) => {
    const { provider, accountData } = req.body;

    if (!accountData) {
        return res.status(400).json({ success: false, error: 'Account data is required.' });
    }

    try {
        if (provider === 'outlook') {
            const parts = accountData.split('|').map(p => p.trim());
            
            let email, refreshToken, clientId;
            if (parts.length >= 4) {
                [email, , refreshToken, clientId] = parts;
            } else if (parts.length === 3) {
                [email, refreshToken, clientId] = parts;
            } else {
                return res.status(400).json({ success: false, error: 'Format error: email|password|refresh_token|client_id required.' });
            }

            // Microsoft OAuth Access Token Request
            const tokenResponse = await axios.post(
                'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                new URLSearchParams({
                    client_id: clientId,
                    grant_type: 'refresh_token',
                    refresh_token: refreshToken,
                    scope: 'https://graph.microsoft.com/Mail.Read'
                }),
                { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
            );

            const accessToken = tokenResponse.data.access_token;
            if (!accessToken) {
                return res.json({ success: false, error: 'Failed to get Microsoft access token.' });
            }

            // Fetch Latest Email using Microsoft Graph API
            const graphResponse = await axios.get('https://graph.microsoft.com/v1.0/me/messages?$top=1', {
                headers: { Authorization: `Bearer ${accessToken}` }
            });

            const messages = graphResponse.data.value;
            if (!messages || messages.length === 0) {
                return res.json({ success: false, error: 'Inbox is empty or no messages found.' });
            }

            const latestMail = messages[0];
            const subject = latestMail.subject || '';
            const html = latestMail.body ? latestMail.body.content : '';
            const text = latestMail.bodyPreview || '';

            // Extract Code and Link using priority logic
            const { code, link } = parseVerificationDetails(subject, html, text);

            return res.json({
                success: true,
                subject: subject,
                code: code,
                link: link,
                fullHtml: html || text
            });

        } else if (provider === 'gmail' || provider === 'att') {
            const parts = accountData.split('|').map(p => p.trim());
            if (parts.length < 2) {
                return res.status(400).json({ success: false, error: 'Format error: email|app_password required.' });
            }

            const [email, appPassword] = parts;
            const host = provider === 'gmail' ? 'imap.gmail.com' : 'imap.mail.att.net';

            const imap = new Imap({
                user: email,
                password: appPassword,
                host: host,
                port: 993,
                tls: true,
                tlsOptions: { rejectUnauthorized: false }
            });

            fetchLatestEmail(imap, res);
        } else {
            return res.status(400).json({ success: false, error: 'Invalid provider selected.' });
        }
    } catch (err) {
        const errMsg = err.response?.data?.error_description || err.message || 'Server error occurred.';
        res.json({ success: false, error: errMsg });
    }
});

function fetchLatestEmail(imap, res) {
    imap.once('ready', () => {
        imap.openBox('INBOX', true, (err, box) => {
            if (err || box.messages.total === 0) {
                imap.end();
                return res.json({ success: false, error: 'Inbox is empty or inaccessible.' });
            }

            const fetch = imap.seq.fetch(`${box.messages.total}:${box.messages.total}`, { bodies: '' });
            
            fetch.on('message', (msg) => {
                msg.on('body', (stream) => {
                    simpleParser(stream, async (err, parsed) => {
                        imap.end();
                        if (err) return res.json({ success: false, error: 'Failed to parse email.' });

                        const subject = parsed.subject || '';
                        const text = parsed.text || '';
                        const html = parsed.html || parsed.textAsHtml || '';

                        const { code, link } = parseVerificationDetails(subject, html, text);

                        res.json({
                            success: true,
                            subject: subject,
                            code: code,
                            link: link,
                            fullHtml: html
                        });
                    });
                });
            });
        });
    });

    imap.once('error', (err) => {
        res.json({ success: false, error: 'IMAP connection error: ' + err.message });
    });

    imap.connect();
}

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});