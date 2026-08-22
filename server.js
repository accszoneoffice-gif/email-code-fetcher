const express = require('express');
const axios = require('axios');
const Imap = require('node-imap');
const { simpleParser } = require('mailparser');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static('public'));

// --- 1. EMAIL FETCHER API ---
app.post('/api/get-code', async (req, res) => {
    const { provider, accountData } = req.body;

    if (!accountData) {
        return res.status(400).json({ success: false, error: 'Account data is required.' });
    }

    try {
        if (provider === 'outlook') {
            const parts = accountData.split('|').map(p => p.trim());
            if (parts.length < 4) {
                return res.status(400).json({ success: false, error: 'Format error: email|password|refresh_token|client_id required.' });
            }

            const [email, password, refreshToken, clientId] = parts;

            const tokenResponse = await axios.post(
                'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                new URLSearchParams({
                    client_id: clientId,
                    grant_type: 'refresh_token',
                    refresh_token: refreshToken,
                    scope: 'https://outlook.office.com/IMAP.AccessAsUser.All'
                }),
                { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
            );

            const accessToken = tokenResponse.data.access_token;
            if (!accessToken) {
                return res.json({ success: false, error: 'Failed to get Microsoft access token.' });
            }

            const authString = `user=${email}\x01auth=Bearer ${accessToken}\x01\x01`;
            const xoauth2Token = Buffer.from(authString).toString('base64');

            const imap = new Imap({
                xoauth2: xoauth2Token,
                host: 'outlook.office365.com',
                port: 993,
                tls: true,
                tlsOptions: { rejectUnauthorized: false }
            });

            fetchLatestEmail(imap, res);

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
        res.json({ success: false, error: err.message || 'Server error occurred.' });
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

                        const codeMatch = text.match(/\b\d{4,8}\b/) || html.match(/\b\d{4,8}\b/);
                        const code = codeMatch ? codeMatch[0] : null;

                        const linkMatch = html.match(/href=["'](https?:\/\/[^"']+)["']/i) || text.match(/(https?:\/\/[^\s]+)/i);
                        const link = linkMatch ? linkMatch[1] : null;

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

// --- 2. SOCIAL MEDIA CHECKER API ---
app.post('/api/check-social', async (req, res) => {
    const { urls } = req.body;
    if (!urls || !Array.isArray(urls)) {
        return res.status(400).json({ error: 'URLs array is required.' });
    }

    const results = [];

    for (let url of urls) {
        try {
            const response = await axios.get(url, {
                headers: {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
                },
                timeout: 5000,
                validateStatus: false
            });

            if (response.status === 200) {
                results.push({ url, status: 'LIVE' });
            } else {
                results.push({ url, status: 'DISABLED' });
            }
        } catch (err) {
            results.push({ url, status: 'DISABLED' });
        }
    }

    res.json({ results });
});

app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});