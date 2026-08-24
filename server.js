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
            // email|password|refresh_token|client_id
            if (parts.length >= 4) {
                [email, , refreshToken, clientId] = parts;
            } else if (parts.length === 3) {
                [email, refreshToken, clientId] = parts;
            } else {
                return res.status(400).json({ success: false, error: 'Format error: email|password|refresh_token|client_id required.' });
            }

            // Fallback Client ID (Default Live/Outlook Client ID)
            const finalClientId = clientId || '0000000048170277';

            let accessToken = null;

            // Attempt 1: Fetch Token via Microsoft Live OAuth Endpoint (Recommended for Personal Outlook/Hotmail)
            try {
                const liveTokenRes = await axios.post(
                    'https://login.live.com/oauth20_token.srf',
                    new URLSearchParams({
                        client_id: finalClientId,
                        grant_type: 'refresh_token',
                        refresh_token: refreshToken,
                        scope: 'service::outlook.com::MBI_SSL'
                    }),
                    { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
                );
                accessToken = liveTokenRes.data.access_token;
            } catch (err1) {
                // Attempt 2: Fallback to Microsoft Online OAuth Endpoint if Live endpoint fails
                try {
                    const onlineTokenRes = await axios.post(
                        'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                        new URLSearchParams({
                            client_id: finalClientId,
                            grant_type: 'refresh_token',
                            refresh_token: refreshToken,
                            scope: 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access'
                        }),
                        { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
                    );
                    accessToken = onlineTokenRes.data.access_token;
                } catch (err2) {
                    const errMsg = err2.response?.data?.error_description || err1.response?.data?.error_description || 'Failed to authenticate refresh token with Microsoft.';
                    return res.json({ success: false, error: errMsg });
                }
            }

            if (!accessToken) {
                return res.json({ success: false, error: 'Failed to obtain access token.' });
            }

            // Fetch Email via Microsoft Graph API
            const graphResponse = await axios.get('https://graph.microsoft.com/v1.0/me/messages?$top=1', {
                headers: { Authorization: `Bearer ${accessToken}` }
            }).catch(async () => {
                // Fallback to Outlook REST API if Graph API scopes fail
                return await axios.get('https://outlook.office.com/api/v2.0/me/messages?$top=1', {
                    headers: { Authorization: `Bearer ${accessToken}` }
                });
            });

            const messages = graphResponse.data.value || graphResponse.data.value;
            if (!messages || messages.length === 0) {
                return res.json({ success: false, error: 'Inbox is empty or no messages found.' });
            }

            const latestMail = messages[0];
            const subject = latestMail.subject || '';
            const html = latestMail.body ? latestMail.body.content : '';
            const text = latestMail.bodyPreview || '';

            // Extract Verification Code (4 to 8 digits)
            const codeMatch = text.match(/\b\d{4,8}\b/) || html.match(/\b\d{4,8}\b/);
            const code = codeMatch ? codeMatch[0] : null;

            // Extract Direct Verification Link
            const linkMatch = html.match(/href=["'](https?:\/\/[^"']+)["']/i) || text.match(/(https?:\/\/[^\s]+)/i);
            const link = linkMatch ? linkMatch[1] : null;

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

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});