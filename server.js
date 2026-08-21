const express = require('express');
const axios = require('axios');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/get-code', async (req, res) => {
    try {
        const { accountData } = req.body;
        if (!accountData) {
            return res.status(400).json({ success: false, error: 'Account data is required.' });
        }

        const parts = accountData.split('|');
        if (parts.length < 4) {
            return res.status(400).json({ success: false, error: 'Invalid format. Use email|password|refresh_token|client_id' });
        }

        const [email, password, refreshToken, clientId] = parts.map(p => p.trim());

        // OAuth2 Access Token Fetching
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

        // Fetch Recent Emails via Microsoft Graph API
        const mailRes = await axios.get('https://graph.microsoft.com/v1.0/me/messages?$top=5&$select=subject,bodyPreview,body', {
            headers: { Authorization: `Bearer ${accessToken}` }
        });

        const messages = mailRes.data.value;
        if (!messages || messages.length === 0) {
            return res.json({ success: false, error: 'No emails found in the inbox.' });
        }

        let foundCode = null;
        let actionLink = null;

        for (const msg of messages) {
            const bodyContent = msg.body?.content || msg.bodyPreview || '';
            const fullText = `${msg.subject} ${bodyContent}`;

            // OTP Code Matching (4-8 digits)
            if (!foundCode) {
                const codeMatch = fullText.match(/\b\d{4,8}\b/);
                if (codeMatch) foundCode = codeMatch[0];
            }

            // Link Extraction (Filtering out tracking images/assets)
            if (!actionLink) {
                const urlRegex = /(https?:\/\/[^\s"'<>]+)/gi;
                const foundUrls = bodyContent.match(urlRegex) || [];
                actionLink = foundUrls.find(link => 
                    !link.includes('.png') && 
                    !link.includes('.jpg') && 
                    !link.includes('.gif') &&
                    !link.includes('schemas.microsoft')
                ) || null;
            }

            if (foundCode && actionLink) break;
        }

        return res.json({
            success: true,
            code: foundCode,
            link: actionLink
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