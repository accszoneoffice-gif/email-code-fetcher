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
            return res.status(400).json({ success: false, error: "Account data is required" });
        }

        // Parse: email|password|refresh_token|client_id
        const parts = accountData.trim().split('|').map(p => p.trim());
        if (parts.length < 3) {
            return res.status(400).json({ 
                success: false, 
                error: "Invalid input format. Expected: email|password|refresh_token|client_id" 
            });
        }

        const email = parts[0];
        const password = parts[1];
        const refreshToken = parts[2];
        const clientId = parts[3] || '9e5f94bc-e8a4-4e73-b8be-63364c29d753';

        // Direct token exchange endpoint used by legacy Hotmail tools
        const tokenParams = new URLSearchParams({
            client_id: clientId,
            grant_type: 'refresh_token',
            refresh_token: refreshToken
        });

        let accessToken;
        try {
            // Using login.live.com endpoint directly for Hotmail/Outlook tokens
            const tokenRes = await axios.post(
                'https://login.live.com/oauth20_token.srf',
                tokenParams.toString(),
                {
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
                }
            );
            accessToken = tokenRes.data.access_token;
        } catch (err) {
            // Fallback to Microsoft v2.0 if login.live fails
            try {
                const tokenRes2 = await axios.post(
                    'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                    tokenParams.toString(),
                    { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
                );
                accessToken = tokenRes2.data.access_token;
            } catch (err2) {
                return res.status(400).json({
                    success: false,
                    error: err2.response ? err2.response.data : err2.message
                });
            }
        }

        if (!accessToken) {
            return res.status(400).json({ success: false, error: "Access token could not be generated." });
        }

        // Fetch Recent Messages via Graph API
        const mailRes = await axios.get(
            'https://graph.microsoft.com/v1.0/me/messages?$top=10&$select=subject,bodyPreview,body',
            {
                headers: { 'Authorization': `Bearer ${accessToken}` }
            }
        );

        const messages = mailRes.data.value;
        if (!messages || messages.length === 0) {
            return res.json({ success: false, error: "No emails found in this account." });
        }

        // Extract 4-8 digit verification code
        let foundCode = null;
        for (const msg of messages) {
            const content = `${msg.subject || ''} ${msg.bodyPreview || ''} ${msg.body?.content || ''}`;
            const match = content.match(/(?:code|pin|verification|otp|is)[\s:\-]*([0-9]{4,8})/i) || content.match(/\b[0-9]{4,8}\b/);
            
            if (match) {
                foundCode = match[1] || match[0];
                break;
            }
        }

        if (foundCode) {
            return res.json({ success: true, code: foundCode });
        } else {
            return res.json({ success: false, error: "No verification code found in recent emails." });
        }

    } catch (globalErr) {
        return res.status(500).json({
            success: false,
            error: globalErr.response ? globalErr.response.data : globalErr.message
        });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});