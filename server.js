const express = require('express');
const axios = require('axios');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/get-code', async (req, res) => {
    const { accountData } = req.body;

    if (!accountData) {
        return res.status(400).json({ success: false, error: "Account data is required" });
    }

    const parts = accountData.trim().split('|');
    if (parts.length < 4) {
        return res.status(400).json({ success: false, error: "Invalid format. Expected: email|password|refresh_token|client_id" });
    }

    const [email, password, refreshToken, clientId] = parts.map(p => p.trim());

    try {
        // Exchange Refresh Token for Access Token
        const tokenParams = new URLSearchParams({
            client_id: clientId,
            grant_type: 'refresh_token',
            refresh_token: refreshToken
        });

        const tokenResponse = await axios.post(
            'https://login.microsoftonline.com/common/oauth2/v2.0/token',
            tokenParams.toString(),
            { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
        );

        const accessToken = tokenResponse.data.access_token;

        // Fetch Recent Messages
        const mailResponse = await axios.get(
            'https://graph.microsoft.com/v1.0/me/messages?$top=5&$select=subject,bodyPreview,body',
            { headers: { Authorization: `Bearer ${accessToken}` } }
        );

        const messages = mailResponse.data.value;
        if (!messages || messages.length === 0) {
            return res.json({ success: false, error: "No emails found in inbox." });
        }

        let foundCode = null;
        for (const msg of messages) {
            const content = `${msg.subject || ''} ${msg.bodyPreview || ''} ${msg.body?.content || ''}`;
            
            // Search for 4 to 8 digit verification code
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

    } catch (err) {
        const errorData = err.response ? err.response.data : err.message;
        return res.status(500).json({
            success: false,
            error: errorData
        });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});