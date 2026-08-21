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

        // Split exact format: email|password|refresh_token|client_id
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

        // Step 1: Exchange Refresh Token for Access Token using consumers endpoint
        const tokenParams = new URLSearchParams();
        tokenParams.append('client_id', clientId);
        tokenParams.append('grant_type', 'refresh_token');
        tokenParams.append('refresh_token', refreshToken);
        tokenParams.append('scope', 'https://graph.microsoft.com/Mail.Read offline_access');

        let accessToken;
        try {
            const tokenRes = await axios.post(
                'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
                tokenParams.toString(),
                {
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
                }
            );
            accessToken = tokenRes.data.access_token;
        } catch (tokenErr) {
            const errDetails = tokenErr.response ? tokenErr.response.data : tokenErr.message;
            return res.status(400).json({
                success: false,
                error: { message: "Failed to exchange refresh token", details: errDetails }
            });
        }

        if (!accessToken) {
            return res.status(400).json({ success: false, error: "Access token generation returned empty." });
        }

        // Step 2: Fetch last 5 messages from Microsoft Graph API
        const mailRes = await axios.get(
            'https://graph.microsoft.com/v1.0/me/messages?$top=5&$select=subject,bodyPreview,body',
            {
                headers: { 'Authorization': `Bearer ${accessToken}` }
            }
        );

        const messages = mailRes.data.value;
        if (!messages || messages.length === 0) {
            return res.json({ success: false, error: "No emails found in this account." });
        }

        // Step 3: Extract verification code using RegEx
        let foundCode = null;
        for (const msg of messages) {
            const fullText = `${msg.subject || ''} ${msg.bodyPreview || ''} ${msg.body?.content || ''}`;
            
            const match = fullText.match(/(?:code|pin|verification|otp|is|g-)[\s:\-]*([0-9]{4,8})/i) || fullText.match(/\b[0-9]{4,8}\b/);
            
            if (match) {
                foundCode = match[1] || match[0];
                break;
            }
        }

        if (foundCode) {
            return res.json({ success: true, code: foundCode });
        } else {
            return res.json({ success: false, error: "No 4-8 digit verification code found in recent emails." });
        }

    } catch (globalErr) {
        const errorData = globalErr.response ? globalErr.response.data : globalErr.message;
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