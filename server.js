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

        // Split data safely using pipe (|)
        const rawParts = accountData.trim().split('|');
        if (rawParts.length < 3) {
            return res.status(400).json({ 
                success: false, 
                error: "Invalid format. Expected: email|password|refresh_token|client_id" 
            });
        }

        const email = rawParts[0].trim();
        const password = rawParts[1].trim();
        const refreshToken = rawParts[2].trim();
        const clientId = rawParts[3] ? rawParts[3].trim() : '9e5f94bc-e8a4-4e73-b8be-63364c29d753';

        // Step 1: Request Access Token from Microsoft OAuth Server
        const postData = new URLSearchParams({
            client_id: clientId,
            grant_type: 'refresh_token',
            refresh_token: refreshToken,
            scope: 'https://graph.microsoft.com/Mail.Read offline_access'
        });

        let accessToken = null;
        try {
            const tokenRes = await axios.post(
                'https://login.microsoftonline.com/consumers/oauth2/v2.0/token',
                postData.toString(),
                {
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
                }
            );
            accessToken = tokenRes.data.access_token;
        } catch (tokenErr) {
            const errorDetails = tokenErr.response ? tokenErr.response.data : tokenErr.message;
            return res.status(400).json({
                success: false,
                error: errorDetails
            });
        }

        if (!accessToken) {
            return res.status(400).json({ success: false, error: "Failed to generate access token." });
        }

        // Step 2: Fetch Recent Emails via Graph API
        const mailRes = await axios.get(
            'https://graph.microsoft.com/v1.0/me/messages?$top=5&$select=subject,bodyPreview,body',
            {
                headers: { 
                    'Authorization': `Bearer ${accessToken}`,
                    'Prefer': 'outlook.body-type="text"'
                }
            }
        );

        const messages = mailRes.data.value;
        if (!messages || messages.length === 0) {
            return res.json({ success: false, error: "No emails found in this account." });
        }

        // Step 3: Extract verification code using Regex
        let foundCode = null;
        for (const msg of messages) {
            const fullText = `${msg.subject || ''} ${msg.bodyPreview || ''} ${msg.body?.content || ''}`;
            
            // Search for 4 to 8 digit verification codes
            const match = fullText.match(/(?:code|pin|verification|otp|is|g-)[\s:\-]*([0-9]{4,8})/i) || fullText.match(/\b[0-9]{4,8}\b/);
            
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