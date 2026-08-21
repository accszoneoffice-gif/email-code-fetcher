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

    // Format check: email|password|refresh_token|client_id
    const parts = accountData.trim().split('|');
    if (parts.length < 4) {
        return res.status(400).json({ success: false, error: "Invalid format. Expected: email|password|refresh_token|client_id" });
    }

    const [email, password, refreshToken, clientId] = parts.map(p => p.trim());

    try {
        // Step 1: Exchange Refresh Token for Access Token
        const params = new URLSearchParams();
        params.append('client_id', clientId);
        params.append('grant_type', 'refresh_token');
        params.append('refresh_token', refreshToken);
        params.append('scope', 'https://graph.microsoft.com/Mail.Read offline_access');

        const tokenResponse = await axios.post(
            'https://login.microsoftonline.com/common/oauth2/v2.0/token',
            params,
            {
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            }
        );

        const accessToken = tokenResponse.data.access_token;

        if (!accessToken) {
            return res.status(400).json({ success: false, error: "Failed to obtain access token." });
        }

        // Step 2: Retrieve Recent Messages using Access Token
        const mailResponse = await axios.get(
            'https://graph.microsoft.com/v1.0/me/messages?$top=5&$select=subject,bodyPreview,body',
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                }
            }
        );

        const messages = mailResponse.data.value;
        if (!messages || messages.length === 0) {
            return res.json({ success: false, error: "No emails found in inbox." });
        }

        // Step 3: Extract verification code
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