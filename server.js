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

    // Parse format: email|password|refresh_token|client_id
    const parts = accountData.trim().split('|');
    if (parts.length < 4) {
        return res.status(400).json({ success: false, error: "Invalid format. Expected: email|password|refresh_token|client_id" });
    }

    const [email, password, refreshToken, clientId] = parts;

    try {
        // Step 1: Exchange Refresh Token for Access Token with required scopes
        const tokenParams = new URLSearchParams({
            client_id: clientId,
            grant_type: 'refresh_token',
            refresh_token: refreshToken,
            scope: 'https://graph.microsoft.com/Mail.Read offline_access'
        });

        const tokenResponse = await axios.post(
            'https://login.microsoftonline.com/common/oauth2/v2.0/token',
            tokenParams.toString(),
            { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
        );

        const accessToken = tokenResponse.data.access_token;

        // Step 2: Fetch recent messages from Microsoft Graph API
        const mailResponse = await axios.get(
            'https://graph.microsoft.com/v1.0/me/messages?$top=5&$select=subject,bodyPreview,receivedDateTime',
            { headers: { Authorization: `Bearer ${accessToken}` } }
        );

        const messages = mailResponse.data.value;
        if (!messages || messages.length === 0) {
            return res.json({ success: false, error: "No emails found" });
        }

        // Step 3: Extract 4 to 8 digit verification codes using Regex
        let foundCode = null;
        for (const msg of messages) {
            const text = `${msg.subject} ${msg.bodyPreview}`;
            const match = text.match(/\b\d{4,8}\b/);
            if (match) {
                foundCode = match[0];
                break;
            }
        }

        if (foundCode) {
            return res.json({ success: true, code: foundCode });
        } else {
            return res.json({ success: false, error: "Verification code not found in recent emails" });
        }

    } catch (err) {
        return res.status(500).json({
            success: false,
            error: err.response ? err.response.data : err.message
        });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});