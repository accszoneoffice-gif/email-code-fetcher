const express = require('express');
const axios = require('axios');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static('public'));

app.post('/api/fetch-hotmail', async (req, res) => {
    const { email, password, refresh_token, client_id } = req.body;

    try {
        const tokenResponse = await axios.post(
            'https://login.microsoftonline.com/common/oauth2/v2.0/token',
            new URLSearchParams({
                client_id: client_id,
                grant_type: 'refresh_token',
                refresh_token: refresh_token,
                scope: 'https://graph.microsoft.com/Mail.Read'
            }).toString(),
            { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
        );

        const accessToken = tokenResponse.data.access_token;

        const mailResponse = await axios.get(
            'https://graph.microsoft.com/v1.0/me/messages?$top=1',
            { headers: { Authorization: `Bearer ${accessToken}` } }
        );

        const messages = mailResponse.data.value;
        if (!messages || messages.length === 0) {
            return res.json({ success: false, message: 'No emails found.' });
        }

        const latestMail = messages[0].body.content || messages[0].bodyPreview;
        const codeMatch = latestMail.match(/\b\d{4,8}\b/); 
        const code = codeMatch ? codeMatch[0] : 'Code not found in latest email';

        return res.json({ success: true, code: code, subject: messages[0].subject });

    } catch (error) {
        return res.status(500).json({ success: false, error: error.response?.data || error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));