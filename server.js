const express = require('express');
const axios = require('axios');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ইমেইল বডি থেকে ৬ বা ৫ ডিজিটের ওটিপি (OTP) বা ভেরিফিকেশন কোড খুঁজে বের করার ফাংশন
function extractOTP(text) {
    if (!text) return null;
    
    // ১. সাধারণ ৬ বা ৪-৮ ডিজিটের কোড খোঁজা (যেমন: 123456, 47164)
    const codeMatch = text.match(/\b\d{4,8}\b/);
    if (codeMatch) return codeMatch[0];

    // ২. আলফানিউমেরিক কোড খোঁজা (যেমন: G-123456)
    const alphaCodeMatch = text.match(/\b[A-Za-z0-9]{5,8}\b/);
    if (alphaCodeMatch) return alphaCodeMatch[0];

    return null;
}

app.post('/api/get-code', async (req, res) => {
    try {
        const { accountData } = req.body;

        if (!accountData) {
            return res.status(400).json({ success: false, error: "Account details are required" });
        }

        // ইনপুট ফরম্যাট: email|password|refresh_token|client_id
        const parts = accountData.trim().split('|').map(p => p.trim());
        if (parts.length < 4) {
            return res.status(400).json({ 
                success: false, 
                error: "Invalid format. Required: email|password|refresh_token|client_id" 
            });
        }

        const [email, password, refreshToken, clientId] = parts;

        // ১. মাইক্রোসফটের অফিশিয়াল ওথ২ এন্ডপয়েন্ট থেকে নতুন Access Token নেওয়া
        const tokenParams = new URLSearchParams({
            client_id: clientId,
            grant_type: 'refresh_token',
            refresh_token: refreshToken,
            scope: 'https://outlook.office.com/mail.read'
        });

        let tokenResponse;
        try {
            tokenResponse = await axios.post(
                'https://login.microsoftonline.com/common/oauth2/v2.0/token',
                tokenParams.toString(),
                { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
            );
        } catch (tokenErr) {
            return res.status(400).json({
                success: false,
                error: "Failed to authenticate with Microsoft. Invalid Refresh Token or Client ID."
            });
        }

        const accessToken = tokenResponse.data.access_token;

        // ২. মাইক্রোসফট আউটলুক/গ্রাফ এপিআই দিয়ে সাম্প্রতিক সব ইমেইল ফেচ করা (কোনো প্ল্যাটফর্ম ফিল্টার ছাড়া)
        const mailResponse = await axios.get(
            'https://outlook.office.com/api/v2.0/me/messages?$top=5&$select=Subject,From,Body,ReceivedDateTime',
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Accept': 'application/json'
                }
            }
        );

        const messages = mailResponse.data.value;

        if (!messages || messages.length === 0) {
            return res.json({ success: false, error: "No emails found in this inbox." });
        }

        // ৩. যেকোনো সোশ্যাল মিডিয়া বা সার্ভিস থেকে আসা সাম্প্রতিক ইমেইল প্রসেস করা
        let foundCode = null;
        let formattedMessages = [];

        for (const msg of messages) {
            const sender = msg.From?.EmailAddress?.Address || msg.From?.EmailAddress?.Name || 'Unknown';
            const subject = msg.Subject || 'No Subject';
            const bodyContent = msg.Body?.Content || '';

            // ওটিপি বা কোড এক্সট্রাক্ট করা
            const code = extractOTP(subject) || extractOTP(bodyContent);

            if (!foundCode && code) {
                foundCode = code; // সর্বশেষে আসা ইমেইলের কোডটি ধরা হবে
            }

            formattedMessages.push({
                from: sender,
                subject: subject,
                date: msg.ReceivedDateTime,
                code: code || 'N/A'
            });
        }

        if (foundCode) {
            return res.json({
                success: true,
                code: foundCode,
                messages: formattedMessages
            });
        } else {
            return res.json({
                success: false,
                error: "Emails retrieved, but no verification code was detected in recent messages."
            });
        }

    } catch (err) {
        return res.status(500).json({
            success: false,
            error: err.response ? JSON.stringify(err.response.data) : err.message
        });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});