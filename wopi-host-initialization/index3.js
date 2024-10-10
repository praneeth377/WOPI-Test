const express = require('express');
const jwt = require('jsonwebtoken');
const axios = require('axios');
const bodyParser = require('body-parser');
const cors = require('cors');

const app = express();
const secret = 'yoyo';
const filePath = 'Sample Data.xlsx';

app.use(bodyParser.json());
app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    credentials: true,
}));

// Middleware to authenticate JWT
const authenticateToken = (req, res, next) => {
    const authHeader = req.headers['authorization'];
    const token = authHeader && authHeader.split(' ')[1];

    if (!token) return res.sendStatus(401); // Unauthorized if no token is provided

    jwt.verify(token, secret, (err, user) => {
        if (err) return res.sendStatus(403); // Forbidden if token is invalid
        req.user = user;
        next();
    });
};

// Endpoint: Generate Token
app.get('/generate-token', (req, res) => {
    const fileId = '1';
    const userId = 'user1';
    const token = jwt.sign(
        { fileId, userId, permissions: 'read-write' },
        secret,
        { expiresIn: '3h' }
    );
    res.json({ accessToken: token });
});

// Endpoint: Generate Share URL
app.post('/generate-share-url', authenticateToken, async (req, res) => {
    const { fileId } = req.body;

    if (!fileId) {
        return res.status(400).json({ error: 'fileId is required' });
    }

    try {
        const accessToken = req.headers.authorization.split(' ')[1];
        const response = await axios.post(
            `https://api.office.com/v1.0/me/drive/items/${fileId}/action.createLink`,
            {
                type: "view"
            },
            {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            }
        );

        const shareUrl = response.data.link.webUrl;
        res.json({ shareUrl });
    } catch (error) {
        console.error('Error generating share URL:', error);
        res.status(500).json({ error: 'Failed to generate share URL' });
    }
});

// Start the server
app.listen(3000, () => {
    console.log('Server is running on http://localhost:3000');
});
