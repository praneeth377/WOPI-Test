const express = require('express');
const fs = require('fs');
const path = require('path');
const jwt = require('jsonwebtoken');
const bodyParser = require('body-parser');
const cors = require('cors');

const app = express();
const secret = 'yoyo';
const filePath = path.join(__dirname, 'Sample Data.xlsx');

let lockStatus = { locked: false, lockId: null };

app.use(bodyParser.raw({ type: 'application/octet-stream', limit: '50mb' }));

app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    credentials: true,
}));

// Endpoint: CheckFileInfo
app.get('/wopi/files/:fileId', (req, res) => {
    const fileStats = fs.statSync(filePath);
    const fileInfo = {
        BaseFileName: 'Sample Data.xlsx',
        Size: fileStats.size,
        OwnerId: 'user1',
        UserId: 'user1',
        UserCanWrite: true,
        Version: 'v1',
        AllowExternalMarketplace: false,
        BreadcrumbBrandName: 'My Company',
        BreadcrumbBrandUrl: 'https://mycompany.com',
        BreadcrumbDocName: 'Sample Data.xlsx',
        BreadcrumbDocUrl: 'https://mycompany.com/sample-data.xlsx',
        CloseUrl: 'https://mycompany.com/close',
        FileExtension: '.xlsx',
        LastModifiedTime: fileStats.mtime.toISOString(),
        SupportsLocks: true,
        UserFriendlyName: 'User1',
        ReadOnly: false,
        SupportsExtendedLockLength: true,
        SupportsGetLock: true,
        SupportsEcosystem: true,
        SupportsUpdate: true
    };

    res.json(fileInfo);
});

// Endpoint: GetFile (returns the actual file contents)
app.get('/wopi/files/:fileId/contents', (req, res) => {
    // Read the file and stream it back to the client
    const fileStream = fs.createReadStream(filePath);
    res.setHeader('Content-Disposition', 'attachment; filename="Sample Data.xlsx"');
    fileStream.pipe(res);
});

// Endpoint: PutFile (saves the updated file content after editing)
app.post('/wopi/files/:fileId/contents', (req, res) => {
    const writeStream = fs.createWriteStream(filePath);
    req.pipe(writeStream);
    writeStream.on('finish', () => {
        res.sendStatus(200);
    });
    writeStream.on('error', (err) => {
        console.error('Error saving file:', err);
        res.sendStatus(500);
    });
});

// Endpoint for locking the file
app.post('/wopi/files/:fileId/lock', (req, res) => {
    const lockId = req.headers['x-wopi-lock'];

    if (!lockId || lockStatus.locked) {
        return res.status(409).send('File is already locked');
    }

    lockStatus = {
        locked: true,
        lockId,
    };

    res.sendStatus(200);
});

// Endpoint for unlocking the file
app.post('/wopi/files/:fileId/unlock', (req, res) => {
    const lockId = req.headers['x-wopi-lock'];

    if (lockStatus.lockId !== lockId) {
        return res.status(409).send('Invalid lock');
    }

    lockStatus = { locked: false, lockId: null };
    res.sendStatus(200);
});

// Generate an access token (JWT)
app.get('/generate-token', (req, res) => {
    const fileId = '1';
    const userId = 'user1';
    const token = jwt.sign(
        { fileId, userId, permissions: 'read-write' },
        secret,
        { expiresIn: '3h' }
    );

    const wopiSrc = encodeURIComponent(`http://localhost:3000/wopi/files/${fileId}`);
    const excelUrl = `https://excel.officeapps.live.com/x/_layouts/xlviewerinternal.aspx?WOPISrc=${wopiSrc}&access_token=${token}`;

    res.json({ token, excelUrl });
});

app.listen(3000, () => {
    console.log('WOPI Host is running on http://localhost:3000');
});
