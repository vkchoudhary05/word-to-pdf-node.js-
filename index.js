const express = require('express');
const multer = require('multer');
const pdf = require('pdf-parse');
const { Document, Packer, Paragraph, HeadingLevel } = require('docx');
const fs = require('fs');
const path = require('path');

const app = express();
const port = 3000;


app.use(express.static(path.join(__dirname, 'public')));


const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/');
    },
    filename: (req, file, cb) => {
        cb(null, file.originalname);
    },
});

const upload = multer({
    storage: storage,
    limits: { fileSize: 10 * 1024 * 1024 }, 
    fileFilter: (req, file, cb) => {
        if (file.mimetype !== 'application/pdf') {
            return cb(new Error('Only PDF files are allowed.'));
        }
        cb(null, true);
    }
});


app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'homePage.html'));
});


app.get('/downloadPage', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'download.html'));
});


app.post('/upload', upload.single('pdfFile'), async (req, res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    if (req.file.mimetype !== 'application/pdf') {
        return res.status(400).send('Only PDF files are allowed.');
    }

    try {
        const pdfPath = path.join(__dirname, 'uploads', req.file.filename);
        const pdfBuffer = fs.readFileSync(pdfPath);
        const data = await pdf(pdfBuffer);
        const text = data.text;

        const doc = new Document({
            sections: [
                {
                    properties: {},
                    children: [
                        new Paragraph({
                            text: text,
                            heading: HeadingLevel.HEADING_1,
                        }),
                    ],
                },
            ],
        });

        const docxFilename = req.file.filename.replace('.pdf', '.docx');
        const docxPath = path.join(__dirname, 'uploads', docxFilename);
        const docxBuffer = await Packer.toBuffer(doc);
        fs.writeFileSync(docxPath, docxBuffer);

    
        res.redirect(`/downloadPage?link=/download/${docxFilename}`);
    } catch (error) {
        console.error('Error during conversion:', error);
        res.status(500).send('Internal Server Error');
    }
});


app.get('/download/:filename', (req, res) => {
    const filePath = path.join(__dirname, 'uploads', req.params.filename);
    if (fs.existsSync(filePath)) {
        res.download(filePath, (err) => {
            if (err) {
                console.error('Error downloading file:', err);
            } else {
                
                fs.unlink(filePath, (unlinkErr) => {
                    if (unlinkErr) {
                        console.error('Error deleting file:', unlinkErr);
                    } else {
                        console.log('File deleted successfully after download:', req.params.filename);
                    }
                });
            }
        });
    } else {
        res.status(404).send('File not found.');
    }
});


if (!fs.existsSync('uploads')) {
    fs.mkdirSync('uploads');
}

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
