import express, { Request, Response } from 'express';
import multer from 'multer';
import path from 'path';
import { extractTablesFromWorkbook } from './extractService';

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Serve static files from public directory
app.use(express.static(path.join(__dirname, '../public')));

app.post('/extract', upload.single('file'), (req: Request, res: Response) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded. Use form-data with key "file".' });
        }

        const inputBuffer = req.file.buffer;
        const outputBuffer = extractTablesFromWorkbook(inputBuffer);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="combined_output.xlsx"');
        res.send(outputBuffer);
    } catch (error: any) {
        console.error('Error processing file:', error);
        res.status(500).json({ error: error.message || 'Failed to process Excel file' });
    }
});

app.get('/health', (_req: Request, res: Response) => {
    res.json({ status: 'ok' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
    console.log(`POST /extract - Upload Excel file to extract tables`);
    console.log(`GET /health - Health check`);
});
