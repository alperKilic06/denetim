const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
const port = 3000;

// Create uploads directory if it doesn't exist
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir);
}

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, uploadsDir);
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + '-' + file.originalname);
    }
});

const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        const ext = path.extname(file.originalname).toLowerCase();
        if (ext !== '.xlsx' && ext !== '.xls') {
            return cb(new Error('Sadece Excel dosyaları yüklenebilir (.xlsx, .xls)'));
        }
        cb(null, true);
    }
});

// Serve static files
app.use(express.static('public'));
app.use(express.json());

// Excel file upload and processing endpoint
app.post('/upload', upload.single('excelFile'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'Lütfen bir Excel dosyası yükleyin' });
        }

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(req.file.path);

        // Extract data from all sheets
        const sheetsData = [];
        
        workbook.eachSheet((worksheet, sheetId) => {
            const sheetData = {
                name: worksheet.name,
                rowCount: worksheet.rowCount,
                columnCount: worksheet.columnCount,
                data: []
            };

            worksheet.eachRow((row, rowNumber) => {
                const rowData = [];
                row.eachCell((cell, colNumber) => {
                    rowData.push({
                        value: cell.value,
                        type: cell.type
                    });
                });
                sheetData.data.push(rowData);
            });

            sheetsData.push(sheetData);
        });

        // Clean up uploaded file
        fs.unlinkSync(req.file.path);

        res.json({
            success: true,
            fileName: req.file.originalname,
            sheets: sheetsData
        });

    } catch (error) {
        console.error('Excel işleme hatası:', error);
        res.status(500).json({ error: 'Excel dosyası işlenirken hata oluştu: ' + error.message });
    }
});

// Generate report endpoint
app.post('/generate-report', express.json(), async (req, res) => {
    try {
        const { sheetsData } = req.body;
        
        if (!sheetsData || !Array.isArray(sheetsData)) {
            return res.status(400).json({ error: 'Geçersiz veri formatı' });
        }

        // Generate HTML report
        let reportHtml = `
            <!DOCTYPE html>
            <html lang="tr">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Denetim Raporu</title>
                <style>
                    body { font-family: Arial, sans-serif; margin: 20px; }
                    h1 { color: #333; }
                    h2 { color: #666; margin-top: 30px; }
                    table { border-collapse: collapse; width: 100%; margin-top: 10px; }
                    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                    th { background-color: #4CAF50; color: white; }
                    tr:nth-child(even) { background-color: #f2f2f2; }
                    .summary { background-color: #e7f3ff; padding: 15px; margin: 20px 0; border-radius: 5px; }
                </style>
            </head>
            <body>
                <h1>Denetim Raporu</h1>
                <div class="summary">
                    <strong>Rapor Tarihi:</strong> ${new Date().toLocaleDateString('tr-TR')}<br>
                    <strong>Toplam Sayfa Sayısı:</strong> ${sheetsData.length}
                </div>
        `;

        sheetsData.forEach((sheet, index) => {
            reportHtml += `
                <h2>Sayfa ${index + 1}: ${sheet.name}</h2>
                <p><strong>Satır Sayısı:</strong> ${sheet.rowCount} | <strong>Sütun Sayısı:</strong> ${sheet.columnCount}</p>
                <table>
            `;

            // Add table rows (limit to first 100 rows for performance)
            const maxRows = Math.min(sheet.data.length, 100);
            sheet.data.slice(0, maxRows).forEach((row, rowIndex) => {
                reportHtml += '<tr>';
                row.forEach(cell => {
                    const tag = rowIndex === 0 ? 'th' : 'td';
                    let value = cell.value;
                    
                    // Handle different cell types
                    if (value && typeof value === 'object') {
                        if (value.formula) {
                            value = value.result || '';
                        } else if (value.richText) {
                            value = value.richText.map(rt => rt.text).join('');
                        } else {
                            value = JSON.stringify(value);
                        }
                    }
                    
                    reportHtml += `<${tag}>${value || ''}</${tag}>`;
                });
                reportHtml += '</tr>';
            });

            if (sheet.data.length > maxRows) {
                reportHtml += `<tr><td colspan="${sheet.data[0]?.length || 1}"><em>... ve ${sheet.data.length - maxRows} satır daha</em></td></tr>`;
            }

            reportHtml += '</table>';
        });

        reportHtml += `
            </body>
            </html>
        `;

        res.send(reportHtml);

    } catch (error) {
        console.error('Rapor oluşturma hatası:', error);
        res.status(500).json({ error: 'Rapor oluşturulurken hata oluştu: ' + error.message });
    }
});

app.listen(port, () => {
    console.log(`Denetim uygulaması http://localhost:${port} adresinde çalışıyor`);
});
