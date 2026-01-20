let uploadedData = null;

// File input change event
document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (file) {
        showFileInfo(file);
    }
});

// Drag and drop functionality
const uploadBox = document.getElementById('uploadBox');

uploadBox.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadBox.style.background = '#e8f0ff';
});

uploadBox.addEventListener('dragleave', (e) => {
    e.preventDefault();
    uploadBox.style.background = '#f8f9ff';
});

uploadBox.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadBox.style.background = '#f8f9ff';
    
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
        document.getElementById('fileInput').files = e.dataTransfer.files;
        showFileInfo(file);
    } else {
        showError('Lütfen geçerli bir Excel dosyası seçin (.xlsx veya .xls)');
    }
});

function showFileInfo(file) {
    document.getElementById('fileName').textContent = file.name;
    document.getElementById('fileInfo').style.display = 'block';
}

function resetUpload() {
    document.getElementById('fileInput').value = '';
    document.getElementById('fileInfo').style.display = 'none';
    document.getElementById('results').style.display = 'none';
    document.getElementById('report').style.display = 'none';
    uploadedData = null;
}

// Upload button click event
document.getElementById('uploadBtn').addEventListener('click', async function() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    
    if (!file) {
        showError('Lütfen bir dosya seçin');
        return;
    }

    const formData = new FormData();
    formData.append('excelFile', file);

    // Show loading
    document.getElementById('loading').style.display = 'block';
    document.getElementById('fileInfo').style.display = 'none';
    document.getElementById('results').style.display = 'none';
    document.getElementById('report').style.display = 'none';

    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (response.ok && result.success) {
            uploadedData = result;
            displayResults(result);
        } else {
            showError(result.error || 'Dosya yüklenirken bir hata oluştu');
        }
    } catch (error) {
        showError('Sunucu ile bağlantı kurulamadı: ' + error.message);
    } finally {
        document.getElementById('loading').style.display = 'none';
    }
});

function displayResults(data) {
    const sheetsInfoDiv = document.getElementById('sheetsInfo');
    sheetsInfoDiv.innerHTML = '';

    data.sheets.forEach((sheet, index) => {
        const sheetCard = document.createElement('div');
        sheetCard.className = 'sheet-card';
        sheetCard.innerHTML = `
            <h3>📄 ${sheet.name}</h3>
            <p><strong>Satır Sayısı:</strong> ${sheet.rowCount}</p>
            <p><strong>Sütun Sayısı:</strong> ${sheet.columnCount}</p>
            <p><strong>Toplam Hücre:</strong> ${sheet.rowCount * sheet.columnCount}</p>
        `;
        sheetsInfoDiv.appendChild(sheetCard);
    });

    document.getElementById('results').style.display = 'block';
}

// Generate report button
document.getElementById('generateReportBtn').addEventListener('click', async function() {
    if (!uploadedData) {
        showError('Önce bir Excel dosyası yükleyin');
        return;
    }

    document.getElementById('loading').style.display = 'block';

    try {
        const response = await fetch('/generate-report', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ sheetsData: uploadedData.sheets })
        });

        if (response.ok) {
            const reportHtml = await response.text();
            displayReport(reportHtml);
        } else {
            const error = await response.json();
            showError(error.error || 'Rapor oluşturulurken bir hata oluştu');
        }
    } catch (error) {
        showError('Rapor oluşturulamadı: ' + error.message);
    } finally {
        document.getElementById('loading').style.display = 'none';
    }
});

function displayReport(html) {
    const reportSection = document.getElementById('report');
    reportSection.innerHTML = `
        <h2>✅ Rapor Oluşturuldu</h2>
        <div style="margin: 20px 0;">
            <button class="btn btn-primary" onclick="printReport()">🖨️ Yazdır</button>
            <button class="btn btn-primary" onclick="downloadReport()">💾 İndir</button>
            <button class="btn btn-secondary" onclick="resetUpload()">🔄 Yeni Rapor</button>
        </div>
        <iframe id="reportFrame" style="width: 100%; min-height: 600px; border: 1px solid #ddd; border-radius: 5px;"></iframe>
    `;
    
    const iframe = document.getElementById('reportFrame');
    iframe.contentWindow.document.open();
    iframe.contentWindow.document.write(html);
    iframe.contentWindow.document.close();
    
    reportSection.style.display = 'block';
    reportSection.scrollIntoView({ behavior: 'smooth' });
}

function printReport() {
    const iframe = document.getElementById('reportFrame');
    if (iframe) {
        iframe.contentWindow.print();
    }
}

function downloadReport() {
    const iframe = document.getElementById('reportFrame');
    if (iframe) {
        const reportHtml = iframe.contentWindow.document.documentElement.outerHTML;
        const blob = new Blob([reportHtml], { type: 'text/html' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'denetim-raporu-' + new Date().toISOString().split('T')[0] + '.html';
        a.click();
        URL.revokeObjectURL(url);
    }
}

function showError(message) {
    const errorDiv = document.createElement('div');
    errorDiv.className = 'error';
    errorDiv.innerHTML = `<strong>Hata:</strong> ${message}`;
    
    const container = document.querySelector('.container');
    const existingError = container.querySelector('.error');
    if (existingError) {
        existingError.remove();
    }
    
    container.insertBefore(errorDiv, container.firstChild.nextSibling);
    
    setTimeout(() => {
        errorDiv.remove();
    }, 5000);
}
