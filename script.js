document.getElementById('generateButton').addEventListener('click', handleFile);
document.getElementById('downloadButton').addEventListener('click', downloadExcel);

let summaryData = [];

function handleFile() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    const warningDiv = document.getElementById('warning');
    
    if (file) {
        if (file.name.endsWith('.xls')) {
            warningDiv.textContent = "Uyarı: .xls formatı kullanıyorsunuz. Bu format hatalı sonuçlara neden olabilir. Lütfen .xlsx formatını kullanın.";
            warningDiv.style.display = 'block';
        } else {
            warningDiv.style.display = 'none';
        }

        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            processExcel(data);
        };
        reader.readAsArrayBuffer(file);
    } else {
        alert('Lütfen bir Excel dosyası seçin.');
    }
}

function processExcel(data) {
    const workbook = XLSX.read(data, {type: 'array'});
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 'A' });

    const summary = jsonData.reduce((acc, row, index) => {
        if (index < 4) return acc; // Başlık satırlarını atla
        
        const Urun = row['N'];
        const Miktar = parseFloat(row['O']) || 0;
        const Tutar = parseFloat(row['S']) || 0;
        
        if (Urun) {
            if (!acc[Urun]) {
                acc[Urun] = { toplamMiktar: 0, toplamTutar: 0 };
            }
            acc[Urun].toplamMiktar += Miktar;
            acc[Urun].toplamTutar += Tutar;
        }
        
        return acc;
    }, {});

    let genelToplamMiktar = 0;
    let genelToplamTutar = 0;
    summaryData = Object.entries(summary)
        .filter(([_, { toplamMiktar, toplamTutar }]) => toplamMiktar > 0 || toplamTutar > 0)
        .map(([urun, { toplamMiktar, toplamTutar }]) => {
            genelToplamMiktar += toplamMiktar;
            genelToplamTutar += toplamTutar;
            return { Urun: urun, ToplamMiktar: toplamMiktar, ToplamTutar: toplamTutar };
        });

    if (genelToplamMiktar > 0 || genelToplamTutar > 0) {
        summaryData.push({
            Urun: 'GENEL TOPLAM',
            ToplamMiktar: genelToplamMiktar,
            ToplamTutar: genelToplamTutar
        });
    }

    displayResults(summaryData);
    document.getElementById('downloadButton').style.display = 'block';
}

function displayResults(data) {
    const resultDiv = document.getElementById('result');
    let html = '<table><tr><th>Ürün</th><th>Toplam Miktar</th><th>Toplam Tutar</th></tr>';
    
    data.forEach((item, index) => {
        html += `<tr${index === data.length - 1 ? ' class="total-row"' : ''}>
            <td>${item.Urun}</td>
            <td>${item.ToplamMiktar}</td>
            <td>${item.ToplamTutar.toFixed(2)} ₺</td>
        </tr>`;
    });
    
    html += '</table>';
    resultDiv.innerHTML = html;
}

function downloadExcel() {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(summaryData.slice(0, -1)); // Genel toplamı hariç tut

    // Başlık stilini ayarla
    const headerStyle = { font: { bold: true } };
    ['A1', 'B1', 'C1'].forEach(cell => {
        if (!ws[cell]) ws[cell] = {};
        ws[cell].s = headerStyle;
    });

    // Para birimi formatını uygula
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const address = XLSX.utils.encode_col(C) + "1";
        if (ws[address].v === 'ToplamTutar') {
            for (let R = range.s.r + 1; R <= range.e.r; ++R) {
                const cell = XLSX.utils.encode_cell({c:C, r:R});
                if (!ws[cell]) continue;
                ws[cell].z = '#,##0.00 ₺';
                ws[cell].v = parseFloat(ws[cell].v).toFixed(2);
            }
            break;
        }
    }

    // Genel toplamı ekle
    const totalRow = summaryData[summaryData.length - 1];
    XLSX.utils.sheet_add_json(ws, [totalRow], {
        origin: -1,
        skipHeader: true
    });

    // Genel toplam satırını kalın yap
    const lastRowNum = range.e.r + 2; // +2 çünkü bir satır ekledik ve sıfır tabanlı indeksleme
    ['A', 'B', 'C'].forEach(col => {
        const cell = ws[col + lastRowNum];
        if (!cell) return;
        if (!cell.s) cell.s = {};
        cell.s.font = { bold: true };
    });

    XLSX.utils.book_append_sheet(wb, ws, "Özet");
    XLSX.writeFile(wb, 'ozet-rapor.xlsx');
}