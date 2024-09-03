const XLSX = require('xlsx');

function summarizeSales(inputFilePath, outputFilePath) {
  const workbook = XLSX.readFile(inputFilePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  
  const data = XLSX.utils.sheet_to_json(sheet, { header: 'A' });

  const summary = data.reduce((acc, row, index) => {
    if (index < 4) return acc; // Başlık satırlarını atla
    
    const Urun = row['N'];
    const Miktar = row['O'];
    const Tutar = row['S'];
    
    if (Urun && Miktar) {
      if (!acc[Urun]) {
        acc[Urun] = { toplamMiktar: 0, toplamTutar: 0 };
      }
      const miktar = Number(Miktar) || 0;
      const tutar = Number(Tutar) || 0;
      acc[Urun].toplamMiktar += miktar;
      acc[Urun].toplamTutar += tutar;
    }
    
    return acc;
  }, {});

  let genelToplamMiktar = 0;
  let genelToplamTutar = 0;
  const summaryArray = Object.entries(summary)
    .filter(([_, { toplamMiktar, toplamTutar }]) => toplamMiktar > 0 || toplamTutar > 0)
    .map(([urun, { toplamMiktar, toplamTutar }]) => {
      genelToplamMiktar += toplamMiktar;
      genelToplamTutar += toplamTutar;
      return {
        'Ürün': urun,
        'Toplam Miktar': toplamMiktar,
        'Toplam Tutar': toplamTutar
      };
    });

  // Genel toplam satırını ekle
  if (genelToplamMiktar > 0 || genelToplamTutar > 0) {
    summaryArray.push({
      'Ürün': 'GENEL TOPLAM',
      'Toplam Miktar': genelToplamMiktar,
      'Toplam Tutar': genelToplamTutar
    });
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(summaryArray);

  // Para birimi formatını uygula
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const address = XLSX.utils.encode_col(C) + "1";
    if (ws[address].v === 'Toplam Tutar') {
      for (let R = range.s.r + 1; R <= range.e.r; ++R) {
        const cell = XLSX.utils.encode_cell({c:C, r:R});
        if (!ws[cell]) continue;
        ws[cell].z = '#,##0.00 ₺';
      }
      break;
    }
  }

  // Genel toplam satırını kalın yap
  if (summaryArray.length > 0) {
    const lastRow = range.e.r;
    ['A', 'B', 'C'].forEach(col => {
      const cell = ws[`${col}${lastRow + 1}`];
      if (cell) {
        cell.s = { font: { bold: true } };
      }
    });
  }

  XLSX.utils.book_append_sheet(wb, ws, "Özet");
  XLSX.writeFile(wb, outputFilePath);

  console.log(`Özet Excel dosyası oluşturuldu: ${outputFilePath}`);
}

summarizeSales('excel-rapor2.xlsx', 'ozet-rapor2.xlsx');