// ============================================
// KODE GOOGLE APPS SCRIPT - VERSI SEDERHANA
// ============================================
// 
// CARA PAKAI:
// 1. Buka Google Sheets baru (kosong)
// 2. Extensions > Apps Script
// 3. Hapus semua kode, tempel kode ini
// 4. Save
// 5. Deploy > New deployment > Web app
// 6. Set "Who has access" = "Anyone"
// 7. Deploy
// 8. Salin URL yang diberikan
// 9. Ganti URL di form website dengan URL tersebut
//
// ============================================

function doGet(e) {
  try {
    // Buka spreadsheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Buat header jika belum ada
    if (sheet.getRange(1, 1).getValue() !== 'Timestamp') {
      sheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Nama', 'NIM', 'Email', 'Pesan/Aspirasi']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      sheet.getRange(1, 1, 1, 5).setBackground('#f0f0f0');
      
      // Set border untuk header
      const headerRange = sheet.getRange(1, 1, 1, 5);
      headerRange.setBorder(true, true, true, true, true, true);
    }
    
    // Ambil data dari form
    const nama = e.parameter.nama || '';
    const nim = e.parameter.nim || '';
    const email = e.parameter.email || '';
    const pesan = e.parameter.pesan || '';
    
    // Simpan ke spreadsheet
    const waktu = new Date();
    const lastRow = sheet.getLastRow() + 1;
    sheet.appendRow([waktu, nama, nim, email, pesan]);
    
    // Set border untuk baris baru
    const newRowRange = sheet.getRange(lastRow, 1, 1, 5);
    newRowRange.setBorder(true, true, true, true, true, true);
    
    // Auto-resize kolom berdasarkan konten
    sheet.autoResizeColumns(1, 5);
    
    // Set lebar kolom minimum dan maksimum
    sheet.setColumnWidth(1, 150); // Timestamp: 150px
    sheet.setColumnWidth(2, 200); // Nama: 200px
    sheet.setColumnWidth(3, 120); // NIM: 120px
    sheet.setColumnWidth(4, 250); // Email: 250px
    sheet.setColumnWidth(5, 400); // Pesan: 400px (lebih lebar untuk pesan panjang)
    
    // Wrap text untuk kolom pesan agar teks panjang bisa dibaca
    sheet.getRange(lastRow, 5).setWrap(true);
    
    // Set border untuk semua data yang ada (jika belum ada)
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() > 1) {
      const allDataRange = sheet.getRange(2, 1, dataRange.getNumRows() - 1, 5);
      allDataRange.setBorder(true, true, true, true, true, true);
    }
    
    // Return pesan sukses
    return ContentService.createTextOutput('Berhasil!');
    
  } catch (error) {
    return ContentService.createTextOutput('Error: ' + error.toString());
  }
}

// ============================================
// SELESAI - HANYA INI SAJA!
// ============================================

