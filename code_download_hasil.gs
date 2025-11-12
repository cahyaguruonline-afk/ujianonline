// [GANTI SELURUH FUNGSI getDownloadLinkXLS DI code.gs.txt]

function getDownloadLinkXLS(sessionId) {
  const session = getUserSession(sessionId);
  if (!session.success || !session.data || session.data.role !== 'admin') {
    return { success: false, message: 'Akses ditolak.' };
  }

  try {
    // 1. Buka Spreadsheet (Cek SPREADSHEET_ID)
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); 
    Logger.log('Spreadsheet berhasil dibuka dengan ID: ' + SPREADSHEET_ID);
    
    // 2. Dapatkan Sheet 'peserta' (Cek Nama Sheet)
    // PASTIKAN NAMA INI SAMA PERSIS dengan sheet di Spreadsheet Anda
    const SHEET_NAME = 'peserta'; // Ubah 'peserta' menjadi 'Peserta' jika P kapital
    const sheet = ss.getSheetByName(SHEET_NAME); 
    
    if (!sheet) {
      Logger.log('ERROR: Sheet "' + SHEET_NAME + '" tidak ditemukan.');
      return { success: false, message: 'Sheet "' + SHEET_NAME + '" tidak ditemukan. Pastikan nama sheet benar (case-sensitive).' };
    }

    // 3. Dapatkan Sheet ID (gid)
    const sheetId = sheet.getSheetId(); 
    Logger.log('Sheet "' + SHEET_NAME + '" ditemukan. Sheet ID (gid): ' + sheetId);
    
    // 4. Buat URL ekspor .xlsx
    // Ini adalah format URL standar untuk download .xlsx dari Google Sheets
    const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=xlsx&gid=${sheetId}`;
    
    Logger.log('Generated URL: ' + url);
    
    // 5. Kembalikan URL
    return { success: true, url: url };

  } catch (e) {
    // Tangkap error jika ada masalah saat membuka SS atau Sheet
    Logger.log('Error in getDownloadLinkXLS: ' + e.message + ' Stack: ' + e.stack);
    return { success: false, message: 'Terjadi kesalahan server saat membuat link: ' + e.message };
  }
}
