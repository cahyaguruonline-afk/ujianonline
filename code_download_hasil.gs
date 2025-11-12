// [TAMBAHKAN FUNGSI INI KE code.gs.txt]

/**
 * Mendapatkan link download .xlsx untuk sheet 'peserta'.
 * Dipanggil oleh admin dari halaman Hasil Ujian.
 */
function getDownloadLinkXLS(sessionId) {
  // 1. Validasi Sesi Admin
  const session = getUserSession(sessionId);
  if (!session.success || !session.data || session.data.role !== 'admin') {
    return { success: false, message: 'Akses ditolak.' };
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID); [cite_start]// [cite: 537]
    const sheet = ss.getSheetByName('peserta'); [cite_start]// [cite: 593, 642, 677, 753]
    
    if (!sheet) {
      return { success: false, message: 'Sheet "peserta" tidak ditemukan.' }; [cite_start]// [cite: 754, 786]
    }

    const sheetId = sheet.getSheetId(); // Dapatkan 'gid' (Grid ID)
    
    // 2. Buat URL ekspor .xlsx
    const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=xlsx&gid=${sheetId}`;
    
    // 3. Kembalikan URL ke frontend
    return { success: true, url: url };

  } catch (e) {
    Logger.log('Error in getDownloadLinkXLS: ' + e.message);
    return { success: false, message: 'Gagal membuat link download: ' + e.message };
  }
}
