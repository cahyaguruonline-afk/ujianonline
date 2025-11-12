// [TAMBAHKAN FUNGSI INI KE code.gs.txt]

/**
 * Mengambil detail pertanyaan dan jawaban esai untuk seorang peserta.
 * Dipanggil dari modal 'Lihat Esai' di halaman Rekap Admin.
 */
function getDetailEsaiPeserta(sessionId, noPeserta) {
  // 1. Validasi Sesi Admin
  const session = getUserSession(sessionId);
  if (!session.success || !session.data || session.data.role !== 'admin') {
    return { success: false, message: 'Akses ditolak.' };
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const rekapEsaiSheet = ss.getSheetByName('rekap_esai');
    if (!rekapEsaiSheet) {
      return { success: false, message: 'Sheet "rekap_esai" tidak ditemukan.' };
    }

    // 2. Siapkan Map untuk Soal (Mengambil teks pertanyaan)
    const allSoalDetails = {};
    const allSubjectSheets = getSubjectSheets().concat(['soal_esai']);
    
    allSubjectSheets.forEach(sheetName => {
      const sheetSoal = ss.getSheetByName(sheetName);
      if (sheetSoal) {
        const dataSoal = sheetSoal.getDataRange().getValues();
        allSoalDetails[sheetName] = {};
        for (let i = 1; i < dataSoal.length; i++) {
          if (dataSoal[i].join('').trim() === '') continue;
          const noSoal = parseInt(dataSoal[i][0], 10);
          if (isNaN(noSoal)) continue;
          
          if (dataSoal[i][1] === 'Esai') { // Hanya ambil soal Tipe Esai
            allSoalDetails[sheetName][noSoal] = dataSoal[i][2]; // Kolom [2] adalah 'Pertanyaan'
          }
        }
      }
    });

    // 3. Cari Jawaban Peserta di 'rekap_esai'
    const dataRekap = rekapEsaiSheet.getDataRange().getValues();
    const headers = dataRekap[0];
    const noPesertaCol = headers.indexOf('Nomor Peserta');
    const mapelCol = headers.indexOf('Mata Pelajaran');

    const hasilEsai = []; // Array untuk menyimpan {pertanyaan, jawaban}

    for (let i = 1; i < dataRekap.length; i++) {
      if (dataRekap[i][noPesertaCol] == noPeserta) {
        const row = dataRekap[i];
        const mapel = row[mapelCol];
        const soalSheetName = 'soal_' + mapel.toLowerCase().replace(/ /g, '_');

        for (let j = 4; j < headers.length; j++) { // Mulai dari kolom jawaban pertama (indeks 4)
          if (row[j]) { // Jika ada jawaban di kolom ini
            const soalNumMatch = headers[j].match(/\d+/);
            const soalNum = soalNumMatch ? parseInt(soalNumMatch[0], 10) : null;
            
            // Cocokkan nomor soal dengan pertanyaan
            if (soalNum && allSoalDetails[soalSheetName] && allSoalDetails[soalSheetName][soalNum]) {
              const pertanyaan = allSoalDetails[soalSheetName][soalNum];
              const jawaban = row[j];
              
              hasilEsai.push({
                mapel: mapel,
                pertanyaan: pertanyaan,
                jawaban: jawaban
              });
            }
          }
        }
      }
    }
    
    return { success: true, data: hasilEsai };

  } catch (e) {
    Logger.log('Error in getDetailEsaiPeserta: ' + e.message);
    return { success: false, message: 'Gagal mengambil detail esai: ' + e.message };
  }
}
