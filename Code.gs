/**
 * Penyimpan nilai & ranking untuk Ulangan Online.
 * Cara pasang:
 * 1. Buat Google Sheet baru → menu Ekstensi → Apps Script.
 * 2. Hapus isi Code.gs, tempel seluruh kode ini, lalu Simpan.
 * 3. Terapkan → Deployment baru → jenis "Aplikasi web".
 *    Jalankan sebagai: Saya. Yang memiliki akses: Siapa saja.
 * 4. Salin URL Web App (berakhiran /exec) ke CONFIG.SCRIPT_URL di index.html.
 */
const NAMA_SHEET = 'Nilai';
const KOLOM = ['waktu', 'kelas', 'mapel', 'nama', 'benar', 'total', 'nilai', 'durasi'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(NAMA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(NAMA_SHEET);
    sh.appendRow(KOLOM);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Menerima nilai dari siswa
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const nama = String(d.nama || '').trim().replace(/\s+/g, ' ').slice(0, 40);
    if (!nama) throw new Error('nama kosong');
    sheet_().appendRow([
      new Date(d.waktu || Date.now()), String(d.kelas), String(d.mapel), nama,
      Number(d.benar) || 0, Number(d.total) || 0, Number(d.nilai) || 0, Number(d.durasi) || 0
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Mengirim data nilai untuk ranking (?kelas=5)
function doGet(e) {
  const kelas = e && e.parameter && e.parameter.kelas;
  const rows = sheet_().getDataRange().getValues().slice(1);
  const data = rows
    .filter(r => r[3] && (!kelas || String(r[1]) === String(kelas)))
    .map(r => ({
      waktu: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      kelas: String(r[1]), mapel: String(r[2]), nama: String(r[3]),
      benar: Number(r[4]), total: Number(r[5]), nilai: Number(r[6]), durasi: Number(r[7])
    }));
  return json_(data);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
