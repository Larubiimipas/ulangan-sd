/**
 * Server Ulangan Online (Google Apps Script + Google Sheets).
 * Sheet yang dipakai (dibuat otomatis):
 *   Nilai      : nilai akhir setiap siswa
 *   Aktivitas  : catatan mulai / keluar / kembali / selesai
 *   Pantauan   : status terkini setiap siswa yang sedang ulangan
 * Kata sandi guru disimpan di Project Settings → Script properties → ADMIN_PASSWORD
 * (tidak ditulis di kode ini).
 */
const KOLOM = {
  Nilai: ['waktu', 'kelas', 'mapel', 'nama', 'benar', 'total', 'nilai', 'durasi', 'keluar', 'alasan', 'sesi'],
  Aktivitas: ['waktu', 'kelas', 'mapel', 'nama', 'jenis', 'keluar', 'terjawab', 'keterangan', 'sesi'],
  Pantauan: ['sesi', 'kelas', 'mapel', 'nama', 'status', 'terjawab', 'total', 'keluar', 'mulai', 'terakhir']
};

function sheet_(nama) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nama);
  if (!sh) {
    sh = ss.insertSheet(nama);
    sh.appendRow(KOLOM[nama]);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < KOLOM[nama].length) {
    sh.getRange(1, 1, 1, KOLOM[nama].length).setValues([KOLOM[nama]]);
  }
  return sh;
}

function rows_(nama) {
  const sh = sheet_(nama);
  const v = sh.getDataRange().getValues();
  const head = KOLOM[nama];
  return v.slice(1).filter(r => r.join('') !== '').map(r => {
    const o = {};
    head.forEach((k, i) => { o[k] = r[i] instanceof Date ? r[i].toISOString() : r[i]; });
    return o;
  });
}

const teks_ = (s, n) => String(s == null ? '' : s).trim().replace(/\s+/g, ' ').slice(0, n || 60);
const angka_ = x => Number(x) || 0;

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'data tidak valid' }); }
  const aksi = d.aksi || 'nilai';
  if (aksi === 'admin') return admin_(d);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const nama = teks_(d.nama, 40);
    if (!nama) throw new Error('nama kosong');
    const now = new Date();
    if (aksi === 'nilai') {
      sheet_('Nilai').appendRow([
        new Date(d.waktu || now), teks_(d.kelas, 3), teks_(d.mapel, 10), nama,
        angka_(d.benar), angka_(d.total), angka_(d.nilai), angka_(d.durasi),
        angka_(d.keluar), teks_(d.alasan, 40), teks_(d.sesi, 40)
      ]);
      upsertPantauan_(d, nama, 'selesai', now);
      log_(d, nama, 'selesai', now);
    } else if (aksi === 'log') {
      log_(d, nama, teks_(d.jenis, 20), now);
      upsertPantauan_(d, nama, d.jenis === 'kembali' || d.jenis === 'progres' ? 'mengerjakan' : teks_(d.jenis, 20), now);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function log_(d, nama, jenis, now) {
  if (jenis === 'progres') return; // progres hanya memperbarui Pantauan
  sheet_('Aktivitas').appendRow([
    now, teks_(d.kelas, 3), teks_(d.mapel, 10), nama, jenis,
    angka_(d.keluar), angka_(d.terjawab), teks_(d.keterangan, 80), teks_(d.sesi, 40)
  ]);
}

function upsertPantauan_(d, nama, status, now) {
  const sesi = teks_(d.sesi, 40);
  if (!sesi) return;
  const sh = sheet_('Pantauan');
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const i = ids.indexOf(sesi);
  const row = [sesi, teks_(d.kelas, 3), teks_(d.mapel, 10), nama, status,
    angka_(d.terjawab), angka_(d.total), angka_(d.keluar), d.mulai ? new Date(d.mulai) : now, now];
  if (i < 0) sh.appendRow(row);
  else {
    if (!d.mulai) row[8] = sh.getRange(i + 2, 9).getValue();
    sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
  }
}

function admin_(d) {
  const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!pw) return json_({ ok: false, error: 'Kata sandi guru belum diatur di Script properties (ADMIN_PASSWORD).' });
  if (String(d.password || '') !== pw) {
    Utilities.sleep(800);
    return json_({ ok: false, error: 'Kata sandi salah.' });
  }
  const akt = rows_('Aktivitas');
  return json_({
    ok: true,
    nilai: rows_('Nilai'),
    pantauan: rows_('Pantauan'),
    aktivitas: akt.slice(-1000),
    sekarang: new Date().toISOString()
  });
}

// Data ranking untuk siswa (?kelas=5) — hanya kolom yang perlu ditampilkan
function doGet(e) {
  const kelas = e && e.parameter && e.parameter.kelas;
  const data = rows_('Nilai')
    .filter(r => r.nama && (!kelas || String(r.kelas) === String(kelas)))
    .map(r => ({ waktu: r.waktu, kelas: String(r.kelas), mapel: String(r.mapel), nama: String(r.nama),
      benar: Number(r.benar), total: Number(r.total), nilai: Number(r.nilai), durasi: Number(r.durasi) }));
  return json_(data);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
