/**
 * Server Ulangan Online (Google Apps Script + Google Sheets).
 * Sheet yang dipakai (dibuat otomatis):
 *   Nilai      : nilai akhir setiap siswa
 *   Aktivitas  : catatan mulai / keluar / kembali / selesai
 *   Pantauan   : status terkini setiap siswa yang sedang ulangan
 * Kata sandi guru dibuat & diganti dari web (menu "Masuk sebagai guru").
 * Disimpan sebagai hash di Script properties (ADMIN_HASH, ADMIN_SALT).
 * Lupa kata sandi? Pemilik script menghapus ADMIN_HASH di Project Settings →
 * Script properties, lalu guru membuat kata sandi baru dari web.
 */
const KOLOM = {
  Nilai: ['waktu', 'kelas', 'mapel', 'nama', 'benar', 'total', 'nilai', 'durasi', 'keluar', 'alasan', 'sesi', 'kode'],
  Aktivitas: ['waktu', 'kelas', 'mapel', 'nama', 'jenis', 'keluar', 'terjawab', 'keterangan', 'sesi', 'kode'],
  Pantauan: ['sesi', 'kelas', 'mapel', 'nama', 'status', 'terjawab', 'total', 'keluar', 'mulai', 'terakhir', 'kode'],
  Ulangan: ['kode', 'nama', 'kelas', 'mapel', 'durasi', 'aktif', 'dibuat']
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
const kode_ = x => String(x == null ? '' : x).toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 8);

// ---------- Ulangan & kode ----------
function cariUlangan_(kode) {
  kode = kode_(kode);
  return rows_('Ulangan').find(u => String(u.kode) === kode) || null;
}
function cekKode_(d) {
  const u = cariUlangan_(d.kode);
  if (!u) return json_({ ok: false, error: 'Kode tidak ditemukan. Periksa lagi kode dari gurumu.' });
  if (!(u.aktif === true || String(u.aktif).toUpperCase() === 'TRUE')) return json_({ ok: false, error: 'Ulangan ini sudah ditutup oleh guru.' });
  if (d.kelas && String(u.kelas) !== String(d.kelas)) return json_({ ok: false, error: 'Kode ini untuk Kelas ' + u.kelas + ', bukan kelasmu.' });
  if (d.mapel && String(u.mapel) !== String(d.mapel)) return json_({ ok: false, error: 'Kode ini bukan untuk mata pelajaran yang kamu pilih.' });
  return json_({ ok: true, ulangan: { kode: String(u.kode), nama: String(u.nama), kelas: String(u.kelas), mapel: String(u.mapel), durasi: Number(u.durasi) || 0 } });
}
function kodeBaru_() {
  const huruf = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // tanpa O/0/I/1 agar tidak tertukar
  const ada = rows_('Ulangan').map(u => String(u.kode));
  let k;
  do { k = ''; for (let i = 0; i < 6; i++) k += huruf[Math.floor(Math.random() * huruf.length)]; } while (ada.indexOf(k) >= 0);
  return k;
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'data tidak valid' }); }
  const aksi = d.aksi || 'nilai';
  if (aksi === 'admin') return admin_(d);
  if (aksi === 'cekKode') return cekKode_(d);

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
        angka_(d.keluar), teks_(d.alasan, 40), teks_(d.sesi, 40), kode_(d.kode)
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
    angka_(d.keluar), angka_(d.terjawab), teks_(d.keterangan, 80), teks_(d.sesi, 40), kode_(d.kode)
  ]);
}

function upsertPantauan_(d, nama, status, now) {
  const sesi = teks_(d.sesi, 40);
  if (!sesi) return;
  const sh = sheet_('Pantauan');
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const i = ids.indexOf(sesi);
  const row = [sesi, teks_(d.kelas, 3), teks_(d.mapel, 10), nama, status,
    angka_(d.terjawab), angka_(d.total), angka_(d.keluar), d.mulai ? new Date(d.mulai) : now, now, kode_(d.kode)];
  if (i < 0) sh.appendRow(row);
  else {
    if (!d.mulai) row[8] = sh.getRange(i + 2, 9).getValue();
    sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
  }
}

// ---------- Kata sandi guru (disimpan sebagai hash, bukan teks asli) ----------
function hash_(pw, salt) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return b.map(x => ('0' + (x & 255).toString(16)).slice(-2)).join('');
}
function sandiAda_() {
  const p = PropertiesService.getScriptProperties();
  return !!(p.getProperty('ADMIN_HASH') || p.getProperty('ADMIN_PASSWORD'));
}
function sandiCocok_(pw) {
  const p = PropertiesService.getScriptProperties();
  const h = p.getProperty('ADMIN_HASH'), salt = p.getProperty('ADMIN_SALT');
  if (h && salt) return hash_(String(pw || ''), salt) === h;
  const lama = p.getProperty('ADMIN_PASSWORD'); // kompatibel dengan cara lama
  return !!lama && String(pw || '') === lama;
}
function simpanSandi_(pw) {
  const salt = Utilities.getUuid();
  const p = PropertiesService.getScriptProperties();
  p.setProperties({ ADMIN_SALT: salt, ADMIN_HASH: hash_(pw, salt) });
  p.deleteProperty('ADMIN_PASSWORD');
}

function admin_(d) {
  const sub = d.sub || 'data';
  if (sub === 'cek') return json_({ ok: true, ada: sandiAda_() });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (sub === 'buat') {
      if (sandiAda_()) return json_({ ok: false, error: 'Kata sandi guru sudah pernah dibuat. Silakan masuk.' });
      if (String(d.baru || '').length < 6) return json_({ ok: false, error: 'Kata sandi minimal 6 karakter.' });
      simpanSandi_(String(d.baru));
      return json_({ ok: true });
    }
    if (!sandiAda_()) return json_({ ok: false, error: 'Kata sandi guru belum dibuat.', belum: true });
    if (!sandiCocok_(d.password)) {
      Utilities.sleep(800);
      return json_({ ok: false, error: sub === 'ganti' ? 'Kata sandi lama salah.' : 'Kata sandi salah.' });
    }
    if (sub === 'buatUlangan') {
      const nama = teks_(d.nama, 60), kelas = teks_(d.kelas, 3), mapel = teks_(d.mapel, 10);
      if (!nama || !kelas || !mapel) return json_({ ok: false, error: 'Nama ulangan, kelas, dan mapel wajib diisi.' });
      const durasi = d.tanpaWaktu ? 0 : Math.min(240, Math.max(5, angka_(d.durasi) || 60));
      const kode = kodeBaru_();
      sheet_('Ulangan').appendRow([kode, nama, kelas, mapel, durasi, true, new Date()]);
      return json_({ ok: true, kode: kode });
    }
    if (sub === 'aturUlangan') {
      const sh = sheet_('Ulangan'), kode = kode_(d.kode);
      const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
      const i = ids.indexOf(kode);
      if (i < 0) return json_({ ok: false, error: 'Ulangan tidak ditemukan.' });
      sh.getRange(i + 2, 6).setValue(!!d.aktif);
      return json_({ ok: true });
    }
    if (sub === 'ganti') {
      if (String(d.baru || '').length < 6) return json_({ ok: false, error: 'Kata sandi baru minimal 6 karakter.' });
      simpanSandi_(String(d.baru));
      return json_({ ok: true });
    }
  } finally {
    lock.releaseLock();
  }
  const akt = rows_('Aktivitas');
  return json_({
    ok: true,
    ulangan: rows_('Ulangan'),
    nilai: rows_('Nilai'),
    pantauan: rows_('Pantauan'),
    aktivitas: akt.slice(-1000),
    sekarang: new Date().toISOString()
  });
}

// Data ranking untuk siswa (?kelas=5) — hanya kolom yang perlu ditampilkan
function doGet(e) {
  const p = (e && e.parameter) || {};
  const kodes = String(p.kode || '').split(',').map(kode_).filter(String);
  const data = rows_('Nilai')
    .filter(r => r.nama && (!p.kelas || String(r.kelas) === String(p.kelas)) && (!kodes.length || kodes.indexOf(String(r.kode)) >= 0))
    .map(r => ({ waktu: r.waktu, kelas: String(r.kelas), mapel: String(r.mapel), nama: String(r.nama),
      benar: Number(r.benar), total: Number(r.total), nilai: Number(r.nilai), durasi: Number(r.durasi), kode: String(r.kode || '') }));
  return json_(data);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
