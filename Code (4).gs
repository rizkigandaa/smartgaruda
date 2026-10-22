/** RUANG KELAS – backend Google Apps Script (data tersimpan di Google Sheet ini) */
const KODE_GURU = 'GURU2026';          // ganti dengan kode rahasia sekolah Anda
const SALT = 'ganti-teks-rahasia-ini'; // ganti sebelum dipakai
const MAPEL = ['Matematika', 'B. Indonesia', 'IPA', 'B. Inggris'];
const QS = ['Di kota mana Anda lahir?', 'Siapa nama hewan peliharaan pertama Anda?', 'Apa nama sekolah dasar Anda?'];
const HEAD = {
  Akun: ['Username', 'Nama', 'Peran', 'SiswaID', 'Hash', 'Q', 'AHash', 'Dibuat', 'Sekolah'],
  Sekolah: ['NPSN', 'Nama', 'Jenjang', 'Provinsi', 'KabKota', 'Dibuat'],
  Siswa: ['ID', 'Nama', 'Username', 'Matematika', 'B. Indonesia', 'IPA', 'B. Inggris'],
  Tugas: ['ID', 'Judul', 'Mapel', 'Tenggat', 'Selesai'],
  Info: ['ID', 'Teks', 'Tanggal'],
  Pesan: ['ID', 'Dari', 'Username', 'SiswaID', 'Teks', 'Tanggal'],
  Absensi: ['Tanggal', 'SiswaID', 'Status']
};

function doGet() {
  Object.keys(HEAD).forEach(sheet_); // membuat semua sheet otomatis
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('GarudaSmart')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- util ---------- */
const norm_ = s => String(s || '').trim().toLowerCase();
const today_ = () => Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
const txt_ = v => v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v);
const hash_ = s => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SALT + s));
function sheet_(n) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(HEAD[n]); s.setFrozenRows(1); s.getRange('A:A').setNumberFormat('@'); }
  else if (s.getLastColumn() < HEAD[n].length) s.getRange(1, 1, 1, HEAD[n].length).setValues([HEAD[n]]);
  return s;
}
function all_(n) {
  const v = sheet_(n).getDataRange().getValues(), h = v.shift();
  return v.map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; });
}
function lock_(fn) {
  const l = LockService.getScriptLock(); l.waitLock(15000);
  try { return fn(); } finally { l.releaseLock(); }
}
function tries_(key) {
  const c = CacheService.getScriptCache(), n = +(c.get(key) || 0);
  if (n >= 5) throw new Error('Terlalu banyak percobaan. Coba lagi dalam 15 menit.');
  return () => c.put(key, n + 1, 900);
}
function auth_(token, role) {
  const raw = CacheService.getScriptCache().get('s:' + token);
  if (!raw) throw new Error('Sesi berakhir. Silakan masuk lagi.');
  const me = JSON.parse(raw);
  if (role && me.role !== role) throw new Error('Akses ditolak untuk peran Anda.');
  return me;
}

/* ---------- akun: daftar, masuk, lupa sandi ---------- */
function register(p) {
  return lock_(() => {
    const u = norm_(p.u), nama = String(p.nama || '').trim(), a = norm_(p.a);
    if (!/^[a-z0-9_.]{3,}$/.test(u)) throw new Error('Nama pengguna minimal 3 karakter: huruf kecil, angka, titik, atau garis bawah.');
    if (nama.length < 3) throw new Error('Isi nama lengkap dengan benar.');
    if (String(p.pw || '').length < 6) throw new Error('Kata sandi minimal 6 karakter.');
    if (!a) throw new Error('Isi jawaban pemulihan.');
    if (all_('Akun').some(x => x.Username === u)) throw new Error('Nama pengguna sudah dipakai. Pilih yang lain.');
    const sk = String(p.sekolah || '').trim();
    if (sk.length < 3) throw new Error('Isi nama sekolah atau kampus Anda.');
    let sid = '';
    if (p.role === 'guru') {
      if (p.kode !== KODE_GURU) throw new Error('Kode guru tidak benar.');
    } else if (p.role === 'siswa') {
      sid = Utilities.getUuid().slice(0, 8);
      sheet_('Siswa').appendRow([sid, nama, u, 0, 0, 0, 0]); // nama siswa = nama yang mendaftar
    } else if (p.role === 'ortu') {
      const m = all_('Siswa').filter(s => norm_(s.Nama) === norm_(p.anak));
      if (m.length !== 1) throw new Error(m.length ? 'Ada lebih dari satu siswa dengan nama itu. Hubungi guru.' : 'Siswa belum terdaftar. Minta anak Anda mendaftar lebih dulu.');
      sid = m[0].ID;
    } else throw new Error('Peran tidak valid.');
    ensureSekolah_(sk, p.jenjang);
    sheet_('Akun').appendRow([u, nama, p.role, sid, hash_(p.pw), +p.q || 0, hash_(a), new Date(), sk]);
    return true;
  });
}
function login(u, pw) {
  u = norm_(u); const fail = tries_('l:' + u);
  const a = all_('Akun').find(x => x.Username === u);
  if (!a || a.Hash !== hash_(pw)) { fail(); throw new Error('Nama pengguna atau kata sandi tidak cocok.'); }
  const me = { u, nama: a.Nama, role: a.Peran, sid: String(a.SiswaID || ''), sekolah: String(a.Sekolah || '') }, token = Utilities.getUuid();
  CacheService.getScriptCache().put('s:' + token, JSON.stringify(me), 21600);
  return { token, me };
}
function recoveryQuestion(u) {
  const a = all_('Akun').find(x => x.Username === norm_(u));
  if (!a) throw new Error('Nama pengguna tidak ditemukan.');
  return QS[+a.Q] || QS[0];
}
function resetPassword(u, jawaban, pw) {
  u = norm_(u); const fail = tries_('r:' + u);
  const a = all_('Akun').find(x => x.Username === u);
  if (!a || a.AHash !== hash_(norm_(jawaban))) { fail(); throw new Error('Jawaban pemulihan tidak cocok.'); }
  if (String(pw || '').length < 6) throw new Error('Kata sandi minimal 6 karakter.');
  sheet_('Akun').getRange(a._row, 5).setValue(hash_(pw));
  return true;
}

/* ---------- data ---------- */
function getData(token) {
  const me = auth_(token), t = today_(), ab = all_('Absensi');
  let ss = all_('Siswa');
  if (me.role !== 'guru') ss = ss.filter(s => String(s.ID) === me.sid);
  const students = ss.map(s => {
    const mine = ab.filter(x => String(x.SiswaID) === String(s.ID)), c = st => mine.filter(x => x.Status === st).length;
    const td = mine.find(x => txt_(x.Tanggal) === t);
    return { id: String(s.ID), nama: s.Nama, nilai: MAPEL.map(m => +s[m] || 0), att: { H: c('H'), S: c('S'), I: c('I'), A: c('A') }, today: td ? td.Status : '' };
  });
  const tugas = all_('Tugas').map(x => ({ id: String(x.ID), judul: x.Judul, mapel: x.Mapel, tenggat: txt_(x.Tenggat), selesai: String(x.Selesai || '').split(',').filter(Boolean) }));
  const info = all_('Info').map(x => ({ id: String(x.ID), teks: x.Teks, tanggal: txt_(x.Tanggal) })).reverse();
  let pesan = [];
  if (me.role !== 'siswa') pesan = all_('Pesan').filter(x => me.role === 'guru' || x.Username === me.u).map(x => ({ dari: x.Dari, teks: x.Teks, tanggal: txt_(x.Tanggal) })).reverse();
  return { me, students, tugas, info, pesan };
}
function setNilai(token, id, i, v) {
  auth_(token, 'guru'); i = +i; v = Math.max(0, Math.min(100, +v || 0));
  return lock_(() => { const s = all_('Siswa').find(x => String(x.ID) === String(id)); if (!s || i < 0 || i > 3) throw new Error('Data siswa tidak ditemukan.'); sheet_('Siswa').getRange(s._row, 4 + i).setValue(v); return v; });
}
function setHadir(token, id, st) {
  auth_(token, 'guru'); if ('HSIA'.indexOf(st) < 0 || !st) throw new Error('Status tidak valid.');
  return lock_(() => {
    const t = today_(), r = all_('Absensi').find(x => String(x.SiswaID) === String(id) && txt_(x.Tanggal) === t);
    if (r) sheet_('Absensi').getRange(r._row, 3).setValue(st); else sheet_('Absensi').appendRow([t, id, st]);
    return true;
  });
}
function addTugas(token, judul, mapel, tenggat) {
  auth_(token, 'guru'); if (!String(judul).trim() || !tenggat) throw new Error('Isi judul dan tenggat tugas.');
  sheet_('Tugas').appendRow([Utilities.getUuid().slice(0, 8), String(judul).trim(), mapel, tenggat, '']); return true;
}
function addInfo(token, teks) {
  auth_(token, 'guru'); if (!String(teks).trim()) throw new Error('Isi pengumuman dulu.');
  sheet_('Info').appendRow([Utilities.getUuid().slice(0, 8), String(teks).trim(), today_()]); return true;
}
function toggleTugas(token, tid) {
  const me = auth_(token, 'siswa');
  return lock_(() => {
    const t = all_('Tugas').find(x => String(x.ID) === String(tid)); if (!t) throw new Error('Tugas tidak ditemukan.');
    const l = String(t.Selesai || '').split(',').filter(Boolean), i = l.indexOf(me.sid);
    i < 0 ? l.push(me.sid) : l.splice(i, 1);
    sheet_('Tugas').getRange(t._row, 5).setValue(l.join(',')); return true;
  });
}
function kirimPesan(token, teks) {
  const me = auth_(token, 'ortu'); if (!String(teks).trim()) throw new Error('Tulis pesan dulu.');
  sheet_('Pesan').appendRow([Utilities.getUuid().slice(0, 8), me.nama, me.u, me.sid, String(teks).trim(), today_()]); return true;
}

/* ---------- direktori sekolah & kampus ---------- */
// Sheet "Sekolah": NPSN | Nama | Jenjang | Provinsi | KabKota | Dibuat
// Jenjang memakai 4 label ini: "SD / MI", "SMP / MTs", "SMA / SMK / MA", "Perguruan tinggi"
function infoPublik() { return { sekolah: Math.max(0, sheet_('Sekolah').getLastRow() - 1) }; }
function cariSekolah(q, jenjang) {
  q = String(q || '').trim(); if (q.length < 3) return [];
  const sh = sheet_('Sekolah'), n = sh.getLastRow() - 1; if (n < 1) return [];
  const hits = sh.getRange(2, 2, n, 1).createTextFinder(q).matchCase(false).findAll().slice(0, 60), out = [];
  for (const h of hits) {
    const r = sh.getRange(h.getRow(), 1, 1, 5).getValues()[0];
    if (jenjang && r[2] && r[2] !== jenjang) continue;
    out.push({ npsn: String(r[0]), nama: r[1], kab: r[4], prov: r[3] });
    if (out.length >= 15) break;
  }
  return out;
}
function ensureSekolah_(nama, jenjang) {
  const sh = sheet_('Sekolah'), n = sh.getLastRow() - 1;
  if (n > 0 && sh.getRange(2, 2, n, 1).createTextFinder(nama).matchCase(false).matchEntireCell(true).findNext()) return;
  sh.appendRow(['', nama, jenjang || '', '', '', new Date()]); // nama baru dari pendaftar
}
// Impor/perbarui direktori dari file CSV (kolom: NPSN, Nama, Jenjang, Provinsi, KabKota).
// Isi properti skrip SEKOLAH_CSV_URL (Project Settings), lalu jalankan sekali atau pasang pembaruan harian.
function perbaruiSekolah() {
  const url = PropertiesService.getScriptProperties().getProperty('SEKOLAH_CSV_URL');
  if (!url) throw new Error('Isi properti skrip SEKOLAH_CSV_URL dengan tautan CSV data sekolah.');
  const rows = Utilities.parseCsv(UrlFetchApp.fetch(url).getContentText());
  const head = rows.shift().map(norm_), cols = ['npsn', 'nama', 'jenjang', 'provinsi', 'kabkota'].map(k => head.indexOf(k));
  if (cols[1] < 0 || cols[2] < 0) throw new Error('CSV harus punya kolom: NPSN, Nama, Jenjang, Provinsi, KabKota.');
  const now = new Date(), data = rows.filter(r => r[cols[1]]).map(r => cols.map(i => i < 0 ? '' : r[i]).concat([now]));
  const sh = sheet_('Sekolah'), last = sh.getLastRow();
  const manual = last > 1 ? sh.getRange(2, 1, last - 1, 6).getValues().filter(r => !r[0]) : [];
  if (last > 1) sh.getRange(2, 1, last - 1, 6).clearContent();
  const all = data.concat(manual);
  if (sh.getMaxRows() < all.length + 1) sh.insertRowsAfter(sh.getMaxRows(), all.length + 1 - sh.getMaxRows());
  for (let i = 0; i < all.length; i += 5000) { const c = all.slice(i, i + 5000); sh.getRange(2 + i, 1, c.length, 6).setValues(c); }
}
function pasangPembaruanHarian() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'perbaruiSekolah').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('perbaruiSekolah').timeBased().everyDays(1).atHour(2).create();
}
