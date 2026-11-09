/**
 * STOCK OPNAME GUDANG - backend Google Apps Script (terikat ke Google Sheet)
 * Sheet yang dipakai: Barang, Akun, Meta, Log, Laporan (dibuat otomatis).
 */

// ====== ISI SEBELUM MENJALANKAN setupAkunIT() ======
// Setelah setup berhasil, HAPUS password ini dari kode.
const AKUN_IT_AWAL = { username: 'admin', nama: 'Admin IT', password: 'GANTI-PASSWORD-INI' };
// ===================================================

const TZ = 'Asia/Jakarta';
const SESSION_TTL = 21600; // detik (batas maksimum CacheService = 6 jam)
const HASH_ITER = 2000;
// Catatan: di sheet kolom B berjudul "sku", tetapi di kode program/JSON propertinya tetap bernama "kode".
const H_BARANG = ['id', 'sku', 'nama', 'satuan', 'lokasi', 'sistem', 'fisik', 'catatan', 'updatedAt', 'updatedBy', 'barcode'];
const T_BARANG = [1, 2, 3, 4, 5, 8, 9, 10, 11]; // kolom bertipe teks (SKU & barcode tidak boleh jadi angka)
const H_AKUN = ['username', 'nama', 'role', 'salt', 'hash'];

// Backend ini hanya API (tanpa halaman). Halaman web (index.html) di-hosting di GitHub Pages
// dan memanggil API ini lewat POST JSON: {fn: "api_login", args: [...]}.
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'Stock Opname Gudang API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const API = {
    api_login: api_login, api_logout: api_logout, api_getData: api_getData, api_saveMeta: api_saveMeta,
    api_setFisik: api_setFisik, api_addItem: api_addItem, api_updateItem: api_updateItem,
    api_deleteItem: api_deleteItem, api_importItems: api_importItems, api_buatLaporan: api_buatLaporan,
    api_listUsers: api_listUsers, api_saveUser: api_saveUser, api_deleteUser: api_deleteUser
  };
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = Object.prototype.hasOwnProperty.call(API, req.fn) ? API[req.fn] : null;
    if (!fn) throw new Error('Perintah tidak dikenal.');
    out = { ok: true, data: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- SETUP (dijalankan sekali dari editor oleh IT) ---------- */
function setupAkunIT() {
  const a = AKUN_IT_AWAL;
  if (!a.password || a.password === 'GANTI-PASSWORD-INI' || a.password.length < 8) {
    throw new Error('Ubah password di AKUN_IT_AWAL dulu (minimal 8 karakter).');
  }
  PropertiesService.getScriptProperties().setProperty('SSID', SpreadsheetApp.getActiveSpreadsheet().getId());
  sheet_('Barang', H_BARANG, T_BARANG);
  const sh = sheet_('Akun', H_AKUN, [1, 2, 3, 4, 5]);
  sheet_('Meta', ['key', 'value'], [1, 2]);
  sheet_('Log', ['waktu', 'user', 'aksi', 'detail'], [1, 2, 3, 4]);
  const uname = a.username.trim().toLowerCase();
  if (sh.getDataRange().getValues().slice(1).some(r => String(r[0]) === uname)) throw new Error('Akun "' + uname + '" sudah ada.');
  appendRow_(sh, userRow_(uname, a.nama, 'it', a.password));
  log_('setup', 'Akun IT dibuat', uname);
  Logger.log('Setup selesai. Akun IT: ' + uname + '. Sekarang hapus password dari kode, lalu Deploy sebagai Web app.');
}

/* ---------- Impor akun dari daftar di Google Sheet (dijalankan IT dari editor) ----------
 * Buat sheet bernama "Akun_Import" dengan kolom: username | nama | role (IT / Gudang) | password
 * Jalankan fungsi ini. Password diubah menjadi hash di sheet "Akun", lalu kolom password
 * di Akun_Import diganti "(sudah diimpor)" supaya tidak tersimpan sebagai teks biasa. */
function importAkunDariSheet() {
  const ss = ss_();
  const src = ss.getSheetByName('Akun_Import');
  if (!src) throw new Error('Buat sheet "Akun_Import" dengan kolom: username, nama, role, password.');
  const rows = src.getDataRange().getValues().slice(1);
  const sh = sheet_('Akun', H_AKUN, [1, 2, 3, 4, 5]);
  const existing = readUsers_().map(u => u.username);
  let add = 0;
  const skip = [];
  rows.forEach((r, i) => {
    const uname = String(r[0] || '').trim().toLowerCase();
    if (!uname) return;
    const nama = String(r[1] || '').trim();
    const role = String(r[2] || '').trim().toLowerCase() === 'it' ? 'it' : 'gudang';
    const pw = String(r[3] === undefined || r[3] === null ? '' : r[3]);
    if (!/^[a-z0-9._-]{3,}$/.test(uname) || !nama || pw.length < 8 || existing.indexOf(uname) >= 0) { skip.push(uname); return; }
    appendRow_(sh, userRow_(uname, nama, role, pw));
    existing.push(uname);
    src.getRange(i + 2, 4).setValue('(sudah diimpor)');
    add++;
  });
  log_('setup', 'import akun', add + ' akun, ' + skip.length + ' dilewati');
  Logger.log('Selesai: ' + add + ' akun ditambahkan. Dilewati (kosong/duplikat/password < 8 karakter): ' + (skip.join(', ') || '-'));
}

/* ---------- Helper umum ---------- */
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SSID');
  if (!id) throw new Error('Setup belum dijalankan. Jalankan fungsi setupAkunIT di editor Apps Script.');
  return SpreadsheetApp.openById(id);
}
function sheet_(name, headers, textCols) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    (textCols || []).forEach(c => sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'));
  }
  if (headers && sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function appendRow_(sh, row) {
  const r = Math.max(sh.getLastRow(), 1) + 1;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100);
  sh.getRange(r, 1, 1, row.length).setValues([row]);
}
function withLock_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}
function log_(user, aksi, detail) {
  try { appendRow_(sheet_('Log', ['waktu', 'user', 'aksi', 'detail'], [1, 2, 3, 4]), [Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), user, aksi, detail || '']); } catch (e) {}
}
const todayStr_ = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

/* ---------- Akun, password, sesi ---------- */
function hash_(pw, salt) {
  let h = salt + ':' + pw;
  for (let i = 0; i < HASH_ITER; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + ':' + i));
  }
  return h;
}
function userRow_(username, nama, role, pw) {
  const salt = Utilities.getUuid();
  return [username, nama, role, salt, hash_(pw, salt)];
}
function readUsers_() {
  const sh = sheet_('Akun', H_AKUN, [1, 2, 3, 4, 5]);
  return sh.getDataRange().getValues().slice(1).filter(r => r[0] !== '').map(r => ({
    username: String(r[0]), nama: String(r[1]), role: String(r[2]), salt: String(r[3]), hash: String(r[4])
  }));
}
function auth_(token, roles) {
  const cache = CacheService.getScriptCache();
  const raw = token ? cache.get('s_' + token) : null;
  if (!raw) throw new Error('SESI_HABIS');
  const s = JSON.parse(raw);
  const rev = cache.get('r_' + s.username);
  if (rev && Number(rev) > s.iat) throw new Error('SESI_HABIS');
  if (roles && roles.indexOf(s.role) < 0) throw new Error('Anda tidak punya izin untuk tindakan ini.');
  return s;
}
function revoke_(username) {
  CacheService.getScriptCache().put('r_' + username, String(Date.now()), SESSION_TTL);
}

function api_login(username, password) {
  const u = String(username || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const fk = 'f_' + u;
  const fails = Number(cache.get(fk) || 0);
  if (fails >= 5) throw new Error('Terlalu banyak percobaan gagal. Coba lagi 5 menit lagi.');
  const r = readUsers_().find(x => x.username === u);
  if (!r || hash_(String(password || ''), r.salt) !== r.hash) {
    cache.put(fk, String(fails + 1), 300);
    Utilities.sleep(700);
    throw new Error('Username atau password salah.');
  }
  cache.remove(fk);
  const token = Utilities.getUuid() + Utilities.getUuid();
  const user = { username: r.username, nama: r.nama, role: r.role };
  cache.put('s_' + token, JSON.stringify({ iat: Date.now(), username: r.username, nama: r.nama, role: r.role }), SESSION_TTL);
  log_(r.username, 'login', '');
  return { token: token, user: user };
}
function api_logout(token) {
  CacheService.getScriptCache().remove('s_' + token);
  return true;
}

/* ---------- Meta ---------- */
function getMeta_() {
  const sh = sheet_('Meta', ['key', 'value'], [1, 2]);
  const m = { tanggal: todayStr_(), petugas: '' };
  sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0] === 'tanggal' && r[1]) m.tanggal = String(r[1]); if (r[0] === 'petugas') m.petugas = String(r[1]); });
  return m;
}
function api_saveMeta(token, meta) {
  const s = auth_(token);
  return withLock_(() => {
    const sh = sheet_('Meta', ['key', 'value'], [1, 2]);
    const vals = { tanggal: String(meta.tanggal || ''), petugas: String(meta.petugas || '').slice(0, 100) };
    const rows = sh.getDataRange().getValues();
    Object.keys(vals).forEach(k => {
      const i = rows.findIndex((r, idx) => idx > 0 && r[0] === k);
      if (i > 0) sh.getRange(i + 1, 2).setValue(vals[k]); else appendRow_(sh, [k, vals[k]]);
    });
    return true;
  });
}

/* ---------- Barang ---------- */
function barangSheet_() {
  const sh = sheet_('Barang', H_BARANG, T_BARANG);
  // Migrasi otomatis sheet lama: judul kolom B jadi "sku" dan tambah kolom K "barcode"
  const hdr = sh.getRange(1, 1, 1, H_BARANG.length).getValues()[0];
  if (hdr[1] !== 'sku' || hdr[10] !== 'barcode') {
    sh.getRange(1, 11, sh.getMaxRows(), 1).setNumberFormat('@');
    sh.getRange(1, 1, 1, H_BARANG.length).setValues([H_BARANG]).setFontWeight('bold');
  }
  return sh;
}
function readItems_() {
  const v = barangSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!r[0]) continue;
    out.push({
      id: String(r[0]), kode: String(r[1]), nama: String(r[2]), satuan: String(r[3]), lokasi: String(r[4]),
      sistem: Number(r[5]) || 0, fisik: (r[6] === '' || r[6] === null) ? null : Number(r[6]),
      catatan: String(r[7]), updatedAt: String(r[8]), updatedBy: String(r[9]),
      barcode: String(r[10] === undefined || r[10] === null ? '' : r[10])
    });
  }
  return out;
}
function itemRow_(it) {
  return [it.id, it.kode, it.nama, it.satuan || '', it.lokasi || '', Number(it.sistem) || 0,
    (it.fisik === null || it.fisik === undefined || it.fisik === '') ? '' : Number(it.fisik),
    it.catatan || '', it.updatedAt || '', it.updatedBy || '', it.barcode || ''];
}
function writeItems_(items) {
  const sh = barangSheet_();
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, H_BARANG.length).clearContent();
  if (!items.length) return;
  if (sh.getMaxRows() < items.length + 1) sh.insertRowsAfter(sh.getMaxRows(), items.length + 1 - sh.getMaxRows());
  T_BARANG.forEach(c => sh.getRange(2, c, items.length, 1).setNumberFormat('@'));
  sh.getRange(2, 1, items.length, H_BARANG.length).setValues(items.map(itemRow_));
}
function findRow_(sh, id) {
  const n = Math.max(sh.getLastRow() - 1, 1);
  const ids = sh.getRange(2, 1, n, 1).getValues();
  const i = ids.findIndex(r => String(r[0]) === String(id));
  if (i < 0) throw new Error('Barang tidak ditemukan (mungkin sudah dihapus). Muat ulang data.');
  return i + 2;
}
function cleanItem_(it) {
  const kode = String(it.kode || '').trim(), nama = String(it.nama || '').trim();
  if (!kode || !nama) throw new Error('SKU dan nama barang wajib diisi.');
  return { kode: kode.slice(0, 60), nama: nama.slice(0, 200), satuan: String(it.satuan || '').slice(0, 30),
    lokasi: String(it.lokasi || '').slice(0, 60), sistem: Number(it.sistem) || 0, catatan: String(it.catatan || '').slice(0, 300),
    barcode: String(it.barcode || '').trim().slice(0, 60) };
}

function api_getData(token) {
  const s = auth_(token);
  return { user: { username: s.username, nama: s.nama, role: s.role }, meta: getMeta_(), items: readItems_() };
}
function api_setFisik(token, id, val) {
  const s = auth_(token);
  const num = (val === null || val === '' || val === undefined) ? '' : Number(val);
  if (num !== '' && !isFinite(num)) throw new Error('Nilai stok fisik tidak valid.');
  return withLock_(() => {
    const sh = barangSheet_();
    const row = findRow_(sh, id);
    const now = new Date().toISOString();
    sh.getRange(row, 7).setValue(num);
    sh.getRange(row, 9, 1, 2).setValues([[now, s.username]]);
    return { updatedAt: now };
  });
}
function api_addItem(token, it) {
  const s = auth_(token);
  const c = cleanItem_(it);
  return withLock_(() => {
    const items = readItems_();
    if (items.some(x => x.kode.toLowerCase() === c.kode.toLowerCase())) throw new Error('SKU "' + c.kode + '" sudah ada.');
    if (c.barcode) {
      const b = items.find(x => x.barcode && x.barcode.toLowerCase() === c.barcode.toLowerCase());
      if (b) throw new Error('Barcode "' + c.barcode + '" sudah dipakai barang "' + b.nama + '" (' + b.kode + ').');
    }
    const item = Object.assign({ id: Utilities.getUuid(), fisik: null, updatedAt: new Date().toISOString(), updatedBy: s.username }, c);
    appendRow_(barangSheet_(), itemRow_(item));
    log_(s.username, 'tambah barang', c.kode);
    return item;
  });
}
function api_updateItem(token, it) {
  const s = auth_(token, ['it']);
  const c = cleanItem_(it);
  return withLock_(() => {
    const items = readItems_();
    if (items.some(x => x.id !== it.id && x.kode.toLowerCase() === c.kode.toLowerCase())) throw new Error('SKU "' + c.kode + '" sudah dipakai barang lain.');
    if (c.barcode) {
      const b = items.find(x => x.id !== it.id && x.barcode && x.barcode.toLowerCase() === c.barcode.toLowerCase());
      if (b) throw new Error('Barcode "' + c.barcode + '" sudah dipakai barang "' + b.nama + '" (' + b.kode + ').');
    }
    const cur = items.find(x => x.id === it.id);
    if (!cur) throw new Error('Barang tidak ditemukan. Muat ulang data.');
    const item = Object.assign({}, cur, c, { updatedAt: new Date().toISOString(), updatedBy: s.username });
    const sh = barangSheet_();
    sh.getRange(findRow_(sh, it.id), 1, 1, H_BARANG.length).setValues([itemRow_(item)]);
    log_(s.username, 'ubah barang', c.kode);
    return item;
  });
}
function api_deleteItem(token, id) {
  const s = auth_(token, ['it']);
  return withLock_(() => {
    const sh = barangSheet_();
    const row = findRow_(sh, id);
    const kode = sh.getRange(row, 2).getValue();
    sh.deleteRow(row);
    log_(s.username, 'hapus barang', String(kode));
    return true;
  });
}
function api_importItems(token, rows) {
  const s = auth_(token, ['it']);
  if (!Array.isArray(rows) || !rows.length) throw new Error('Tidak ada baris untuk diimpor.');
  return withLock_(() => {
    const items = readItems_();
    const byKode = {};
    items.forEach(x => byKode[x.kode.toLowerCase()] = x);
    let add = 0, upd = 0, skipBarcode = 0;
    const now = new Date().toISOString();
    const byBarcode = {};
    items.forEach(x => { if (x.barcode) byBarcode[x.barcode.toLowerCase()] = x; });
    rows.forEach(r => {
      if (!String(r.kode || '').trim() || !String(r.nama || '').trim()) return;
      const c = cleanItem_(r);
      const ex = byKode[c.kode.toLowerCase()];
      const fis = (r.fisik === '' || r.fisik === null || r.fisik === undefined) ? null : Number(r.fisik);
      // barcode kosong di Excel = pertahankan barcode lama; barcode yang sudah dipakai barang lain dilewati
      if (c.barcode) {
        const owner = byBarcode[c.barcode.toLowerCase()];
        if (owner && owner !== ex) { c.barcode = ''; skipBarcode++; }
      }
      if (ex) {
        if (!c.barcode) delete c.barcode;
        Object.assign(ex, c, { updatedAt: now, updatedBy: s.username });
        if (c.barcode) byBarcode[c.barcode.toLowerCase()] = ex;
        if (fis !== null && isFinite(fis)) ex.fisik = fis;
        upd++;
      } else {
        const item = Object.assign({ id: Utilities.getUuid(), fisik: (fis !== null && isFinite(fis)) ? fis : null, updatedAt: now, updatedBy: s.username }, c);
        items.push(item); byKode[c.kode.toLowerCase()] = item;
        if (item.barcode) byBarcode[item.barcode.toLowerCase()] = item;
        add++;
      }
    });
    writeItems_(items);
    log_(s.username, 'import excel', add + ' baru, ' + upd + ' diperbarui, ' + skipBarcode + ' barcode dilewati');
    return { add: add, upd: upd, skipBarcode: skipBarcode };
  });
}

/* ---------- Laporan (untuk diunduh sebagai Excel dari Google Sheet) ---------- */
function api_buatLaporan(token) {
  const s = auth_(token, ['it']);
  const ss = ss_();
  let sh = ss.getSheetByName('Laporan');
  if (!sh) sh = ss.insertSheet('Laporan');
  sh.clear();
  const items = readItems_(), meta = getMeta_();
  const st = it => {
    if (it.fisik === null) return ['', 'Belum dihitung'];
    const d = it.fisik - it.sistem;
    return [d, d === 0 ? 'Sesuai' : (d > 0 ? 'Lebih' : 'Kurang')];
  };
  const done = items.filter(i => i.fisik !== null).length;
  const diff = items.filter(i => i.fisik !== null && i.fisik !== i.sistem).length;
  const top = [
    ['LAPORAN STOCK OPNAME', '', '', '', '', '', '', ''],
    ['Tanggal', meta.tanggal, '', '', '', '', '', ''],
    ['Petugas', meta.petugas, '', '', '', '', '', ''],
    ['Dibuat oleh', s.nama + ' (' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') + ')', '', '', '', '', '', ''],
    ['Total barang', items.length, '', '', '', '', '', ''],
    ['Sudah dihitung', done, '', '', '', '', '', ''],
    ['Belum dihitung', items.length - done, '', '', '', '', '', ''],
    ['Barang ada selisih', diff, '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '']
  ];
  const head = ['SKU', 'Barcode', 'Nama Barang', 'Satuan', 'Lokasi', 'Stok Sistem', 'Stok Fisik', 'Selisih', 'Status'];
  const body = items.map(i => { const x = st(i); return [i.kode, i.barcode, i.nama, i.satuan, i.lokasi, i.sistem, i.fisik === null ? '' : i.fisik, x[0], x[1]]; });
  sh.getRange(1, 1, top.length, 8).setValues(top);
  sh.getRange(top.length + 1, 1, 1, 9).setValues([head]).setFontWeight('bold').setBackground('#e8eef7');
  if (body.length) {
    sh.getRange(top.length + 2, 1, body.length, 2).setNumberFormat('@'); // SKU & barcode tetap teks
    sh.getRange(top.length + 2, 1, body.length, 9).setValues(body);
  }
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  sh.autoResizeColumns(1, 9);
  log_(s.username, 'buat laporan', items.length + ' barang');
  return { url: ss.getUrl() + '#gid=' + sh.getSheetId() };
}

/* ---------- Kelola akun (IT) ---------- */
function api_listUsers(token) {
  auth_(token, ['it']);
  return readUsers_().map(u => ({ username: u.username, nama: u.nama, role: u.role }));
}
function api_saveUser(token, u) {
  const s = auth_(token, ['it']);
  const uname = String(u.username || '').trim().toLowerCase();
  const nama = String(u.nama || '').trim();
  const role = u.role === 'it' ? 'it' : 'gudang';
  const pw = String(u.password || '');
  if (!/^[a-z0-9._-]{3,}$/.test(uname)) throw new Error('Username minimal 3 karakter (huruf kecil, angka, titik, strip, garis bawah).');
  if (!nama) throw new Error('Nama wajib diisi.');
  return withLock_(() => {
    const sh = sheet_('Akun', H_AKUN, [1, 2, 3, 4, 5]);
    const users = readUsers_();
    const ex = users.find(x => x.username === uname);
    if (!ex) {
      if (pw.length < 8) throw new Error('Password minimal 8 karakter.');
      appendRow_(sh, userRow_(uname, nama, role, pw));
      log_(s.username, 'tambah akun', uname + ' (' + role + ')');
      return true;
    }
    if (pw && pw.length < 8) throw new Error('Password minimal 8 karakter.');
    if (ex.role === 'it' && role !== 'it' && users.filter(x => x.role === 'it').length <= 1) throw new Error('Harus ada minimal satu akun IT.');
    const vals = sh.getDataRange().getValues();
    const idx = vals.findIndex((r, i) => i > 0 && String(r[0]) === uname);
    const row = pw ? userRow_(uname, nama, role, pw) : [uname, nama, role, ex.salt, ex.hash];
    sh.getRange(idx + 1, 1, 1, H_AKUN.length).setValues([row]);
    revoke_(uname);
    log_(s.username, 'ubah akun', uname + (pw ? ' (password diganti)' : ''));
    return true;
  });
}
function api_deleteUser(token, username) {
  const s = auth_(token, ['it']);
  const uname = String(username || '').trim().toLowerCase();
  if (uname === s.username) throw new Error('Tidak bisa menghapus akun yang sedang dipakai.');
  return withLock_(() => {
    const sh = sheet_('Akun', H_AKUN, [1, 2, 3, 4, 5]);
    const users = readUsers_();
    const ex = users.find(x => x.username === uname);
    if (!ex) throw new Error('Akun tidak ditemukan.');
    if (ex.role === 'it' && users.filter(x => x.role === 'it').length <= 1) throw new Error('Harus ada minimal satu akun IT.');
    const vals = sh.getDataRange().getValues();
    const idx = vals.findIndex((r, i) => i > 0 && String(r[0]) === uname);
    sh.deleteRow(idx + 1);
    revoke_(uname);
    log_(s.username, 'hapus akun', uname);
    return true;
  });
}
