/**
 * INBIONUSA - Quality Control System. Versi 4.3
 * Jalankan setupDatabase() sekali dari editor (juga untuk migrasi data lama).
 * Semua panggilan klien lewat api(token, fn, args). Fungsi berakhiran _ tidak bisa dipanggil dari browser.
 *
 * Perubahan v4.3:
 *  - Auto-create sheet jika belum ada (fix: "Cannot read properties of null (reading 'getLastRow')")
 *  - Cache layer untuk semua read (rows_) + auto bust setiap write
 *  - Endpoint dashboardFull() — 1 round-trip untuk dashboard
 *  - Endpoint deleteUji(), deleteCOA() untuk admin
 *  - Role-based permission matrix (ROLE_ALLOW)
 *  - Normalisasi spesifikasi min>max + cleanup duplikat HasilUji saat setup
 */

const SH = {
  Users: ['Username', 'Password', 'Role', 'Nama Lengkap'],
  MasterProduk: ['Kode Produk', 'Nama Produk', 'Kategori', 'Satuan', 'Status Halal', 'Tipe Kemasan'],
  Spesifikasi: ['ID Spek', 'Kode Produk', 'Parameter Uji', 'Nilai Min', 'Nilai Maks', 'Satuan', 'Keterangan', 'Tipe Batas'],
  HasilUji: ['ID Uji', 'No Lot', 'Kode Produk', 'Tgl Sampling', 'Tgl Uji', 'Inspector', 'Hasil Detail JSON', 'Status', 'Catatan'],
  COA: ['No COA', 'Kode Produk', 'No Lot', 'Tanggal Rilis', 'Tanggal Expired', 'Hasil Uji JSON', 'Tanda Tangan Digital', 'Mfg Date'],
  DokumenArsip: ['ID Dokumen', 'Judul Dokumen', 'Kategori', 'Tanggal Berlaku', 'Masa Kadaluarsa', 'Link Drive'],
  AuditLog: ['Waktu', 'User', 'Aksi', 'Referensi']
};

const SALT = 'INBIO::';

const WRITE = {
  saveProduk: 1, deleteProduk: 1,
  saveUji: 1, updateUji: 1, deleteUji: 1,
  issueCOA: 1, updateCOA: 1, deleteCOA: 1,
  uploadDoc: 1, deleteDoc: 1
};

// Role yang boleh melakukan aksi. Tidak ada di list = semua role login boleh.
const ROLE_ALLOW = {
  deleteProduk: ['admin qc', 'administrator inbionusa'],
  deleteUji: ['admin qc', 'administrator inbionusa'],
  deleteCOA: ['admin qc', 'administrator inbionusa'],
  deleteDoc: ['admin qc', 'administrator inbionusa'],
  updateCOA: ['admin qc', 'administrator inbionusa', 'qc supervisor'],
  issueCOA: ['admin qc', 'administrator inbionusa', 'qc supervisor']
};

const CACHE_TTL = 180; // detik

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('INBIONUSA - Quality Control System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// =========================================================
// CACHE LAYER
// =========================================================
function ver_() {
  const P = PropertiesService.getScriptProperties();
  return P.getProperty('DATA_VER') || '1';
}
function bust_() {
  PropertiesService.getScriptProperties().setProperty('DATA_VER', String(Date.now()));
}
function cached_(key, fn, ttl) {
  const c = CacheService.getScriptCache();
  const k = 'c:' + key + ':' + ver_();
  const hit = c.get(k);
  if (hit) { try { return JSON.parse(hit); } catch (e) { } }
  const val = fn();
  try {
    const s = JSON.stringify(val);
    if (s.length < 95000) c.put(k, s, ttl || CACHE_TTL);
  } catch (e) { }
  return val;
}

// =========================================================
// HELPER DASAR
// =========================================================
const ss_ = () => SpreadsheetApp.getActiveSpreadsheet();
const tz_ = () => ss_().getSpreadsheetTimeZone();
const str_ = v => v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : (v == null ? '' : String(v));
const same_ = (a, b) => str_(a).trim() === str_(b).trim();

/** Ambil sheet; jika belum ada, buat otomatis lengkap dengan header. */
function sh_(n) {
  if (!SH[n]) throw new Error('Nama sheet tidak dikenal: ' + n);
  const ss = ss_();
  let sh = ss.getSheetByName(n);
  if (!sh) {
    sh = ss.insertSheet(n);
    initSheet_(sh, n);
  }
  return sh;
}

function initSheet_(sh, n) {
  const h = SH[n], c = h.length;
  const lr = Math.max(sh.getLastRow(), 1);
  const r = sh.getRange(1, 1, lr, c), v = r.getValues().map(x => x.map(str_));
  v[0] = h.slice();
  r.setNumberFormat('@').setValues(v);
  if (sh.getMaxRows() > lr) sh.getRange(lr + 1, 1, sh.getMaxRows() - lr, c).setNumberFormat('@');
  sh.getRange(1, 1, 1, c).setFontWeight('bold').setBackground('#047857').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function hash_(p) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SALT + p)
    .map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');
}

function locked_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(20000);
  try {
    const r = fn();
    bust_(); // invalidate cache setelah write
    return r;
  } finally {
    l.releaseLock();
  }
}

function rows_(n) {
  return cached_('rows:' + n, () => {
    const sh = sh_(n);
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, 1, sh.getLastRow() - 1, SH[n].length).getValues()
      .map(r => r.map(str_)).filter(r => r[0].trim() !== '');
  });
}

function find_(n, col, val) {
  const sh = sh_(n), lr = sh.getLastRow();
  if (lr < 2) return -1;
  const v = sh.getRange(2, col, lr - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (same_(v[i][0], val)) return i + 2;
  return -1;
}

function add_(n, arr) {
  if (!arr || !arr.length) return;
  const sh = sh_(n), lr = sh.getLastRow(), c = SH[n].length;
  if (sh.getMaxRows() < lr + arr.length) {
    sh.insertRowsAfter(sh.getMaxRows(), lr + arr.length - sh.getMaxRows() + 5);
  }
  sh.getRange(lr + 1, 1, arr.length, c).setNumberFormat('@').setValues(arr);
}

function audit_(u, a, ref) {
  try {
    add_('AuditLog', [[Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'), u && u.u || '-', a, ref]]);
  } catch (e) { /* jangan gagalkan aksi utama */ }
}

function need_(v, label) {
  v = v == null ? '' : String(v).trim();
  if (!v) throw new Error(label + ' wajib diisi');
  return v;
}

function num_(v) {
  const s = String(v == null ? '' : v).trim().replace(',', '.');
  return s === '' || isNaN(+s) ? null : +s;
}

function date_(v, label) {
  const s = String(v || '').trim().slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s) || isNaN(new Date(s))) throw new Error(label + ' tidak valid');
  return s;
}

const nameMap_ = () => {
  const m = {};
  rows_('MasterProduk').forEach(r => m[r[0].trim()] = r[1]);
  return m;
};

// =========================================================
// SETUP & MIGRASI
// =========================================================
function setupDatabase() {
  const ss = ss_();
  Object.keys(SH).forEach(n => {
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    initSheet_(sh, n);
  });

  // Users: hash password lama jika perlu
  const us = sh_('Users');
  if (us.getLastRow() < 2) {
    add_('Users', [['INBIONUSA', hash_('INBIONUSA'), 'Admin QC', 'Administrator INBIONUSA']]);
  } else {
    const n = us.getLastRow() - 1;
    const pw = us.getRange(2, 2, n, 1).getValues()
      .map(x => [String(x[0]).length === 64 ? x[0] : hash_(x[0])]);
    us.getRange(2, 2, n, 1).setValues(pw);
  }

  // Hapus sheet default jika kosong
  ['Sheet1', 'Lembar1', 'Muka1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  // Migrasi data lama
  fixLegacyData_();

  bust_(); // refresh semua cache
  return 'Database siap. Password lama sudah di-hash. Data lama sudah dimigrasi. Segera ganti password lewat menu Ganti Password.';
}

/** Perbaiki data lama: spesifikasi min>max dibalik, duplikat HasilUji dibuang. */
function fixLegacyData_() {
  // 1) Perbaiki Spesifikasi yang min > max (RANGE)
  const sp = sh_('Spesifikasi'), lr = sp.getLastRow();
  let fixedSpec = 0;
  if (lr >= 2) {
    const v = sp.getRange(2, 1, lr - 1, 8).getValues();
    for (let i = 0; i < v.length; i++) {
      const t = String(v[i][7] || 'RANGE');
      if (t !== 'RANGE') continue;
      const a = num_(v[i][3]), b = num_(v[i][4]);
      if (a !== null && b !== null && a > b) {
        v[i][3] = String(b);
        v[i][4] = String(a);
        fixedSpec++;
      }
    }
    if (fixedSpec) sp.getRange(2, 1, v.length, 8).setNumberFormat('@').setValues(v);
  }

  // 2) Buang duplikat HasilUji: simpan ID Uji paling besar per kode+lot
  const hu = sh_('HasilUji'), lh = hu.getLastRow();
  let removed = 0;
  if (lh >= 2) {
    const v = hu.getRange(2, 1, lh - 1, 9).getValues();
    const best = {};
    v.forEach((r, i) => {
      const k = String(r[2]).trim() + '|' + String(r[1]).trim();
      const id = String(r[0]);
      if (!best[k] || id > String(best[k].id)) best[k] = { id: id, idx: i };
    });
    const keep = new Set(Object.keys(best).map(k => best[k].idx));
    // hapus dari bawah agar index tetap valid
    for (let i = v.length - 1; i >= 0; i--) {
      if (!keep.has(i)) {
        hu.deleteRow(i + 2);
        removed++;
      }
    }
  }

  if (fixedSpec || removed) {
    try {
      add_('AuditLog', [[
        Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'),
        'SYSTEM',
        'MIGRASI',
        'Spec fixed: ' + fixedSpec + ', duplicate removed: ' + removed
      ]]);
    } catch (e) { }
  }
}

// =========================================================
// DISPATCHER + SESI
// =========================================================
function api(token, fn, args) {
  try {
    args = args || [];

    // Pastikan semua sheet ada sebelum apapun
    Object.keys(SH).forEach(n => sh_(n));

    if (fn === 'login') return { ok: true, data: login_(args[0], args[1]) };

    const c = CacheService.getScriptCache();
    const raw = token ? c.get('s_' + token) : null;
    if (!raw) return { ok: false, auth: false, msg: 'Sesi berakhir, silakan login ulang.' };

    // Refresh TTL sesi
    c.put('s_' + token, raw, 21600);

    if (fn === 'logout') { c.remove('s_' + token); return { ok: true }; }

    const u = JSON.parse(raw);
    const f = API[fn];
    if (!f) throw new Error('Fungsi tidak dikenal');

    const role = String(u.role || '').toLowerCase();
    if (WRITE[fn]) {
      if (role === 'viewer') throw new Error('Akun ini hanya memiliki akses lihat.');
      const allow = ROLE_ALLOW[fn];
      if (allow && allow.indexOf(role) < 0) {
        throw new Error('Role "' + u.role + '" tidak berhak melakukan aksi ini.');
      }
    }

    const r = f.apply(null, [u].concat(args));
    return { ok: true, data: r && r.data, msg: r && r.msg };
  } catch (e) {
    return { ok: false, msg: e.message || String(e) };
  }
}

function login_(user, pass) {
  user = String(user || '').trim();
  if (!user) throw new Error('Username wajib diisi');

  const c = CacheService.getScriptCache();
  const k = 'f_' + user.toLowerCase();
  if (+(c.get(k) || 0) >= 5) throw new Error('Terlalu banyak percobaan gagal. Coba lagi dalam 5 menit.');

  const sh = sh_('Users'), lr = sh.getLastRow();
  const v = lr < 2 ? [] : sh.getRange(2, 1, lr - 1, 4).getValues();

  for (let i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() !== user) continue;
    const st = String(v[i][1]);
    const h = hash_(pass);
    if (st.length === 64 ? st === h : st === String(pass)) {
      // Migrasi password plaintext → hash
      if (st.length !== 64) sh.getRange(i + 2, 2).setNumberFormat('@').setValue(h);
      c.remove(k);

      const t = Utilities.getUuid();
      const d = { u: user, role: String(v[i][2]), nama: String(v[i][3]) };
      c.put('s_' + t, JSON.stringify(d), 21600);
      return { token: t, nama: d.nama, role: d.role };
    }
  }
  c.put(k, String(+(c.get(k) || 0) + 1), 300);
  throw new Error('Username atau password salah.');
}

// =========================================================
// NORMALISASI DATA HASIL UJI (format lama & baru)
// =========================================================
function norm_(j, st) {
  let o;
  try { o = JSON.parse(j); } catch (e) { return []; }
  if (o && o.items) o = o.items;
  if (Array.isArray(o)) return o;
  return Object.keys(o || {}).map(k => {
    const i = o[k] && typeof o[k] === 'object' ? o[k] : { val: o[k] };
    const mn = i.min && i.min !== '-' ? String(i.min) : '';
    const mx = i.max && i.max !== '-' ? String(i.max) : '';
    return {
      p: k,
      t: mn && mx ? 'RANGE' : mn ? 'MIN' : mx ? 'MAX' : 'TEXT',
      mn: mn, mx: mx, u: '',
      k: i.specText || '',
      val: String(i.val == null ? '' : i.val),
      st: st || 'PASS'
    };
  });
}

function snap_(j) {
  let o = {};
  try { o = JSON.parse(j) || {}; } catch (e) { }
  return {
    produk: o.produk || '',
    ts: o.ts || '',
    tu: o.tu || '',
    items: norm_(j, 'PASS')
  };
}

function specs_(kode) {
  return rows_('Spesifikasi').filter(r => same_(r[1], kode))
    .map(r => {
      let mn = r[3], mx = r[4];
      const t = r[7] || 'RANGE';
      if (t === 'RANGE') {
        const a = num_(mn), b = num_(mx);
        if (a !== null && b !== null && a > b) { const tmp = mn; mn = mx; mx = tmp; }
      }
      return { p: r[2], mn: mn, mx: mx, u: r[5], k: r[6], t: t };
    });
}

function cleanSpecs_(list) {
  if (!list || !list.length) throw new Error('Minimal satu parameter spesifikasi');
  const seen = {};
  return list.map(s => {
    const p = need_(s.parameter, 'Nama parameter');
    const key = p.toLowerCase();
    if (seen[key]) throw new Error('Nama parameter ganda: ' + p);
    seen[key] = 1;

    const t = ['RANGE', 'MIN', 'MAX', 'TEXT'].indexOf(s.tipeBatas) < 0 ? 'RANGE' : s.tipeBatas;
    const useMn = t === 'RANGE' || t === 'MIN';
    const useMx = t === 'RANGE' || t === 'MAX';
    const mn = useMn ? num_(s.nilaiMin) : null;
    const mx = useMx ? num_(s.nilaiMaks) : null;

    if (useMn && mn === null) throw new Error(p + ': nilai min harus berupa angka');
    if (useMx && mx === null) throw new Error(p + ': nilai maks harus berupa angka');
    if (t === 'RANGE' && mn > mx) throw new Error(p + ': nilai min tidak boleh lebih besar dari maks');

    return {
      p: p,
      t: t,
      mn: useMn ? String(s.nilaiMin).trim() : '',
      mx: useMx ? String(s.nilaiMaks).trim() : '',
      u: String(s.satuan || '').trim(),
      k: String(s.keterangan || '').trim()
    };
  });
}

function delSpecs_(kode) {
  const sh = sh_('Spesifikasi'), lr = sh.getLastRow();
  if (lr < 2) return;
  const v = sh.getRange(2, 2, lr - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    if (same_(v[i][0], kode)) sh.deleteRow(i + 2);
  }
}

function evalItem_(s, val, manual) {
  if (!val) throw new Error(s.p + ': hasil uji wajib diisi');
  if (s.t === 'TEXT') return manual === 'FAIL' ? 'FAIL' : 'PASS';
  const x = num_(val), mn = num_(s.mn), mx = num_(s.mx);
  if (x === null) throw new Error(s.p + ': hasil uji harus berupa angka');
  const low = (s.t === 'RANGE' || s.t === 'MIN') && mn !== null && x < mn;
  const high = (s.t === 'RANGE' || s.t === 'MAX') && mx !== null && x > mx;
  return low || high ? 'FAIL' : 'PASS';
}

function buildItems_(sp, hasil) {
  hasil = hasil || {};
  return sp.map(s => {
    const i = hasil[s.p] || {};
    const v = String(i.val == null ? '' : i.val).trim();
    return {
      p: s.p, t: s.t, mn: s.mn, mx: s.mx, u: s.u, k: s.k,
      val: v, st: evalItem_(s, v, i.st)
    };
  });
}

// =========================================================
// API ENDPOINTS
// =========================================================
const API = {

  // --- Dashboard gabungan (1 round-trip, cached) ---
  dashboardFull() {
    return { data: cached_('dashFull', () => {
      const all = rows_('HasilUji');
      const last = {};
      all.forEach(r => { last[r[2].trim() + '|' + r[1].trim()] = r[7]; });
      const st = Object.keys(last).map(k => last[k]);
      const dk = rows_('DokumenArsip');
      const today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
      const nm = nameMap_();
      const recent = all.slice(-8).reverse().map(r => ({
        id: r[0], lot: r[1], kode: r[2],
        nama: nm[r[2].trim()] || '',
        ts: r[3], tu: r[4], st: r[7]
      }));
      return {
        produk: rows_('MasterProduk').length,
        lot: st.length,
        pass: st.filter(x => x === 'PASS').length,
        fail: st.filter(x => x === 'FAIL').length,
        coa: rows_('COA').length,
        dok: dk.filter(d => d[4] >= today).length,
        dokExp: dk.filter(d => d[4] && d[4] < today).length,
        recent: recent
      };
    }, 60) };
  },

  // --- Dashboard lama (fallback) ---
  dashboard() {
    const last = {};
    rows_('HasilUji').forEach(r => { last[r[2].trim() + '|' + r[1].trim()] = r[7]; });
    const st = Object.keys(last).map(k => last[k]);
    const dk = rows_('DokumenArsip');
    const today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
    return { data: {
      produk: rows_('MasterProduk').length,
      lot: st.length,
      pass: st.filter(x => x === 'PASS').length,
      fail: st.filter(x => x === 'FAIL').length,
      coa: rows_('COA').length,
      dok: dk.filter(d => d[4] >= today).length,
      dokExp: dk.filter(d => d[4] && d[4] < today).length
    } };
  },

  // --- Master Produk ---
  listProduk() {
    return { data: rows_('MasterProduk').map(r => ({
      kode: r[0], nama: r[1], kat: r[2], sat: r[3], halal: r[4], kem: r[5]
    })) };
  },

  specs(u, kode) {
    return { data: specs_(kode) };
  },

  saveProduk(u, p) {
    return locked_(() => {
      const kode = need_(p.kodeProduk, 'Kode produk');
      const edit = !!p.edit;
      const row = find_('MasterProduk', 1, kode);
      if (edit && row < 0) throw new Error('Produk tidak ditemukan');
      if (!edit && row > 0) throw new Error('Kode produk sudah terdaftar');

      const rec = [
        kode,
        need_(p.namaProduk, 'Nama produk'),
        need_(p.kategori, 'Kategori'),
        need_(p.satuan, 'Satuan'),
        p.statusHalal === 'Non-Halal' ? 'Non-Halal' : 'Halal',
        need_(p.tipeKemasan, 'Kemasan')
      ];

      // Validasi spesifikasi SEBELUM tulis apapun
      const sp = cleanSpecs_(p.spekList);

      if (edit) {
        sh_('MasterProduk').getRange(row, 1, 1, 6).setNumberFormat('@').setValues([rec]);
        delSpecs_(kode);
      } else {
        add_('MasterProduk', [rec]);
      }
      add_('Spesifikasi', sp.map((s, i) => [
        'SPK-' + kode + '-' + (i + 1), kode, s.p, s.mn, s.mx, s.u, s.k, s.t
      ]));

      audit_(u, edit ? 'EDIT_PRODUK' : 'ADD_PRODUK', kode);
      return { msg: 'Produk ' + rec[1] + ' berhasil disimpan.' };
    });
  },

  deleteProduk(u, kode) {
    return locked_(() => {
      const row = find_('MasterProduk', 1, kode);
      if (row < 0) throw new Error('Produk tidak ditemukan');
      if (rows_('HasilUji').some(r => same_(r[2], kode))) {
        throw new Error('Produk sudah memiliki hasil uji, tidak dapat dihapus.');
      }
      delSpecs_(kode);
      sh_('MasterProduk').deleteRow(row);
      audit_(u, 'DEL_PRODUK', kode);
      return { msg: 'Produk ' + kode + ' dan spesifikasinya dihapus.' };
    });
  },

  // --- Inspeksi ---
  listUji() {
    const nm = nameMap_();
    return { data: rows_('HasilUji').reverse().map(r => ({
      id: r[0], lot: r[1], kode: r[2],
      nama: nm[r[2].trim()] || '',
      ts: r[3], tu: r[4], insp: r[5],
      st: r[7], cat: r[8],
      items: norm_(r[6], r[7])
    })) };
  },

  saveUji(u, p) {
    return locked_(() => {
      const kode = need_(p.kodeProduk, 'Produk');
      const lot = need_(p.noLot, 'No lot');
      const ts = date_(p.tglSampling, 'Tgl sampling');
      const tu = date_(p.tglUji, 'Tgl uji');
      if (tu < ts) throw new Error('Tgl uji tidak boleh sebelum tgl sampling');
      if (find_('MasterProduk', 1, kode) < 0) throw new Error('Produk tidak ditemukan');
      if (rows_('HasilUji').some(r => same_(r[2], kode) && same_(r[1], lot))) {
        throw new Error('Lot ' + lot + ' untuk produk ini sudah ada. Gunakan tombol Edit.');
      }

      const sp = specs_(kode);
      if (!sp.length) throw new Error('Produk belum memiliki spesifikasi');
      const items = buildItems_(sp, p.hasil);
      const st = items.every(i => i.st === 'PASS') ? 'PASS' : 'FAIL';

      add_('HasilUji', [[
        'INSP-' + Date.now(), lot, kode, ts, tu,
        need_(p.inspector, 'Inspector'),
        JSON.stringify(items), st,
        String(p.catatan || '').trim() || '-'
      ]]);

      audit_(u, 'ADD_UJI', kode + '/' + lot + '/' + st);
      return { msg: 'Hasil uji lot ' + lot + ' tersimpan. Status: ' + st };
    });
  },

  updateUji(u, p) {
    return locked_(() => {
      const row = find_('HasilUji', 1, need_(p.idUji, 'ID uji'));
      if (row < 0) throw new Error('Data tidak ditemukan');
      const o = sh_('HasilUji').getRange(row, 1, 1, 9).getValues()[0].map(str_);

      if (rows_('COA').some(r => same_(r[1], o[2]) && same_(r[2], o[1]))) {
        throw new Error('Lot sudah memiliki COA, hasil uji tidak dapat diubah.');
      }

      const ts = date_(p.tglSampling, 'Tgl sampling');
      const tu = date_(p.tglUji, 'Tgl uji');
      if (tu < ts) throw new Error('Tgl uji tidak boleh sebelum tgl sampling');

      const sp = norm_(o[6], o[7]).map(i => ({
        p: i.p, t: i.t, mn: i.mn, mx: i.mx, u: i.u, k: i.k
      }));
      const items = buildItems_(sp, p.hasil);
      const st = items.every(i => i.st === 'PASS') ? 'PASS' : 'FAIL';

      sh_('HasilUji').getRange(row, 4, 1, 6).setNumberFormat('@')
        .setValues([[ts, tu, need_(p.inspector, 'Inspector'), JSON.stringify(items), st,
          String(p.catatan || '').trim() || '-']]);

      audit_(u, 'EDIT_UJI', o[2] + '/' + o[1] + '/' + o[7] + '>' + st);
      return { msg: 'Hasil uji lot ' + o[1] + ' diperbarui. Status: ' + st };
    });
  },

  deleteUji(u, id) {
    return locked_(() => {
      const row = find_('HasilUji', 1, need_(id, 'ID uji'));
      if (row < 0) throw new Error('Data tidak ditemukan');
      const r = sh_('HasilUji').getRange(row, 1, 1, 9).getValues()[0].map(str_);
      if (rows_('COA').some(x => same_(x[1], r[2]) && same_(x[2], r[1]))) {
        throw new Error('Lot sudah memiliki COA, hasil uji tidak dapat dihapus.');
      }
      sh_('HasilUji').deleteRow(row);
      audit_(u, 'DEL_UJI', r[2] + '/' + r[1]);
      return { msg: 'Hasil uji lot ' + r[1] + ' dihapus.' };
    });
  },

  // --- COA ---
  listCOA() {
    return { data: rows_('COA').reverse().map(r => {
      const s = snap_(r[5]);
      return {
        no: r[0], kode: r[1], lot: r[2],
        rilis: r[3], exp: r[4], ttd: r[6],
        mfg: r[7] || '',
        produk: s.produk, ts: s.ts, tu: s.tu, items: s.items
      };
    }) };
  },

  issueCOA(u, p) {
    return locked_(() => {
      const row = find_('HasilUji', 1, need_(p.idUji, 'Lot'));
      if (row < 0) throw new Error('Data hasil uji tidak ditemukan');
      const r = sh_('HasilUji').getRange(row, 1, 1, 9).getValues()[0].map(str_);

      if (r[7] !== 'PASS') throw new Error('COA hanya dapat diterbitkan untuk lot berstatus PASS');
      if (rows_('COA').some(x => same_(x[1], r[2]) && same_(x[2], r[1]))) {
        throw new Error('Lot ini sudah memiliki COA');
      }

      const rl = date_(p.tglRilis, 'Tgl rilis');
      const ex = date_(p.tglExpired, 'Tgl expired');
      if (ex <= rl) throw new Error('Tgl expired harus setelah tgl rilis');

      const pre = 'COA/IBN/' + Utilities.formatDate(new Date(), tz_(), 'yyyy') + '/';
      let n = 0;
      rows_('COA').forEach(x => {
        if (x[0].indexOf(pre) === 0) {
          n = Math.max(n, parseInt(x[0].slice(pre.length), 10) || 0);
        }
      });
      const no = pre + String(n + 1).padStart(4, '0');

      const mfg = p.tglMfg ? date_(p.tglMfg, 'Mfg date') : '';
      add_('COA', [[
        no, r[2], r[1], rl, ex,
        JSON.stringify({
          produk: nameMap_()[r[2].trim()] || '',
          ts: r[3], tu: r[4],
          items: norm_(r[6], r[7])
        }),
        String(p.ttd || '').trim() || 'Terverifikasi Digital',
        mfg
      ]]);

      audit_(u, 'ISSUE_COA', no);
      return { msg: 'COA ' + no + ' berhasil diterbitkan.' };
    });
  },

  updateCOA(u, p) {
    return locked_(() => {
      const row = find_('COA', 1, need_(p.noCOA, 'No COA'));
      if (row < 0) throw new Error('COA tidak ditemukan');
      const rl = date_(p.tglRilis, 'Tgl rilis');
      const ex = date_(p.tglExpired, 'Tgl expired');
      if (ex <= rl) throw new Error('Tgl expired harus setelah tgl rilis');

      const sh = sh_('COA');
      sh.getRange(row, 4, 1, 2).setNumberFormat('@').setValues([[rl, ex]]);
      sh.getRange(row, 7).setNumberFormat('@')
        .setValue(String(p.ttd || '').trim() || 'Terverifikasi Digital');
      sh.getRange(row, 8).setNumberFormat('@')
        .setValue(p.tglMfg ? date_(p.tglMfg, 'Mfg date') : '');

      audit_(u, 'EDIT_COA', p.noCOA);
      return { msg: 'COA ' + p.noCOA + ' diperbarui.' };
    });
  },

  deleteCOA(u, no) {
    return locked_(() => {
      const row = find_('COA', 1, need_(no, 'No COA'));
      if (row < 0) throw new Error('COA tidak ditemukan');
      sh_('COA').deleteRow(row);
      audit_(u, 'DEL_COA', no);
      return { msg: 'COA ' + no + ' dihapus.' };
    });
  },

  // --- Dokumen ---
  listDocs() {
    return { data: rows_('DokumenArsip').reverse().map(r => ({
      id: r[0], judul: r[1], kat: r[2],
      b: r[3], k: r[4], link: r[5]
    })) };
  },

  uploadDoc(u, f, p) {
    const jd = need_(p.judul, 'Judul');
    const kat = need_(p.kategori, 'Kategori');
    const b = date_(p.tglBerlaku, 'Tgl berlaku');
    const k = date_(p.tglKadaluarsa, 'Masa kadaluarsa');
    if (k < b) throw new Error('Masa kadaluarsa tidak boleh sebelum tgl berlaku');

    return locked_(() => {
      let link;
      if (f && f.data) {
        if (f.data.length > 11000000) throw new Error('Ukuran file maksimal 8 MB');
        if (f.mimeType !== 'application/pdf') throw new Error('Hanya file PDF yang diterima');
        link = folder_()
          .createFile(Utilities.newBlob(
            Utilities.base64Decode(f.data),
            f.mimeType,
            String(f.fileName || 'dokumen.pdf')
          ))
          .getUrl();
      } else {
        link = String(p.linkDrive || '').trim();
        if (!/^https:\/\/(drive|docs)\.google\.com\//.test(link)) {
          throw new Error('Pilih file PDF atau isi link Google Drive yang valid');
        }
      }

      add_('DokumenArsip', [['DOC-' + Date.now(), jd, kat, b, k, link]]);
      audit_(u, 'ADD_DOC', jd);
      return { msg: 'Dokumen diarsipkan. File tersimpan privat di folder Drive "INBIONUSA Dokumen Mutu"; atur akses lewat folder tersebut.' };
    });
  },

  deleteDoc(u, id) {
    return locked_(() => {
      const row = find_('DokumenArsip', 1, need_(id, 'ID Dokumen'));
      if (row < 0) throw new Error('Dokumen tidak ditemukan');
      sh_('DokumenArsip').deleteRow(row);
      audit_(u, 'DEL_DOC', id);
      return { msg: 'Data arsip dihapus (file di Drive tidak ikut terhapus).' };
    });
  },

  // --- User ---
  gantiPassword(u, lama, baru) {
    return locked_(() => {
      if (String(baru || '').length < 8) throw new Error('Password baru minimal 8 karakter');
      const row = find_('Users', 1, u.u);
      if (row < 0) throw new Error('User tidak ditemukan');
      const sh = sh_('Users');
      const st = String(sh.getRange(row, 2).getValue());
      if (st !== hash_(lama)) throw new Error('Password lama salah');
      sh.getRange(row, 2).setNumberFormat('@').setValue(hash_(baru));
      audit_(u, 'GANTI_PASSWORD', u.u);
      return { msg: 'Password berhasil diganti.' };
    });
  },

  // --- Debug / utilitas (opsional) ---
  _cacheInfo() {
    return { data: { ver: ver_(), ttl: CACHE_TTL } };
  }
};

// =========================================================
// FOLDER DRIVE (untuk upload dokumen)
// =========================================================
function folder_() {
  const P = PropertiesService.getScriptProperties();
  const id = P.getProperty('DOC_FOLDER');
  try { if (id) return DriveApp.getFolderById(id); } catch (e) { }
  const f = DriveApp.createFolder('INBIONUSA Dokumen Mutu');
  P.setProperty('DOC_FOLDER', f.getId());
  return f;
}

// =========================================================
// CLEANUP CACHE MANUAL (jalankan dari editor jika perlu)
// =========================================================
function clearAllCache() {
  CacheService.getScriptCache().removeAll(Object.keys(SH).map(n => 'rows:' + n).concat(['dashFull']));
  bust_();
  return 'Cache dibersihkan.';
}
