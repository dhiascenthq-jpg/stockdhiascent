// =====================================================================
// Dhiascent Stok — API untuk app telefon (GitHub Pages)
// Projek Apps Script BERASINGAN. Ia baca & tulis sheet "Dhiascent Stok Operasi".
//
// PIN admin (untuk Dashboard): tab Tetapan, baris "PIN admin".
//
// LANGKAH PASANG (sekali sahaja):
//  1) Buka script.google.com > New project. Namakan "Dhiascent Stok API".
//  2) Padam kod contoh, tampal SEMUA kod fail ini. Save.
//  3) Pilih fungsi `test` di bar atas > Run > benarkan akses (Allow).
//  4) Deploy > New deployment > jenis: Web app
//        Execute as: Me        Who has access: Anyone
//  5) Salin URL /exec dan beri pada Claude (atau tampal dalam config.js).
// PIN app diambil dari tab Tetapan (sel B2). Tukar PIN di situ sahaja.
// =====================================================================

const SHEET_ID = '1wtbnS5w8Nzo4cg0OCaUGgsRoHureQqtlzQ64kuqAvRI';
const TZ = 'Asia/Kuala_Lumpur';

// Langkah yang dibenarkan, dan sheet mana item itu tinggal
const STEPS = {
  'Stock out':    'produk',   // ambil untuk packing / jual
  'Stock in':     'produk',   // refill siap masuk stor
  'Masuk balik':  'produk',   // RETURN pelanggan sahaja (bukan baki)
  'Guna bahan':   'bahan',    // rosak / packing / manual
  'Terima bahan': 'bahan',    // barang sampai dari supplier
  'Kira stok':    'semua'     // kiraan fizikal: reset baki
};

function doGet(e)  { return handle_((e && e.parameter) || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  return handle_(p);
}

function handle_(p) {
  try {
    checkPin_(p.pin);
    switch (p.action) {
      case 'ping':   return json_({ ok: true });
      case 'data':   return json_({ ok: true, data: getData_() });
      case 'record': return json_({ ok: true, result: record_(p) });
      case 'admin':  checkAdmin_(p.adminPin); return json_({ ok: true });
      case 'dash':   checkAdmin_(p.adminPin); return json_({ ok: true, dash: getDash_(p) });
      default:       throw new Error('Action tidak dikenali');
    }
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function checkPin_(pin) {
  const set = ss_().getSheetByName('Tetapan').getRange('B2').getDisplayValue().trim();
  if (set && String(pin || '').trim() !== set) throw new Error('PIN salah');
}

// PIN admin: tab Tetapan, baris berlabel "PIN admin" (lajur B)
function checkAdmin_(pin) {
  const vals = ss_().getSheetByName('Tetapan').getRange('A1:B20').getDisplayValues();
  const row = vals.find(r => String(r[0]).trim().toLowerCase() === 'pin admin');
  const set = row ? String(row[1]).trim() : '';
  if (!set) throw new Error('PIN admin belum diset dalam tab Tetapan');
  if (String(pin || '').trim() !== set) throw new Error('PIN admin salah');
}

// ---------------------------------------------------------------- dashboard
function getDash_(p) {
  const ss = ss_();
  const key = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const addDays = (k, n) => { const d = new Date(k + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const isKey = k => /^\d{4}-\d{2}-\d{2}$/.test(String(k || ''));

  // Julat tarikh: from/to (yyyy-MM-dd) atau days (lalai 14 hari hingga hari ini)
  const today = key(new Date());
  let to = isKey(p.to) ? p.to : today;
  let from = isKey(p.from) ? p.from : addDays(to, -((Number(p.days) || 14) - 1));
  if (from > to) { const t = from; from = to; to = t; }
  if (addDays(from, 180) < to) from = addDays(to, -180);
  let days = 0; for (let k = from; k <= to; k = addDays(k, 1)) days++;

  const rk = ss.getSheetByName('Rekod');
  const last = rk.getLastRow();
  const all = last > 1 ? rk.getRange(2, 1, last - 1, 9).getValues() : [];

  const SET_PARTS = ['getaran-aura-ga', 'kasturi-kijang-kk'];
  const SET_AUTO_DARI = '2026-11-12';   // mulai tarikh ini Set Sinaran tolak GA & KK automatik

  const outBy = {}, inBy = {}, perDay = {}, returns = [], usedBy = {}, rows = [];
  for (let k = from; k <= to; k = addDays(k, 1)) perDay[k] = 0;

  all.forEach(r => {
    if (!(r[1] instanceof Date)) return;
    const k = key(r[1]);
    if (k < from || k > to) return;
    const step = r[2], id = r[3], q = Number(r[5]) || 0;
    if (step === 'Baki dilapor') return;
    rows.push([k, step, id, q, String(r[7] || '').replace(/\s*\((app|sheet|pembetulan)\)\s*$/i, ''), String(r[8] || '')]);
    if (step === 'Stock out') {
      if (id === 'set-sinaran') {
        outBy[id] = (outBy[id] || 0) + q;
        if (k >= SET_AUTO_DARI) SET_PARTS.forEach(x => outBy[x] = (outBy[x] || 0) + q);
      } else {
        outBy[id] = (outBy[id] || 0) + q;
        perDay[k] += q;
      }
    } else if (step === 'Stock in') {
      inBy[id] = (inBy[id] || 0) + q;
    } else if (step === 'Guna bahan') {
      usedBy[id] = (usedBy[id] || 0) + q;
    } else if (step === 'Masuk balik') {
      returns.push({ date: k, id: id, name: r[4], qty: q, by: r[7] });
    }
  });

  // Bahan ikut resepi = Stock in produk x kuantiti seunit
  const rs = ss.getSheetByName('Resepi');
  const rsRows = rs.getRange(2, 1, Math.max(rs.getLastRow() - 1, 1), 4).getValues();
  rsRows.forEach(r => {
    if (!r[0] || !r[2]) return;
    const made = inBy[r[0]] || 0;
    if (made) usedBy[r[2]] = (usedBy[r[2]] || 0) + made * (Number(r[3]) || 0);
  });
  rows.forEach(r => { if (r[1] === 'Stock out' && r[2] === 'set-sinaran' && r[0] >= SET_AUTO_DARI) usedBy['kotak-sinaran'] = (usedBy['kotak-sinaran'] || 0) + r[3]; });

  const avg = {}, avgUse = {};
  Object.keys(outBy).forEach(id => avg[id] = outBy[id] / days);
  Object.keys(usedBy).forEach(id => avgUse[id] = usedBy[id] / days);

  return {
    days: days, from: from, to: to, today: today,
    perDay: Object.keys(perDay).map(k => ({ date: k, out: perDay[k] })),
    outBy: outBy, inBy: inBy, avgOut: avg, avgUse: avgUse,
    returns: returns.slice(-20),
    rows: rows   // [tarikh, langkah, id, kuantiti, oleh, nota]
  };
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return v === '' || v == null ? '' : String(v);
}
function num_(v) { const n = Number(v); return isFinite(n) ? n : null; }

// ---------------------------------------------------------------- baca
function getData_() {
  const ss = ss_();
  const pr = ss.getSheetByName('Produk');
  const prRows = pr.getRange(2, 1, Math.max(pr.getLastRow() - 1, 1), 10).getValues();
  const produk = prRows.filter(r => r[0]).map(r => ({
    id: r[0], name: r[1], group: r[2], price: num_(r[3]),
    count: num_(r[4]), countDate: fmtDate_(r[5]),
    in: num_(r[6]), out: num_(r[7]), ret: num_(r[8]), baki: num_(r[9])
  }));

  const bh = ss.getSheetByName('Bahan');
  const bhRows = bh.getRange(2, 1, Math.max(bh.getLastRow() - 1, 1), 12).getValues();
  const bahan = bhRows.filter(r => r[0] && r[2] !== 'Kos proses').map(r => ({
    id: r[0], name: r[1], group: r[2], unit: r[3], min: num_(r[4]),
    count: num_(r[5]), countDate: fmtDate_(r[6]),
    in: num_(r[7]), used: num_(r[8]), baki: r[9] === '' ? null : num_(r[9]), status: r[10]
  }));

  const st = ss.getSheetByName('Staff');
  // Lajur D "Dashboard" = Ya: staf ini nampak butang Dashboard (masih perlu PIN admin)
  const staff = st.getRange(2, 1, Math.max(st.getLastRow() - 1, 1), 4).getValues()
    .filter(r => r[0]).map(r => ({ name: r[0], role: r[1], dash: String(r[3]).trim().toLowerCase() === 'ya' }));

  const rk = ss.getSheetByName('Rekod');
  const last = rk.getLastRow();
  const n = Math.min(60, Math.max(last - 1, 0));
  const recent = n ? rk.getRange(last - n + 1, 1, n, 9).getValues().reverse().map(r => ({
    time: r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd HH:mm') : String(r[0]),
    date: fmtDate_(r[1]), step: r[2], id: r[3], name: r[4], qty: num_(r[5]), by: r[7], note: r[8]
  })) : [];

  return { produk, bahan, staff, recent, now: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') };
}

// ---------------------------------------------------------------- tulis
function record_(p) {
  const step = String(p.step || '');
  if (!STEPS[step]) throw new Error('Langkah tidak sah: ' + step);
  const by = String(p.by || '').trim();
  if (!by) throw new Error('Pilih nama staf dulu');
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(p.date || ''));
  if (!m) throw new Error('Tarikh tidak sah');
  const date = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const items = (p.items || []).map(i => ({ id: String(i.id), qty: Number(i.qty) }))
    .filter(i => i.id && isFinite(i.qty) && (step === 'Kira stok' ? i.qty >= 0 : i.qty > 0));
  if (!items.length) throw new Error('Tiada kuantiti diisi');

  // elak hantar dua kali (contoh: tekan butang dua kali / internet lambat)
  const cache = CacheService.getScriptCache();
  if (p.key) {
    if (cache.get('k_' + p.key)) return { duplicate: true, rows: 0 };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = ss_();
    const pr = ss.getSheetByName('Produk'), bh = ss.getSheetByName('Bahan');
    const prIds = pr.getRange(2, 1, Math.max(pr.getLastRow() - 1, 1), 2).getValues();
    const bhIds = bh.getRange(2, 1, Math.max(bh.getLastRow() - 1, 1), 2).getValues();
    const find = (list, id) => { for (let k = 0; k < list.length; k++) if (list[k][0] === id) return { row: k + 2, name: list[k][1] }; return null; };

    const kind = STEPS[step];
    const now = new Date();
    const note = String(p.note || '').slice(0, 300);
    const rows = [];
    items.forEach(it => {
      let hit = null, sheet = null;
      if (kind !== 'bahan') { hit = find(prIds, it.id); if (hit) sheet = 'produk'; }
      if (!hit && kind !== 'produk') { hit = find(bhIds, it.id); if (hit) sheet = 'bahan'; }
      if (!hit) throw new Error('Item tidak dijumpai: ' + it.id);

      if (step === 'Kira stok') {
        // reset kiraan: Produk E:F, Bahan F:G
        if (sheet === 'produk') pr.getRange(hit.row, 5, 1, 2).setValues([[it.qty, date]]);
        else bh.getRange(hit.row, 6, 1, 2).setValues([[it.qty, date]]);
      }
      rows.push([now, date, step, it.id, hit.name, it.qty, '', by + ' (app)', note]);
    });

    const rk = ss.getSheetByName('Rekod');
    const start = rk.getLastRow() + 1;
    rk.getRange(start, 1, rows.length, 9).setValues(rows);
    rk.getRange(start, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    rk.getRange(start, 2, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    SpreadsheetApp.flush();
    if (p.key) cache.put('k_' + p.key, '1', 21600);
    return { duplicate: false, rows: rows.length };
  } finally {
    lock.releaseLock();
  }
}

// Jalankan sekali dari editor untuk beri kebenaran & semak sambungan
function test() {
  const d = getData_();
  Logger.log('Produk: %s, Bahan: %s, Staff: %s', d.produk.length, d.bahan.length, d.staff.length);
}
