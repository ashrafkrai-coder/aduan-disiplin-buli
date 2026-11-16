/**
 * ADUAN DISIPLIN & BULI — Backend Google Apps Script
 * Pasang: Google Sheet baharu > Extensions > Apps Script > tampal kod ini
 *  1) Run fungsi  setup  (benarkan akses)
 *  2) Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone
 *  3) Salin URL /exec ke dalam index.html (CONFIG.GAS_URL)
 */
const SH_ADUAN = 'Aduan', SH_TETAPAN = 'Tetapan';
const HEADERS = ['ID','Masa Hantar','Jenis','Kategori','Tahap','Tarikh Kejadian','Masa Kejadian','Lokasi',
  'Murid Terlibat','Mangsa','Keterangan','Bukti','Pelapor','Peranan Pelapor','Hubungi Pelapor',
  'Status','Maklum Balas','Catatan Dalaman','Dikemaskini Oleh','Masa Kemaskini'];
const KEYS = ['id','masaHantar','jenis','kategori','tahap','tarikh','masa','lokasi','terlibat','mangsa',
  'keterangan','bukti','pelapor','peranan','hubungi','status','maklumBalas','catatan','oleh','masaKemaskini'];
const STATUS = ['Baru','Dalam Siasatan','Selesai'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Aduan').addItem('Setup / Baiki helaian', 'setup')
    .addItem('Hantar emel ujian', 'emelUjian').addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Kuala_Lumpur');
  let sh = ss.getSheetByName(SH_ADUAN) || ss.insertSheet(SH_ADUAN);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold')
    .setBackground('#1c1c1e').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, sh.getMaxRows() - 1, HEADERS.length).setNumberFormat('@');
  sh.getRange(2, 2, sh.getMaxRows() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange(2, 20, sh.getMaxRows() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  let st = ss.getSheetByName(SH_TETAPAN);
  if (!st) {
    st = ss.insertSheet(SH_TETAPAN);
    const pin = String(Math.floor(100000 + Math.random() * 900000));
    st.getRange(1, 1, 4, 3).setValues([
      ['Kunci', 'Nilai', 'Nota'],
      ['NAMA_SEKOLAH', 'SMK Taman Medan', 'Dipaparkan dalam aplikasi & emel'],
      ['EMEL_GURU', Session.getEffectiveUser().getEmail(), 'Beberapa emel dipisahkan dengan koma'],
      ['PIN_ADMIN', pin, 'PIN guru untuk tab "Guru" (4-12 huruf/nombor)']
    ]);
    st.getRange('A1:C1').setFontWeight('bold');
    st.getRange('B2:B4').setNumberFormat('@');
    st.setColumnWidth(2, 320); st.setColumnWidth(3, 320);
  }
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('FOLDER_ID')) {
    props.setProperty('FOLDER_ID', DriveApp.createFolder('Bukti Aduan Disiplin & Buli').getId());
  }
  kongsiFolder_();
}

/* ---------- Web app ---------- */
function doGet() {
  return json_({ ok: true, app: 'aduan', sekolah: tetapan_().NAMA_SEKOLAH || '' });
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    switch (b.action) {
      case 'submit': return json_(submit_(b));
      case 'status': return json_(status_(b));
      case 'login': semakPin_(b.pin, b.did); return json_({ ok: true, sekolah: tetapan_().NAMA_SEKOLAH || '' });
      case 'list': semakPin_(b.pin, b.did); return json_({ ok: true, rows: senarai_() });
      case 'update': semakPin_(b.pin, b.did); return json_(kemaskini_(b));
      case 'getSettings': semakPin_(b.pin, b.did); {
        const t = tetapan_(); return json_({ ok: true, sekolah: t.NAMA_SEKOLAH || '', emel: t.EMEL_GURU || '' });
      }
      case 'saveSettings': semakPin_(b.pin, b.did); return json_(simpanTetapan_(b));
      case 'testEmail': semakPin_(b.pin, b.did); emelUjian(); return json_({ ok: true });
      default: throw new Error('Tindakan tidak sah');
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Utiliti ---------- */
function tetapan_() {
  const st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TETAPAN);
  const out = {};
  if (!st) return out;
  st.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[0]) out[r[0].trim()] = String(r[1]).trim(); });
  return out;
}
function setTetapan_(kunci, nilai) {
  const st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TETAPAN);
  const v = st.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === kunci) {
    st.getRange(i + 1, 2).setNumberFormat('@').setValue(nilai); return;
  }
  st.appendRow([kunci, nilai, '']);
}
function senaraiEmel_() {
  return (tetapan_().EMEL_GURU || '').split(/[,;\s]+/).filter(x => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));
}
const PIN_MAKS_GAGAL = 3, PIN_KUNCI_SAAT = 600; // 3 kali salah -> kunci 10 minit (setiap peranti)
function semakPin_(pin, did) {
  const cache = CacheService.getScriptCache();
  const kunci = 'pin_gagal_' + (/^[A-Za-z0-9-]{8,64}$/.test(String(did || '')) ? did : 'tanpa-id');
  const gagal = Number(cache.get(kunci) || 0);
  if (gagal >= PIN_MAKS_GAGAL) throw new Error('Terlalu banyak cubaan salah. Peranti ini dikunci 10 minit.');
  const betul = tetapan_().PIN_ADMIN;
  if (!betul || String(pin || '') !== String(betul)) {
    cache.put(kunci, String(gagal + 1), PIN_KUNCI_SAAT);
    const baki = PIN_MAKS_GAGAL - gagal - 1;
    throw new Error(baki > 0 ? 'PIN salah (baki ' + baki + ' cubaan)' : 'PIN salah. Peranti ini dikunci 10 minit.');
  }
  cache.remove(kunci);
}
function bersih_(s, max) { return String(s == null ? '' : s).replace(/\u0000/g, '').trim().slice(0, max || 300); }
function selamat_(s) { s = String(s); return /^[=+\-@]/.test(s) ? "'" + s : s; } // elak formula injection
function esc_(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function idBaru_() {
  const abjad = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; let r = '';
  for (let i = 0; i < 6; i++) r += abjad[Math.floor(Math.random() * abjad.length)];
  return 'ADU-' + Utilities.formatDate(new Date(), 'Asia/Kuala_Lumpur', 'yyMMdd') + '-' + r;
}
function kongsiFolder_() {
  try {
    const id = PropertiesService.getScriptProperties().getProperty('FOLDER_ID'); if (!id) return;
    const f = DriveApp.getFolderById(id);
    senaraiEmel_().forEach(e => { try { f.addViewer(e); } catch (x) {} });
  } catch (x) {}
}

/* ---------- Hantar aduan ---------- */
function submit_(b) {
  const cache = CacheService.getScriptCache();
  const cid = bersih_(b.cid, 60);
  if (cid) { const dah = cache.get('cid_' + cid); if (dah) return { ok: true, id: dah }; }
  const kadar = Number(cache.get('kadar') || 0);
  if (kadar >= 40) throw new Error('Sistem sibuk. Sila cuba sebentar lagi.');
  cache.put('kadar', String(kadar + 1), 600);

  const jenis = b.jenis === 'Buli' ? 'Buli' : 'Disiplin';
  const keterangan = bersih_(b.keterangan, 3000);
  const kategori = bersih_(b.kategori, 80);
  if (keterangan.length < 10) throw new Error('Keterangan terlalu ringkas');
  if (!kategori) throw new Error('Sila pilih kategori');
  const tanpaNama = !!b.tanpaNama;
  const id = idBaru_();
  const r = {
    id: id, jenis: jenis, kategori: kategori, tahap: b.tahap === 'Segera' ? 'Segera' : 'Biasa',
    tarikh: bersih_(b.tarikh, 20), masa: bersih_(b.masa, 10), lokasi: bersih_(b.lokasi, 150),
    terlibat: bersih_(b.terlibat, 500), mangsa: bersih_(b.mangsa, 300), keterangan: keterangan,
    pelapor: tanpaNama ? 'Tanpa nama' : (bersih_(b.pelapor, 120) || 'Tidak dinyatakan'),
    peranan: bersih_(b.peranan, 30), hubungi: tanpaNama ? '' : bersih_(b.hubungi, 60)
  };

  // Simpan gambar bukti (maks 3)
  const blobs = [], pautan = [];
  try {
    const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
    (Array.isArray(b.bukti) ? b.bukti : []).slice(0, 3).forEach((d, i) => {
      const m = /^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(String(d));
      if (!m || m[2].length > 1500000) return;
      const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], id + '-' + (i + 1) + '.jpg');
      pautan.push(folder.createFile(blob).getUrl()); blobs.push(blob);
    });
  } catch (x) {}
  r.bukti = pautan.join('\n');

  const baris = [r.id, new Date(), r.jenis, r.kategori, r.tahap, r.tarikh, r.masa, r.lokasi, r.terlibat, r.mangsa,
    r.keterangan, r.bukti, r.pelapor, r.peranan, r.hubungi, 'Baru', '', '', '', '']
    .map(v => v instanceof Date ? v : selamat_(v));
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ADUAN).appendRow(baris); }
  finally { lock.releaseLock(); }
  if (cid) cache.put('cid_' + cid, id, 21600);

  try { maklumEmel_(r, blobs); } catch (x) { console.error('Emel gagal: ' + x); }
  return { ok: true, id: id };
}

function maklumEmel_(r, blobs) {
  const emel = senaraiEmel_(); if (!emel.length) return;
  const sekolah = tetapan_().NAMA_SEKOLAH || '';
  const warna = r.jenis === 'Buli' ? '#ff3b30' : '#ff9500';
  const baris = (k, v) => v ? '<tr><td style="padding:8px 0;color:#8e8e93;font-size:13px;width:130px;vertical-align:top">' + k +
    '</td><td style="padding:8px 0;font-size:14px;color:#1c1c1e">' + esc_(v).replace(/\n/g, '<br>') + '</td></tr>' : '';
  const html = '<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#f2f2f7;padding:24px">' +
    '<div style="max-width:520px;margin:auto;background:#fff;border-radius:18px;overflow:hidden">' +
    '<div style="background:' + warna + ';color:#fff;padding:20px 24px"><div style="font-size:12px;opacity:.85;letter-spacing:.5px">' +
    esc_(sekolah.toUpperCase()) + '</div><div style="font-size:21px;font-weight:700;margin-top:4px">Aduan ' + r.jenis + ' Baharu' +
    (r.tahap === 'Segera' ? ' · SEGERA' : '') + '</div><div style="font-size:13px;opacity:.9;margin-top:2px">' + r.id + '</div></div>' +
    '<div style="padding:16px 24px 22px"><table style="width:100%;border-collapse:collapse">' +
    baris('Kategori', r.kategori) + baris('Tarikh kejadian', (r.tarikh + ' ' + r.masa).trim()) + baris('Lokasi', r.lokasi) +
    baris('Murid terlibat', r.terlibat) + baris('Mangsa', r.mangsa) + baris('Keterangan', r.keterangan) +
    baris('Pelapor', r.pelapor + (r.peranan ? ' (' + r.peranan + ')' : '')) + baris('Hubungi', r.hubungi) +
    baris('Bukti', blobs.length ? blobs.length + ' gambar dilampirkan' : '') + '</table>' +
    '<a href="' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '" style="display:block;text-align:center;background:#007aff;color:#fff;' +
    'text-decoration:none;padding:13px;border-radius:12px;font-weight:600;margin-top:14px">Buka Google Sheet</a>' +
    '<p style="font-size:12px;color:#8e8e93;margin:14px 0 0">SULIT — maklumat ini untuk tindakan guru sahaja.</p></div></div></div>';
  MailApp.sendEmail({
    to: emel.join(','), name: 'Aduan Disiplin & Buli',
    subject: (r.tahap === 'Segera' ? '🚨 SEGERA · ' : '') + 'Aduan ' + r.jenis + ' — ' + r.kategori + ' [' + r.id + ']',
    htmlBody: html, attachments: blobs
  });
}

function emelUjian() {
  maklumEmel_({ id: 'ADU-UJIAN', jenis: 'Disiplin', kategori: 'Emel ujian', tahap: 'Biasa', tarikh: '', masa: '', lokasi: '',
    terlibat: '', mangsa: '', keterangan: 'Ini emel ujian. Notifikasi aduan berfungsi dengan baik.', pelapor: 'Sistem', peranan: '', hubungi: '' }, []);
}

/* ---------- Semak status (awam) ---------- */
function cariBaris_(id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ADUAN);
  const c = sh.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
  return c ? { sh: sh, row: c.getRow() } : null;
}
function status_(b) {
  const id = bersih_(b.id, 30).toUpperCase();
  if (!/^ADU-\d{6}-[A-Z0-9]{6}$/.test(id)) throw new Error('No. rujukan tidak sah');
  const f = cariBaris_(id); if (!f) throw new Error('Aduan tidak ditemui');
  const v = f.sh.getRange(f.row, 1, 1, HEADERS.length).getDisplayValues()[0];
  return { ok: true, id: v[0], masaHantar: v[1], jenis: v[2], kategori: v[3], status: v[15], maklumBalas: v[16], masaKemaskini: v[19] };
}

/* ---------- Guru ---------- */
function senarai_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ADUAN);
  const n = sh.getLastRow(); if (n < 2) return [];
  const mula = Math.max(2, n - 499);
  return sh.getRange(mula, 1, n - mula + 1, HEADERS.length).getDisplayValues()
    .filter(r => r[0]).map(r => { const o = {}; KEYS.forEach((k, i) => o[k] = r[i]); return o; }).reverse();
}
function kemaskini_(b) {
  const f = cariBaris_(bersih_(b.id, 30)); if (!f) throw new Error('Aduan tidak ditemui');
  const status = STATUS.indexOf(b.status) >= 0 ? b.status : 'Baru';
  f.sh.getRange(f.row, 16, 1, 5).setValues([[status, selamat_(bersih_(b.maklumBalas, 1000)),
    selamat_(bersih_(b.catatan, 2000)), selamat_(bersih_(b.oleh, 80)), new Date()]]);
  return { ok: true };
}
function simpanTetapan_(b) {
  if (b.sekolah != null) setTetapan_('NAMA_SEKOLAH', bersih_(b.sekolah, 120));
  if (b.emel != null) setTetapan_('EMEL_GURU', bersih_(b.emel, 1000).split(/[,;\s]+/).filter(String).join(', '));
  if (b.pinBaru) {
    if (!/^[A-Za-z0-9]{4,12}$/.test(String(b.pinBaru))) throw new Error('PIN mesti 4-12 huruf/nombor');
    setTetapan_('PIN_ADMIN', String(b.pinBaru));
  }
  kongsiFolder_();
  return { ok: true };
}
