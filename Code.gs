/**
 * BACKEND MODUL MINGGU 6 — Google Apps Script (container-bound ke Google Sheets)
 *
 * LANGKAH PASANG:
 * 1. Buat Google Sheet baru > Extensions > Apps Script.
 * 2. Tempel file ini sebagai Code.gs. Buat file HTML bernama "Index" (tanpa .html), tempel isi Index.html.
 * 3. Deploy > New deployment > Web app:
 *      Execute as: Me   |   Who has access: Anyone
 * 4. Salin URL berakhiran /exec.
 * 5. Google Sites: Insert > Embed > By URL > tempel URL /exec > Insert.
 * Setiap mengubah kode, buat versi baru via Deploy > Manage deployments > Edit > New version.
 *
 * Sheet dibuat otomatis: Akses, Jawaban, Rekap.
 */
const KUNCI = 'BBBBCBBCBB'; // kunci pilihan ganda soal 1-10 (harus sama dengan Index.html)

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Modul Minggu 6 - Kabin dan Keselamatan B737-200')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sh_(nama, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(nama);
  if (!s) { s = ss.insertSheet(nama); s.appendRow(header); s.setFrozenRows(1); s.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#d9e2ec'); }
  return s;
}
const H_AKSES = ['Session ID', 'Waktu Akses', 'Nama', 'NIT/NPM', 'Kelas', 'Terakhir Aktif', 'Durasi Akses (detik)', 'Perangkat'];
const H_JAWAB = ['Waktu Kirim', 'Session ID', 'Nama', 'NIT/NPM', 'Kelas', 'Percobaan ke-', 'Mulai', 'Durasi Pengerjaan (detik)', 'Jawaban PG', 'Skor PG (benar)', 'Nilai PG (0-100)', 'Uraian 1', 'Uraian 2', 'Uraian 3', 'Nilai Uraian (isi dosen)'];
const H_REKAP = ['NIT/NPM', 'Nama', 'Kelas', 'Jumlah Sesi Akses', 'Akses Pertama', 'Akses Terakhir', 'Total Durasi Akses (menit)', 'Status Evaluasi', 'Jumlah Percobaan', 'Nilai PG Terbaik', 'Nilai PG Terakhir', 'Waktu Kirim Terakhir'];

function withLock_(fn) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function logAkses(p, sid, ua) {
  return withLock_(() => {
    const now = new Date();
    sh_('Akses', H_AKSES).appendRow([sid, now, p.nama, "'" + p.nit, p.kelas, now, 0, ua || '']);
    upsertRekap_(p, { akses: now });
    return { ok: true };
  });
}

function ping(sid, detik) {
  return withLock_(() => {
    const s = sh_('Akses', H_AKSES), v = s.getRange(1, 1, s.getLastRow(), 1).getValues();
    for (let i = v.length - 1; i > 0; i--) if (v[i][0] === sid) {
      s.getRange(i + 1, 6, 1, 2).setValues([[new Date(), detik]]);
      const r = s.getRange(i + 1, 3, 1, 3).getValues()[0];
      upsertRekap_({ nama: r[0], nit: String(r[1]).replace("'", ''), kelas: r[2] }, { sid: sid });
      break;
    }
    return { ok: true };
  });
}

function simpanJawaban(d) {
  return withLock_(() => {
    const p = d.profil, now = new Date(), nit = String(p.nit);
    const jw = String(d.jawaban || '').toUpperCase();
    let benar = 0; for (let i = 0; i < KUNCI.length; i++) if (jw[i] === KUNCI[i]) benar++;
    const nilai = Math.round(benar / KUNCI.length * 100);
    const s = sh_('Jawaban', H_JAWAB), data = s.getLastRow() > 1 ? s.getRange(2, 4, s.getLastRow() - 1, 1).getDisplayValues() : [];
    const ke = data.filter(r => String(r[0]).replace("'", '') === nit).length + 1;
    const u = d.uraian || [];
    s.appendRow([now, d.sid, p.nama, "'" + nit, p.kelas, ke, new Date(d.mulai), d.durasiDetik, jw, benar, nilai, u[0] || '', u[1] || '', u[2] || '', '']);
    upsertRekap_(p, { submit: now, ke: ke, nilai: nilai });
    return { ok: true, skor: benar, nilai: nilai };
  });
}

// Satu baris per taruna (kunci: NIT). Total durasi dihitung dari sheet Akses.
function upsertRekap_(p, o) {
  const s = sh_('Rekap', H_REKAP), nit = String(p.nit);
  const last = s.getLastRow(), v = last > 1 ? s.getRange(2, 1, last - 1, H_REKAP.length).getDisplayValues() : [];
  let idx = v.findIndex(r => String(r[0]).replace("'", '') === nit);
  let row = idx >= 0 ? s.getRange(idx + 2, 1, 1, H_REKAP.length).getValues()[0]
                     : ["'" + nit, p.nama, p.kelas, 0, '', '', 0, 'Belum mengerjakan', 0, '', '', ''];
  row[1] = p.nama; row[2] = p.kelas;
  if (o.akses) { row[3] = (+row[3] || 0) + 1; if (!row[4]) row[4] = o.akses; row[5] = o.akses; }
  if (o.sid) row[5] = new Date();
  if (o.submit) { row[7] = 'Sudah mengerjakan'; row[8] = o.ke; row[10] = o.nilai; row[9] = Math.max(+row[9] || 0, o.nilai); row[11] = o.submit; }
  // total durasi akses (menit) = jumlah durasi seluruh sesi taruna
  const a = sh_('Akses', H_AKSES), n = a.getLastRow();
  if (n > 1) { const av = a.getRange(2, 4, n - 1, 4).getDisplayValues(); let t = 0;
    av.forEach(r => { if (String(r[0]).replace("'", '') === nit) t += +r[3] || 0; }); row[6] = Math.round(t / 60 * 10) / 10; }
  if (idx >= 0) s.getRange(idx + 2, 1, 1, H_REKAP.length).setValues([row]); else s.appendRow(row);
}
