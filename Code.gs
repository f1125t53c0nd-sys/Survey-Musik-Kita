/**
 * Code.gs — Survey Musik Kita
 * Membutuhkan: App.gs (APP_CONFIG) dan index.html
 */

// Urutan kolom di sheet "Respon"
const HEADERS = ['Waktu', 'Nama', 'Usia', 'Genre', 'Artis', 'Platform', 'Frekuensi',
                 'Durasi', 'Waktu Favorit', 'Aktivitas', 'Kepuasan', 'Kepentingan'];

// Nilai tengah tiap pilihan durasi (menit) untuk menghitung rata-rata
const DURASI_MENIT = {
  'Kurang dari 30 menit': 15,
  '30 menit - 1 jam': 45,
  '1 - 2 jam': 90,
  '2 - 4 jam': 180,
  'Lebih dari 4 jam': 300
};

/** Menampilkan halaman web (file index.html) */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Jalankan SEKALI dari editor untuk membuat sheet + memberi izin akses */
function setup() {
  getSheet_();
  Logger.log('Siap! Sheet "' + APP_CONFIG.SHEET_NAME + '" sudah tersedia.');
}

/** Opsional: jalankan untuk menambah 1 data contoh (tes koneksi) */
function testTambahData() {
  submitSurvey({
    nama: 'Tes', usia: '13 - 15', genre: ['Pop', 'K-Pop'], artis: 'NIKI, BTS',
    platform: ['Spotify'], frekuensi: 'Setiap hari', durasi: '1 - 2 jam',
    waktu: ['Malam'], aktivitas: ['Belajar'], kepuasan: 4, penting: 5
  });
  Logger.log('Data contoh berhasil ditambahkan.');
}

/** Membuka sheet (dan membuatnya jika belum ada) */
function getSheet_() {
  const url = APP_CONFIG.SPREADSHEET_URL;
  const ss = (url && url.indexOf('docs.google.com') > -1)
    ? SpreadsheetApp.openByUrl(url)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet tidak ditemukan. Isi SPREADSHEET_URL di App.gs.');

  let sh = ss.getSheetByName(APP_CONFIG.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(APP_CONFIG.SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#ffd6e7');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Dipanggil dari form: validasi lalu simpan 1 baris ke sheet */
function submitSurvey(d) {
  const clean = v => {
    const s = String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').trim().slice(0, 150);
    return /^[=+\-@]/.test(s) ? ' ' + s : s;   // cegah formula injection
  };
  const list = v => (Array.isArray(v) ? v : []).map(clean).filter(Boolean);

  const genre = list(d.genre), platform = list(d.platform);
  const waktu = list(d.waktu), aktivitas = list(d.aktivitas);
  const artis = String(d.artis || '').split(/[,;\n]/).map(clean).filter(Boolean).slice(0, 3);
  const kep = Number(d.kepuasan), pen = Number(d.penting);

  if (!d.usia || !genre.length || genre.length > 3 || !artis.length || !platform.length ||
      !d.frekuensi || !(d.durasi in DURASI_MENIT) || !waktu.length || !aktivitas.length ||
      !(kep >= 1 && kep <= 5) || !(pen >= 1 && pen <= 5)) {
    throw new Error('Data belum lengkap. Cek kembali semua pertanyaan.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    getSheet_().appendRow([
      new Date(), clean(d.nama) || 'Anonim', clean(d.usia), genre.join(', '), artis.join(', '),
      platform.join(', '), clean(d.frekuensi), clean(d.durasi), waktu.join(', '),
      aktivitas.join(', '), kep, pen
    ]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/** Menghitung semua statistik dari data terbaru di sheet */
function getStats() {
  const sh = getSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return { total: 0 };

  const rows = sh.getRange(2, 1, n, HEADERS.length).getValues();
  const col = i => rows.map(r => r[i]);
  const split = (vals, re) => vals.reduce((a, v) =>
    a.concat(String(v).split(re).map(s => s.trim()).filter(Boolean)), []);

  // hitung kemunculan, urut dari terbanyak
  const tally = (items, ignoreCase) => {
    const m = {};
    items.forEach(s => {
      const k = ignoreCase ? s.toLowerCase() : s;
      if (!m[k]) m[k] = { name: s, count: 0 };
      m[k].count++;
    });
    return Object.keys(m).map(k => m[k])
      .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
  };
  const avg = arr => arr.reduce((a, b) => a + Number(b), 0) / arr.length;

  const menit = col(7).map(v => DURASI_MENIT[v]).filter(Boolean);

  return {
    total: n,
    genre: tally(split(col(3), /,\s*/)),
    artis: tally(split(col(4), /,\s*/), true),
    platform: tally(split(col(5), /,\s*/)),
    frekuensi: tally(col(6)),
    durasi: tally(col(7)),
    waktu: tally(split(col(8), /,\s*/)),
    aktivitas: tally(split(col(9), /,\s*/)),
    usia: tally(col(2)),
    avgMenit: menit.length ? avg(menit) : 0,
    avgKepuasan: avg(col(10)),
    avgPenting: avg(col(11)),
    updated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy, HH:mm')
  };
}
