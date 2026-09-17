// ============================================================
// CODE.GS GABUNGAN — TBR REGION 10
// Berisi 2 bagian:
//   BAGIAN A: Backend Web App (doGet) — dipanggil dari GitHub Pages
//             (index.html & laporan.html) untuk submit & ambil laporan.
//   BAGIAN B: Rekap Cabang — menu sidebar di Google Sheets untuk
//             memantau cabang yang sudah/belum lapor.
// Cara pasang: ganti SEMUA isi Code.gs di Apps Script dengan file ini,
// lalu Deploy ulang Web App (pilih "New version" di deployment yang sudah ada,
// supaya URL /exec tidak berubah).
// ============================================================


// ============================================================
// BAGIAN A — BACKEND WEB APP (doGet)
// ============================================================

var CONFIG = {
  NAMA_SHEET: "Laporan Harian",         // Nama tab sheet di Google Spreadsheet
  NAMA_SHEET_LOG: "Log Laporan",        // Nama tab log (dibuat otomatis)
  NAMA_SHEET_LIVIN: "Monitoring Livin' Food",  // Nama tab monitoring Livin' Food
  NAMA_SHEET_PIPELINE: "Pipeline EDC to LVM",  // Nama tab pipeline konversi EDC to LVM
  NAMA_SHEET_LEAKAGE: "Pipeline Leakage Top 100"  // Nama tab pipeline leakage top 100
};

// Pilihan dropdown yang valid buat form follow up di monitoring-top100-leakage.html.
var KETERANGAN_LEAKAGE_VALID = [
  '1. Rekening Operasional di Bank lain',
  '2. Rekening Pihak Ketiga di Bank lain (>50%)',
  '3. Memiliki Fasilitas Pinjaman di Bank lain',
  '4. Mendapatkan Special Rate Funding di Bank lain',
  '6. Lainnya'
];
var PENAWARAN_PRIMA_VALID = ['Ya', 'Tidak'];
var SUMBER_PIPELINE_VALID = ['Bottom Up', 'Top Down'];
var KET_KONFIRMASI_VALID = ['Done', 'Belum'];

// Tahapan status progress pipeline akuisisi merchant yang valid.
var STATUS_EDC_LVM_VALID = ["Target", "Penawaran", "Done Konversi to LVM", "Merchant Menolak", "Merchant Tutup", "Perlu Kunjungan MTI", "Tetap pakai EDC & Optimalkan SV"];

/**
 * Semua request dari GitHub Pages masuk ke sini via GET.
 * Parameter "action" menentukan apa yang dilakukan:
 *   - action=submit  → simpan data laporan baru
 *   - action=getData → ambil data laporan berdasarkan ID
 */
function doGet(e) {
  var action = (e.parameter && e.parameter.action) ? e.parameter.action : '';

  // ── MODE MAINTENANCE: kalau aktif, SEMUA action (kecuali cekMaintenance
  //    sendiri) diblokir, biar data di sheet aman gak keubah/kebaca dari
  //    frontend selagi lagi dibenerin manual. Halaman depan (index/pipeline/
  //    laporan/monitoring-area) ngecek ini duluan & ngalihin ke
  //    maintenance.html. Diaktifkan/dimatikan lewat menu "🔧 Maintenance"
  //    di Google Sheets — lihat onOpen(), TANPA perlu buka Apps Script Editor.
  if (action !== 'cekMaintenance' && isMaintenanceAktif()) {
    return jsonResponse({
      maintenance : true,
      pesan       : 'Sistem sedang dalam pemeliharaan data. Coba lagi beberapa saat lagi.'
    });
  }
  if (action === 'cekMaintenance') {
    return jsonResponse({ maintenance: isMaintenanceAktif() });
  }

  // ── SUBMIT: simpan laporan baru ──────────────────────────────
  if (action === 'submit') {
    try {
      var result = prosesSubmit(e.parameter);
      return jsonResponse({ success: true, idLaporan: result.idLaporan });
    } catch (err) {
      return jsonResponse({ success: false, error: err.message });
    }
  }

  // ── GET DATA: ambil data laporan by ID ───────────────────────
  if (action === 'getData') {
    var id = e.parameter.id || '';
    var laporan = cariLaporan(id);
    if (laporan) {
      return jsonResponse(laporan);
    } else {
      return jsonResponse({ error: 'Laporan tidak ditemukan: ' + id });
    }
  }

  // ── GET MERCHANT EDC: ambil daftar merchant EDC (pipeline to LVM)
  //    milik satu cabang, lengkap dengan status progress terkininya ──
  if (action === 'getMerchantEDC') {
    var kodeCabangQ = e.parameter.kodeCabang || '';
    try {
      var merchants = getMerchantEDCByCabang(kodeCabangQ);
      return jsonResponse({ merchants: merchants });
    } catch (err) {
      return jsonResponse({ merchants: [], error: err.message });
    }
  }

  // ── GET PIPELINE ALL: ambil semua data pipeline EDC to LVM,
  //    opsional difilter per kodeCabang (dipakai pipeline.html) ──
  if (action === 'getPipelineAll') {
    var kodeFilter = e.parameter.kodeCabang || '';
    try {
      var pipelineList = getPipelineAll(kodeFilter);
      return jsonResponse({ data: pipelineList });
    } catch (err) {
      return jsonResponse({ data: [], error: err.message });
    }
  }

  // ── UPDATE STATUS: update status progress satu merchant pipeline,
  //    dipanggil langsung dari pipeline.html saat admin klik merchant
  //    lalu pilih status baru di modal ──
  if (action === 'updateStatus') {
    try {
      var ssUpd = SpreadsheetApp.getActiveSpreadsheet();
      var kodeCabangUpd    = e.parameter.kodeCabang   || '';
      var namaMerchantUpd  = e.parameter.namaMerchant || '';
      var alamatUpd        = e.parameter.alamat       || '';
      var midMerchantUpd   = e.parameter.midMerchant  || '';
      var statusBaruUpd    = e.parameter.status       || '';
      var midUpd           = e.parameter.mid          || '';
      var alasanUpd        = e.parameter.alasan       || '';

      if (statusBaruUpd === 'Done Konversi to LVM' && !midUpd.trim()) {
        return jsonResponse({ success: false, error: 'MNDI/MID/Nomor Rekening wajib diisi untuk status Done Konversi to LVM.' });
      }
      if (statusBaruUpd === 'Merchant Menolak' && !alasanUpd.trim()) {
        return jsonResponse({ success: false, error: 'Alasan wajib diisi untuk status Merchant Menolak.' });
      }

      var berhasilUpd = updatePipelineStatus(ssUpd, kodeCabangUpd, namaMerchantUpd, statusBaruUpd, midUpd, alasanUpd, alamatUpd, midMerchantUpd);
      if (berhasilUpd) {
        return jsonResponse({ success: true });
      } else {
        return jsonResponse({ success: false, error: 'Merchant tidak ditemukan atau status tidak valid.' });
      }
    } catch (err) {
      return jsonResponse({ success: false, error: err.message });
    }
  }

  // ── RINGKASAN AREA: data buat monitoring-area.html (rekap per Area
  //    + baris Total Region + Progress DtD) ─────────────────────
  if (action === 'getRingkasanArea') {
    try {
      return jsonResponse(getRingkasanArea());
    } catch (err) {
      return jsonResponse({ error: err.message });
    }
  }

  // ── WA LAPORAN REGION: teks siap kirim buat tombol "Kirim Laporan
  //    Region" di monitoring-area.html ─────────────────────────
  if (action === 'getWaLaporanRegion') {
    try {
      return jsonResponse({ teks: susunWaLaporanRegion() });
    } catch (err) {
      return jsonResponse({ error: err.message });
    }
  }

  // ── WA LAPORAN AREA: teks siap kirim buat tombol "Kirim Laporan"
  //    per baris Area di monitoring-area.html ──────────────────
  if (action === 'getWaLaporanArea') {
    try {
      var namaAreaQ = e.parameter.area || '';
      return jsonResponse({ teks: susunWaLaporanArea(namaAreaQ) });
    } catch (err) {
      return jsonResponse({ error: err.message });
    }
  }

  // ── GET LEAKAGE ALL: ambil semua data Pipeline Leakage Top 100,
  //    opsional difilter per 3 digit kode area (mis. "150" utk Manado) —
  //    dipakai monitoring-top100-leakage.html ─────────────────────
  if (action === 'getLeakageAll') {
    try {
      var kodeAreaFilter = e.parameter.area || '';
      return jsonResponse({ data: getLeakageAll(kodeAreaFilter) });
    } catch (err) {
      return jsonResponse({ data: [], error: err.message });
    }
  }

  // ── DEBUG: bandingkan header ASLI di sheet "Pipeline Leakage Top 100"
  //    vs. HEADER_LEAKAGE yang di-hardcode di Code.gs — dipakai buat
  //    diagnosa kalau ada kolom yang kebaca 0/'-' padahal sheet-nya keisi.
  //    Aman dihapus kapan saja, tidak dipakai oleh halaman manapun.
  if (action === 'debugLeakageHeaders') {
    var ssDebug = SpreadsheetApp.getActiveSpreadsheet();
    var sheetDebug = ssDebug.getSheetByName(CONFIG.NAMA_SHEET_LEAKAGE);
    if (!sheetDebug) return jsonResponse({ error: 'Sheet tidak ditemukan: ' + CONFIG.NAMA_SHEET_LEAKAGE });
    var headersAsli = sheetDebug.getRange(1, 1, 1, sheetDebug.getLastColumn()).getValues()[0];
    var perbandingan = HEADER_LEAKAGE.map(function (h) {
      return { diharapkan: h, ditemukan: headersAsli.indexOf(h) !== -1 };
    });
    return jsonResponse({ headersAsli: headersAsli, perbandingan: perbandingan });
  }

  // ── UPDATE LEAKAGE FOLLOW UP: simpan hasil follow up satu merchant
  //    leakage (dipanggil dari modal monitoring-top100-leakage.html) ──
  if (action === 'updateLeakageFollowUp') {
    try {
      var hasilUpd = updateLeakageFollowUp(e.parameter);
      if (hasilUpd.error) return jsonResponse({ success: false, error: hasilUpd.error });
      return jsonResponse({ success: true });
    } catch (err) {
      return jsonResponse({ success: false, error: err.message });
    }
  }

  // ── EXPORT LEAKAGE EXCEL: bikin/refresh 1 spreadsheet snapshot khusus
  //    berisi data Leakage terkini (tanpa tab lain di file utama), lalu
  //    balikin link exportnya buat tombol "Download Excel" ───────────
  if (action === 'exportLeakageExcel') {
    try {
      return jsonResponse({ url: exportLeakageExcelFile() });
    } catch (err) {
      return jsonResponse({ error: err.message });
    }
  }

  // ── Default: API aktif ───────────────────────────────────────
  return jsonResponse({ status: 'API aktif', versi: '2.5' });
}

// Helper: buat JSON response
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// PROSES SUBMIT: simpan data ke Google Sheet
// ============================================================
function prosesSubmit(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss);

  // Ambil & validasi data dari form
  var tanggal          = p.tanggal          || '';
  var area             = p.area             || '';
  var kodeCabang       = p.kodeCabang       || '';
  var namaCabang       = p.namaCabang       || '';
  var jumlahLVM        = Number(p.jumlahLVM)            || 0;
  var jumlahEDC        = Number(p.jumlahEDC)            || 0;
  var jumlahEDCPOT     = Number(p.jumlahEDCPOT)         || 0;
  var jumlahPemasLVM   = Number(p.jumlahPemasanganLVM)  || 0;
  var jumlahRetensiEDC = Number(p.jumlahRetensiEDC)     || 0;
  var totalLeadsCakra     = Number(p.totalLeadsCakra)     || 0;
  var totalKunjunganCakra = Number(p.totalKunjunganCakra) || 0;
  var gapCakra             = Number(p.gapCakra)            || 0;
  var jumlahLivinFood      = Number(p.jumlahTransaksiLivinFood) || 0;
  var namaMerchantLivinFood = p.namaMerchantLivinFood || '-';
  var kendala          = p.kendala          || '-';
  var keterangan       = p.keterangan       || '-';
  var totalAkuisisi    = jumlahLVM + jumlahEDC + jumlahEDCPOT;
  var tanggalFormatted = formatTanggal(tanggal);
  var timestamp        = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

  // Ambil OTOMATIS semua merchant Pipeline EDC to LVM milik cabang ini yang
  // statusnya diupdate hari yang sama dengan tanggal laporan (lewat pipeline.html) —
  // jadi tidak dibatasi cuma 1 merchant, semua kunjungan hari itu ikut kebawa.
  var updatesEDCLVM = getPipelineUpdatesHariIni(kodeCabang, tanggal);

  // Generate ID unik
  var idLaporan = generateID(kodeCabang);

  // Susun teks WA
  var teksWA = susunTeksWA(
    tanggalFormatted, area, namaCabang, kodeCabang,
    jumlahLVM, jumlahEDC, jumlahEDCPOT, totalAkuisisi,
    jumlahPemasLVM, jumlahRetensiEDC,
    totalLeadsCakra, totalKunjunganCakra, gapCakra,
    jumlahLivinFood, namaMerchantLivinFood,
    updatesEDCLVM,
    kendala, keterangan
  );

  // URL laporan (di GitHub Pages) — diisi kosong, ditentukan di sisi frontend
  var linkLaporan = '';

  // Ringkasan update EDC to LVM hari ini, buat ditulis ke sheet "Laporan Harian"
  // (kolom lama dipertahankan supaya nggak perlu migrasi skema baru lagi).
  var ringkasanMerchantEDCLVM = updatesEDCLVM.length
    ? updatesEDCLVM.map(function (u) { return u.namaMerchant + ' (' + u.status + ')'; }).join('; ')
    : '-';
  var ringkasanStatusEDCLVM = updatesEDCLVM.length
    ? updatesEDCLVM.length + ' merchant diupdate'
    : '-';

  // Susun baris berdasarkan NAMA HEADER asli di sheet (bukan posisi tetap),
  // supaya aman walau urutan kolom "Jumlah Akuisisi EDC POT" berbeda-beda.
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nilaiKolom = {
    'Timestamp'               : timestamp,
    'Tanggal Laporan'         : tanggal,
    'Area'                    : area,
    'Nama Cabang'             : namaCabang,
    'Kode Cabang'             : kodeCabang,
    'Jumlah Akuisisi LVM'     : jumlahLVM,
    'Jumlah Akuisisi EDC'     : jumlahEDC,
    'Jumlah Akuisisi EDC POT' : jumlahEDCPOT,
    'Jumlah Pemasangan LVM'   : jumlahPemasLVM,
    'Jumlah Retensi EDC'      : jumlahRetensiEDC,
    'Total Akuisisi'          : totalAkuisisi,
    'Total Leads Cakra'       : totalLeadsCakra,
    'Total Kunjungan Cakra'   : totalKunjunganCakra,
    'Gap (sesuai Cakra)'      : gapCakra,
    "Jumlah Transaksi Livin' Food" : jumlahLivinFood,
    "Nama Merchant Livin' Food"    : namaMerchantLivinFood,
    'Merchant EDC to LVM'          : ringkasanMerchantEDCLVM,
    'Status Progress EDC to LVM'   : ringkasanStatusEDCLVM,
    'Kendala'                 : kendala,
    'Keterangan'              : keterangan,
    'ID Laporan'              : idLaporan,
    'Teks WA'                 : teksWA,
    'Link Laporan'            : linkLaporan
  };

  var row = headers.map(function (h) {
    return nilaiKolom.hasOwnProperty(h) ? nilaiKolom[h] : '';
  });

  sheet.appendRow(row);

  // Update monitoring Livin' Food (jika eligible)
  if (namaMerchantLivinFood && namaMerchantLivinFood !== '-') {
    try {
      updateMonitoringLivinFood(ss, tanggal, kodeCabang, namaMerchantLivinFood, jumlahLivinFood);
    } catch (e) {
      Logger.log('Monitoring Livin Food error: ' + e.message);
    }
  }

  // Catatan: update status Pipeline EDC to LVM sekarang dilakukan LANGSUNG
  // dari pipeline.html (action=updateStatus) saat merchant diklik & disimpan —
  // bukan lagi lewat form laporan harian ini. Baris ini hanya MEMBACA hasil
  // update yang sudah tersimpan di sheet (lihat updatesEDCLVM di atas).

  // Catat log
  catatLog(ss, idLaporan, tanggalFormatted, namaCabang, kodeCabang);

  Logger.log('Submit berhasil: ' + idLaporan);
  return { idLaporan: idLaporan, success: true };
}

// ============================================================
// CARI LAPORAN BERDASARKAN ID
// ============================================================
function cariLaporan(idLaporan) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET);
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  var headers = data[0];
  var idxID = headers.indexOf('ID Laporan');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxID] === idLaporan) {
      var row = data[i];
      return {
        idLaporan           : row[headers.indexOf('ID Laporan')],
        timestamp           : formatTanggal(row[headers.indexOf('Timestamp')]),
        tanggal             : formatTanggal(row[headers.indexOf('Tanggal Laporan')]),
        area                : row[headers.indexOf('Area')]               || '-',
        namaCabang          : row[headers.indexOf('Nama Cabang')]        || '-',
        kodeCabang          : row[headers.indexOf('Kode Cabang')]        || '-',
        jumlahLVM           : Number(row[headers.indexOf('Jumlah Akuisisi LVM')])       || 0,
        jumlahEDC           : Number(row[headers.indexOf('Jumlah Akuisisi EDC')])       || 0,
        jumlahEDCPOT        : Number(row[headers.indexOf('Jumlah Akuisisi EDC POT')])   || 0,
        jumlahPemasanganLVM : Number(row[headers.indexOf('Jumlah Pemasangan LVM')])     || 0,
        jumlahRetensiEDC    : Number(row[headers.indexOf('Jumlah Retensi EDC')])        || 0,
        total               : Number(row[headers.indexOf('Total Akuisisi')])             || 0,
        totalLeadsCakra     : Number(row[headers.indexOf('Total Leads Cakra')])          || 0,
        totalKunjunganCakra : Number(row[headers.indexOf('Total Kunjungan Cakra')])      || 0,
        gapCakra            : Number(row[headers.indexOf('Gap (sesuai Cakra)')])         || 0,
        jumlahTransaksiLivinFood : Number(row[headers.indexOf("Jumlah Transaksi Livin' Food")]) || 0,
        namaMerchantLivinFood    : row[headers.indexOf("Nama Merchant Livin' Food")] || '-',
        // Diambil ULANG dari sheet Pipeline (bukan dibaca dari kolom ringkasan di
        // sheet ini), supaya kalau ada update susulan buat tanggal yang sama,
        // laporan.html tetap nampilin data terkini.
        updatesEDCLVM : getPipelineUpdatesHariIni(
          row[headers.indexOf('Kode Cabang')],
          row[headers.indexOf('Tanggal Laporan')]
        ),
        kendala             : row[headers.indexOf('Kendala')]            || '-',
        keterangan          : row[headers.indexOf('Keterangan')]         || '-',
        teksWA              : row[headers.indexOf('Teks WA')]            || ''
      };
    }
  }
  return null;
}

// ============================================================
// SUSUN TEKS WHATSAPP
// ============================================================
function susunTeksWA(tanggal, area, cabang, kode, lvm, edc, edcPot, total, plasLVM, retEDC, leadsCakra, kunjunganCakra, gapCakra, livinFood, namaMerchant, updatesEDCLVM, kendala, keterangan) {
  var t = '';
  t += 'Mohon izin melaporkan hasil akuisisi harian:\n\n';
  t += '```\n';
  t += 'Tanggal : ' + tanggal + '\n';
  t += 'Area    : ' + area + '\n';
  t += 'Cabang  : ' + cabang + '\n';
  t += 'Kode    : ' + kode + '\n\n';
  t += '📊 Hasil Akuisisi:\n';
  t += '• LVM         : ' + lvm + '\n';
  t += '• EDC         : ' + edc + '\n';
  t += '• EDC POT     : ' + edcPot + '\n';
  t += '• Total       : ' + total + '\n\n';
  t += '🔧 Pemasangan & Retensi:\n';
  t += '• Pemasangan LVM : ' + plasLVM + '\n';
  t += '• Retensi EDC    : ' + retEDC + '\n\n';
  t += '📊 Hasil Kunjungan Cakra:\n';
  t += '• Leads     : ' + leadsCakra + '\n';
  t += '• Realisasi : ' + kunjunganCakra + '\n';
  t += '• Gap       : ' + gapCakra + '\n\n';
  t += "🛒 Livin' Food:\n";
  if (namaMerchant && namaMerchant !== '-') {
    t += '• Merchant  : ' + namaMerchant + '\n';
  }
  t += '• Transaksi : ' + livinFood + '\n\n';
  if (updatesEDCLVM && updatesEDCLVM.length) {
    t += '🔄 Progress Konversi EDC to LVM (' + updatesEDCLVM.length + ' merchant):\n';
    updatesEDCLVM.forEach(function (u) {
      var baris = '• ' + u.namaMerchant + ' — ' + u.status;
      if (u.status === 'Done Konversi to LVM' && u.mid && u.mid !== '-') {
        baris += ' (MID: ' + u.mid + ')';
      }
      if (u.status === 'Merchant Menolak' && u.alasan && u.alasan !== '-') {
        baris += ' (Alasan: ' + u.alasan + ')';
      }
      t += baris + '\n';
    });
    t += '\n';
  }
  t += 'Kendala    : ' + kendala + '\n';
  t += 'Keterangan : ' + keterangan + '\n';
  t += '```\n\n';
  t += 'Terima kasih.';
  return t;
}

// ============================================================
// HELPER FUNCTIONS (BAGIAN A)
// ============================================================
function generateID(kodeCabang) {
  var now = new Date();
  var tgl = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd');
  var rnd = Math.floor(1000 + Math.random() * 9000);
  var kode = (kodeCabang || 'XXX').toString().replace(/\s/g, '').toUpperCase();
  return 'RPT-' + kode + '-' + tgl + '-' + rnd;
}

function formatTanggal(tgl) {
  if (!tgl) return '-';
  try {
    var d = new Date(tgl);
    if (isNaN(d.getTime())) return tgl.toString();
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } catch (e) { return tgl.toString(); }
}

function catatLog(ss, idLaporan, tanggal, cabang, kode) {
  try {
    var log = ss.getSheetByName(CONFIG.NAMA_SHEET_LOG);
    if (!log) {
      log = ss.insertSheet(CONFIG.NAMA_SHEET_LOG);
      log.getRange(1, 1, 1, 5).setValues([['Timestamp Log', 'ID Laporan', 'Tanggal', 'Cabang', 'Kode']]);
      log.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
    }
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    log.appendRow([now, idLaporan, tanggal, cabang, kode]);
  } catch (e) { Logger.log('Log error: ' + e.message); }
}

// ============================================================
// SETUP: Buat sheet dengan header yang benar (kalau belum ada)
// ============================================================
function getOrCreateSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.NAMA_SHEET);
    var headers = [
      'Timestamp', 'Tanggal Laporan', 'Area', 'Nama Cabang', 'Kode Cabang',
      'Jumlah Akuisisi LVM', 'Jumlah Akuisisi EDC', 'Jumlah Akuisisi EDC POT',
      'Jumlah Pemasangan LVM', 'Jumlah Retensi EDC', 'Total Akuisisi',
      'Total Leads Cakra', 'Total Kunjungan Cakra', 'Gap (sesuai Cakra)',
      "Jumlah Transaksi Livin' Food",
      "Nama Merchant Livin' Food",
      'Kendala', 'Keterangan', 'ID Laporan', 'Teks WA', 'Link Laporan'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Jalankan fungsi ini sekali dari Apps Script Editor
 * untuk membuat sheet dengan header yang benar (kalau sheet belum ada).
 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheet(ss);
  SpreadsheetApp.getUi().alert('✅ Sheet "' + CONFIG.NAMA_SHEET + '" berhasil disiapkan!');
}

/**
 * Jalankan fungsi ini SEKALI dari Apps Script Editor (pilih dari dropdown
 * fungsi, lalu klik Run) untuk otomatis menambahkan kolom
 * "Jumlah Akuisisi EDC POT" ke sheet "Laporan Harian" kalau belum ada.
 * Aman dijalankan berkali-kali — kalau kolomnya sudah ada, tidak akan
 * dibuat dobel, cuma kasih pesan "sudah ada".
 */
function tambahKolomEDCPOT() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET); // "Laporan Harian"
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET + '" tidak ditemukan!');
    return;
  }

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var namaKolomBaru = 'Jumlah Akuisisi EDC POT';

  // Kalau kolom sudah ada, tidak usah buat lagi
  if (headers.indexOf(namaKolomBaru) !== -1) {
    beriTahu('✅ Kolom "' + namaKolomBaru + '" sudah ada, tidak perlu ditambah lagi.');
    return;
  }

  // Cari posisi kolom "Jumlah Akuisisi EDC" supaya kolom baru disisipkan tepat setelahnya
  var idxEDC = headers.indexOf('Jumlah Akuisisi EDC'); // 0-based
  var posisiSisip;

  if (idxEDC !== -1) {
    sheet.insertColumnAfter(idxEDC + 1); // idxEDC+1 = posisi 1-based kolom EDC
    posisiSisip = idxEDC + 2;            // posisi 1-based kolom baru
  } else {
    // Kalau kolom "Jumlah Akuisisi EDC" tidak ditemukan, tambahkan di paling akhir
    sheet.insertColumnAfter(lastCol);
    posisiSisip = lastCol + 1;
  }

  sheet.getRange(1, posisiSisip).setValue(namaKolomBaru);
  sheet.getRange(1, posisiSisip)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff');

  beriTahu('✅ Kolom "' + namaKolomBaru + '" berhasil ditambahkan di posisi kolom ' + posisiSisip + '.');
}

/**
 * Jalankan fungsi ini SEKALI dari Apps Script Editor untuk menambahkan
 * 3 kolom baru ke sheet "Laporan Harian": "Total Leads Cakra",
 * "Total Kunjungan Cakra", dan "Gap (sesuai Cakra)" — kalau belum ada.
 * Aman dijalankan berkali-kali, kolom yang sudah ada tidak akan dibuat dobel.
 */
function tambahKolomCakra() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET); // "Laporan Harian"
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET + '" tidak ditemukan!');
    return;
  }

  var kolomBaru = ['Total Leads Cakra', 'Total Kunjungan Cakra', 'Gap (sesuai Cakra)'];
  var ditambahkan = [];

  kolomBaru.forEach(function (nama) {
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers.indexOf(nama) !== -1) return; // sudah ada, lewati

    sheet.insertColumnAfter(lastCol);
    sheet.getRange(1, lastCol + 1).setValue(nama);
    sheet.getRange(1, lastCol + 1)
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('#ffffff');
    ditambahkan.push(nama);
  });

  if (ditambahkan.length === 0) {
    beriTahu('✅ Semua kolom Cakra sudah ada, tidak perlu ditambah lagi.');
  } else {
    beriTahu('✅ Kolom berhasil ditambahkan: ' + ditambahkan.join(', '));
  }
}

// Helper: tampilkan alert kalau memang ada UI aktif (misal dipanggil dari
// menu/sidebar). Kalau dijalankan langsung dari tombol Run di Apps Script
// Editor, tidak ada UI aktif sehingga getUi() akan error — di kondisi itu,
// cukup catat pesannya di Execution Log saja (tidak masalah, pekerjaan
// utamanya tetap sudah selesai sebelum baris ini dijalankan).
function beriTahu(pesan) {
  try {
    SpreadsheetApp.getUi().alert(pesan);
  } catch (e) {
    Logger.log(pesan);
  }
}

/**
 * Jalankan fungsi ini SEKALI dari Apps Script Editor untuk menambahkan
 * kolom "Jumlah Transaksi Livin' Food" ke sheet "Laporan Harian" kalau belum ada.
 * Aman dijalankan berkali-kali.
 */
function tambahKolomLivinFood() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET);
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET + '" tidak ditemukan!');
    return;
  }
  var namaKolom = "Jumlah Transaksi Livin' Food";
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf(namaKolom) !== -1) {
    beriTahu('✅ Kolom "' + namaKolom + '" sudah ada, tidak perlu ditambah lagi.');
    return;
  }
  sheet.insertColumnAfter(lastCol);
  sheet.getRange(1, lastCol + 1).setValue(namaKolom);
  sheet.getRange(1, lastCol + 1)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff');
  beriTahu('✅ Kolom "' + namaKolom + '" berhasil ditambahkan.');
}

/**
 * Jalankan fungsi ini SEKALI dari Apps Script Editor untuk menambahkan
 * kolom "Nama Merchant Livin' Food" ke sheet "Laporan Harian" kalau belum ada.
 * Aman dijalankan berkali-kali.
 */
function tambahKolomNamaMerchantLivinFood() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET);
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET + '" tidak ditemukan!');
    return;
  }
  var namaKolom = "Nama Merchant Livin' Food";
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf(namaKolom) !== -1) {
    beriTahu('✅ Kolom "' + namaKolom + '" sudah ada, tidak perlu ditambah lagi.');
    return;
  }
  // Sisipkan setelah kolom "Jumlah Transaksi Livin' Food" jika ada
  var idxLivin = headers.indexOf("Jumlah Transaksi Livin' Food");
  var posisi;
  if (idxLivin !== -1) {
    sheet.insertColumnAfter(idxLivin + 1);
    posisi = idxLivin + 2;
  } else {
    sheet.insertColumnAfter(lastCol);
    posisi = lastCol + 1;
  }
  sheet.getRange(1, posisi).setValue(namaKolom);
  sheet.getRange(1, posisi)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff');
  beriTahu('✅ Kolom "' + namaKolom + '" berhasil ditambahkan.');
}

// ============================================================
// PIPELINE EDC TO LVM — daftar merchant & progress konversi
// ============================================================

/**
 * Ambil / buat sheet "Pipeline EDC to LVM". Kalau baru dibuat, cuma
 * diisi header saja — data merchant (Kode Cabang, Nama Cabang, Area,
 * Nama Merchant) diisi manual oleh admin di Google Sheets.
 */
function getOrCreateSheetPipeline(ss) {
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.NAMA_SHEET_PIPELINE);
    var headers = [
      'Kode Cabang', 'Nama Cabang', 'Area', 'Nama Merchant', 'Alamat', 'Kota',
      'Tag Optimal 24 Juli', 'Status Progress', 'MID', 'Alasan', 'Tanggal Update Terakhir', 'Catatan'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Jalankan SEKALI dari Apps Script Editor untuk menyiapkan sheet
 * "Pipeline EDC to LVM" kalau belum ada. Setelah itu isi manual
 * baris-barisnya: Kode Cabang, Nama Cabang, Area, Nama Merchant
 * (kolom Status Progress & Tanggal Update biar terisi otomatis lewat form).
 */
function setupSheetPipeline() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheetPipeline(ss);
  beriTahu('✅ Sheet "' + CONFIG.NAMA_SHEET_PIPELINE + '" berhasil disiapkan! Silakan isi data merchant-nya secara manual.');
}

/**
 * Jalankan SEKALI dari Apps Script Editor untuk menambahkan kolom "MID"
 * dan "Alasan" ke sheet "Pipeline EDC to LVM" kalau belum ada.
 *
 * Kenapa perlu: sheet ini dibuat sebelum kolom MID & Alasan ditambahkan
 * ke kode (lihat updatePipelineStatus()), jadi keduanya tidak pernah ada
 * di sheet live. Akibatnya walau modal update status di pipeline.html
 * sudah mewajibkan isi "Alasan Menolak" (status Merchant Menolak) atau
 * "MNDI/MID/Nomor Rekening" (status Done Konversi to LVM) dan validasinya
 * lolos, nilainya SELALU gagal tersimpan secara diam-diam — karena
 * updatePipelineStatus() cuma menulis ke kolom itu kalau
 * headers.lastIndexOf('MID')/'Alasan' ketemu (bukan -1). Kolom Status
 * Progress & Tanggal Update Terakhir tetap kesimpan normal, makanya bug
 * ini tidak kelihatan dari situ.
 *
 * Kolom baru disisipkan tepat sebelum "Tanggal Update Terakhir" supaya
 * urutannya sama dengan header di getOrCreateSheetPipeline(). Aman
 * dijalankan berkali-kali — kolom yang sudah ada tidak dibuat dobel.
 */
function tambahKolomMidAlasanPipeline() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET_PIPELINE + '" tidak ditemukan!');
    return;
  }

  var kolomBaru = ['MID', 'Alasan'];
  var ditambahkan = [];

  kolomBaru.forEach(function (nama) {
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers.indexOf(nama) !== -1) return; // sudah ada, lewati

    var idxTgl = headers.indexOf('Tanggal Update Terakhir');
    var posisi;
    if (idxTgl !== -1) {
      sheet.insertColumnBefore(idxTgl + 1);
      posisi = idxTgl + 1;
    } else {
      sheet.insertColumnAfter(lastCol);
      posisi = lastCol + 1;
    }
    sheet.getRange(1, posisi).setValue(nama);
    sheet.getRange(1, posisi)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    ditambahkan.push(nama);
  });

  if (ditambahkan.length === 0) {
    beriTahu('✅ Kolom MID & Alasan sudah ada, tidak perlu ditambah lagi.');
  } else {
    beriTahu('✅ Kolom berhasil ditambahkan: ' + ditambahkan.join(', ') + '. Update status Merchant Menolak / Done Konversi to LVM berikutnya akan otomatis kesimpan ke kolom ini.');
  }
}

// Cari index kolom dari beberapa kemungkinan nama header sekaligus — sheet
// sumber "Pipeline EDC to LVM" sudah beberapa kali ganti nama kolom (mis.
// "Nama Merchant" jadi "dbaname", "Alamat"/"Kota" jadi huruf kecil semua
// "alamat"/"kota", "Tag Optimal 24 Juli" jadi "Tag Utilized 31 Juli", "MID"
// sumber jadi "MID_NEW") setiap kali data master di-refresh dari sumbernya.
// Dicoba semua nama yang mungkin (urut prioritas), dipakai yang PERTAMA
// ketemu — jadi kalau nama kolomnya ganti lagi di masa depan, tinggal
// tambah nama barunya di daftar kandidat, tidak perlu ubah logic lain.
// PENTING: ini penyebab bug "semua angka jadi 0" yang pernah kejadian —
// waktu itu headers.indexOf('Nama Merchant') balikin -1 karena kolomnya
// sudah berganti nama jadi 'dbaname', jadi seluruh blok pembacaan data
// di-skip diam-diam (idxMerchant !== -1 gagal, tanpa error apapun).
function _cariIndexHeader(headers, kandidatNama) {
  for (var i = 0; i < kandidatNama.length; i++) {
    var idx = headers.indexOf(kandidatNama[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * Ambil daftar merchant EDC (pipeline konversi ke LVM) milik satu cabang,
 * lengkap dengan status progress terkininya. Dipanggil dari form (index.html)
 * lewat action=getMerchantEDC saat cabang dipilih.
 */
function getMerchantEDCByCabang(kodeCabang) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers      = data[0];
  var idxKode      = headers.indexOf('Kode Cabang');
  var idxMerchant  = _cariIndexHeader(headers, ['Nama Merchant', 'dbaname']);
  var idxStatus    = headers.indexOf('Status Progress');
  var idxTgl       = headers.indexOf('Tanggal Update Terakhir');

  var kodeCari = String(kodeCabang || '').trim();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    var kodeBaris = String(data[i][idxKode] || '').trim();
    var namaMerchant = String(data[i][idxMerchant] || '').trim();
    if (kodeBaris === kodeCari && namaMerchant) {
      hasil.push({
        namaMerchant  : namaMerchant,
        status        : data[i][idxStatus] || 'Target',
        tanggalUpdate : data[i][idxTgl] ? formatTanggal(data[i][idxTgl]) : '-'
      });
    }
  }
  return hasil;
}

/**
 * Ambil SEMUA data pipeline EDC to LVM (lintas cabang), opsional
 * difilter per kodeCabang. Dipanggil dari pipeline.html lewat
 * action=getPipelineAll — kalau parameter kodeCabang dikirim,
 * cuma baris cabang itu yang dikembalikan (dipakai buat link
 * per-cabang, misal pipeline.html?kode=15200).
 */
// Data mentah "Tag Optimal 24 Juli" kadang isinya "#N/A" (excel formula
// error dari sumber data asal) atau kosong buat merchant yang belum
// sempat di-tag — keduanya dianggap "tidak ada tag" ('-'), biar badge
// tag di pipeline.html tidak nampilin teks "#N/A" mentah ke user.
function bersihkanTagOptimal(v) {
  var s = String(v || '').trim();
  if (!s || s === '#N/A') return '-';
  return s;
}

function getPipelineAll(kodeCabangFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers      = data[0];
  var idxKode      = headers.indexOf('Kode Cabang');
  var idxNama      = headers.indexOf('Nama Cabang');
  var idxArea      = headers.indexOf('Area');
  var idxMerchant  = _cariIndexHeader(headers, ['Nama Merchant', 'dbaname']);
  var idxAlamat    = _cariIndexHeader(headers, ['Alamat', 'alamat']);
  var idxKota      = _cariIndexHeader(headers, ['Kota', 'kota']);
  var idxStatus    = headers.indexOf('Status Progress');
  // Sheet hasil gabungan data master punya 2 kolom "MID": yang PERTAMA
  // (dekat kolom awal, sekarang bernama "MID_NEW") itu MID/ID merchant asli
  // dari data core banking, BUKAN yang kita mau. Yang kita catat lewat
  // pipeline.html (MID/rekening yang diisi user pas tandain "Done Konversi
  // to LVM") itu kolom "MID" yang TERAKHIR, persis sebelah "Status
  // Progress"/"Alasan". Makanya pakai lastIndexOf, bukan indexOf.
  var idxMid       = headers.lastIndexOf('MID');
  // Kolom "MID" PERTAMA/sumber (index berbeda dari idxMid di atas) — ini ID
  // unik merchant/EDC dari core banking, hampir selalu unik per baris (beda
  // dengan Nama Merchant/dbaname yang bisa dobel kalau merchant-nya
  // franchise/nama umum). Dipakai sebagai pembeda paling akurat kalau ada
  // 2+ merchant dengan nama sama persis dalam 1 cabang. Nama kolomnya sudah
  // berganti dari "MID" jadi "MID_NEW" di sheet sumber — dicoba "MID_NEW"
  // dulu baru fallback ke "MID" (buat sheet lama yang belum ganti nama).
  var idxMidSumber = _cariIndexHeader(headers, ['MID_NEW', 'MID']);
  var adaDuaKolomMid = idxMidSumber !== -1 && idxMidSumber !== idxMid;
  var idxAlasan    = headers.indexOf('Alasan');
  // Kolom "Tagging Merchant" (lama) sudah tidak ada di database yang sudah
  // diperbaiki — sempat diganti "Tag Optimal 24 Juli", lalu berganti lagi
  // jadi "Tag Utilized 31 Juli" — isinya persis sama (Optimal / Hampir
  // Optimal / Kurang Optimal / Nihil) jadi badge tag di pipeline.html tetap
  // jalan tanpa perlu ubah tampilan.
  var idxTagging   = _cariIndexHeader(headers, ['Tag Optimal 24 Juli', 'Tag Utilized 31 Juli']);
  var idxTgl       = headers.indexOf('Tanggal Update Terakhir');
  var idxCatatan   = headers.indexOf('Catatan');

  var kodeCari = String(kodeCabangFilter || '').trim();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    var kodeBaris     = String(data[i][idxKode] || '').trim();
    var namaMerchant  = String(data[i][idxMerchant] || '').trim();
    if (!kodeBaris || !namaMerchant) continue;
    if (kodeCari && kodeBaris !== kodeCari) continue;

    hasil.push({
      kodeCabang    : kodeBaris,
      namaCabang    : data[i][idxNama] || '-',
      area          : data[i][idxArea] || '-',
      namaMerchant  : namaMerchant,
      alamat        : idxAlamat !== -1 ? (data[i][idxAlamat] || '-') : '-',
      kota          : idxKota   !== -1 ? (data[i][idxKota]   || '-') : '-',
      status        : data[i][idxStatus] || 'Target',
      mid           : idxMid !== -1 ? (data[i][idxMid] || '-') : '-',
      midMerchant   : adaDuaKolomMid ? (data[i][idxMidSumber] || '-') : '-',
      alasan        : idxAlasan !== -1 ? (data[i][idxAlasan] || '-') : '-',
      tagging       : idxTagging !== -1 ? bersihkanTagOptimal(data[i][idxTagging]) : '-',
      tanggalUpdate : data[i][idxTgl] ? formatTanggal(data[i][idxTgl]) : '-',
      catatan       : idxCatatan !== -1 ? (data[i][idxCatatan] || '-') : '-'
    });
  }
  return hasil;
}

/**
 * Ambil semua merchant di "Pipeline EDC to LVM" milik satu cabang yang
 * "Tanggal Update Terakhir"-nya SAMA dengan tanggal laporan yang sedang
 * disubmit. Ini yang bikin teks WA laporan harian otomatis nampilin
 * SEMUA merchant yang statusnya diupdate hari itu (lewat pipeline.html),
 * bukan cuma 1 merchant — jadi kalau cabang kunjungan 5 merchant dalam
 * sehari dan update ke-5nya di pipeline.html, ke-5nya otomatis muncul
 * di teks WA laporan harian tanpa perlu pilih manual di form.
 */
function getPipelineUpdatesHariIni(kodeCabang, tanggalLaporan) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  var headers     = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxKode     = headers.indexOf('Kode Cabang');
  var idxMerchant = _cariIndexHeader(headers, ['Nama Merchant', 'dbaname']);
  var idxStatus   = headers.indexOf('Status Progress');
  var idxMid      = headers.lastIndexOf('MID');
  var idxAlasan   = headers.indexOf('Alasan');
  var idxTgl      = headers.indexOf('Tanggal Update Terakhir');
  if (idxKode === -1 || idxMerchant === -1 || idxTgl === -1) return [];

  var kodeCari = String(kodeCabang || '').trim();
  var tglCari  = formatTanggal(tanggalLaporan); // dd/MM/yyyy, sama format dgn Tanggal Update Terakhir

  // Saring dulu pakai kolom Kode Cabang & Tanggal Update Terakhir saja
  // (ringan buat ribuan baris), baru ambil kolom lain (Nama Merchant,
  // Status, MID, Alasan) khusus buat baris yang cocok saja — bukan baca
  // semua kolom dari semua baris tiap kali submit laporan/buka laporan.
  var kolomKode = sheet.getRange(2, idxKode + 1, lastRow - 1, 1).getValues();
  var kolomTgl  = sheet.getRange(2, idxTgl + 1, lastRow - 1, 1).getValues();

  var barisCocok = [];
  for (var i = 0; i < kolomKode.length; i++) {
    var kodeBaris = String(kolomKode[i][0] || '').trim();
    if (kodeBaris !== kodeCari) continue;
    var tglBaris = kolomTgl[i][0] ? formatTanggal(kolomTgl[i][0]) : '';
    if (!tglBaris || tglBaris !== tglCari) continue;
    barisCocok.push(i + 2); // nomor baris asli di sheet (lewati header)
  }

  return barisCocok.map(function (row) {
    var nilai = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
    return {
      namaMerchant : nilai[idxMerchant] || '-',
      status       : nilai[idxStatus] || 'Target',
      mid          : idxMid !== -1 ? (nilai[idxMid] || '-') : '-',
      alasan       : idxAlasan !== -1 ? (nilai[idxAlasan] || '-') : '-'
    };
  });
}

/**
 * Update status progress satu merchant di sheet "Pipeline EDC to LVM".
 * Dipanggil dari action=updateStatus (pipeline.html). Mencari baris yang
 * cocok, lalu menimpa kolom Status Progress, MID (kalau diisi), Alasan
 * (kalau diisi), & Tanggal Update Terakhir baris itu (bukan menambah
 * baris baru — 1 merchant = 1 baris terus terupdate).
 *
 * Prioritas pencocokan baris (dari paling akurat ke paling longgar):
 *   1. Kode Cabang + Nama Merchant + MID Merchant (ID unik dari core
 *      banking, kolom "MID" PERTAMA di sheet). Ini yang paling diandalkan
 *      — hampir 100% unik per baris walau nama merchant sama persis.
 *   2. Kode Cabang + Nama Merchant + Alamat, kalau midMerchant tidak
 *      dikirim/tidak tersedia di sheet.
 *   3. Kode Cabang + Nama Merchant saja, kalau alamat juga tidak ada.
 * Alasan butuh pembeda sama sekali: database gabungan ternyata punya
 * banyak merchant BEDA dalam 1 cabang yang kebetulan namanya SAMA PERSIS
 * (mis. franchise "Restoran Sederhana" di beberapa lokasi) — kalau cuma
 * cocokkan Kode Cabang + Nama Merchant, update bisa nimpa baris yang salah.
 *
 * Parameter "mid" & "alasan" opsional — hanya ditulis kalau ada isinya.
 * Return true kalau baris ditemukan & berhasil diupdate, false kalau tidak.
 */
function updatePipelineStatus(ss, kodeCabang, namaMerchant, statusBaru, mid, alasan, alamat, midMerchant) {
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) return false;

  if (STATUS_EDC_LVM_VALID.indexOf(statusBaru) === -1) {
    Logger.log('Status progress EDC to LVM tidak valid: ' + statusBaru);
    return false;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return false;

  var headers      = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxKode      = headers.indexOf('Kode Cabang');
  var idxMerchant  = _cariIndexHeader(headers, ['Nama Merchant', 'dbaname']);
  var idxAlamat    = _cariIndexHeader(headers, ['Alamat', 'alamat']);
  var idxStatus    = headers.indexOf('Status Progress');
  var idxMid       = headers.lastIndexOf('MID');
  // Kolom "MID" PERTAMA/sumber (beda dari idxMid di atas) — ID unik
  // merchant/EDC dari core banking, jadi pembeda paling akurat. Nama
  // kolomnya sudah berganti dari "MID" jadi "MID_NEW" di sheet sumber.
  // Lihat komentar di getPipelineAll().
  var idxMidSumber   = _cariIndexHeader(headers, ['MID_NEW', 'MID']);
  var adaDuaKolomMid = idxMidSumber !== -1 && idxMidSumber !== idxMid;
  var idxAlasan    = headers.indexOf('Alasan');
  var idxTgl       = headers.indexOf('Tanggal Update Terakhir');
  if (idxKode === -1 || idxMerchant === -1 || idxStatus === -1) return false;

  // Cuma baca kolom-kolom yang dibutuhkan buat nyari baris yang cocok —
  // bukan getDataRange().getValues() yang narik SEMUA kolom (termasuk
  // Catatan yang teksnya panjang) untuk ribuan baris. Sheet Pipeline
  // sudah >4000 baris; baca semua kolom tiap kali simpan status bikin
  // request lambat dan gampang timeout/"Failed to fetch" di jaringan yang
  // kurang stabil.
  var kolomKode     = sheet.getRange(2, idxKode + 1, lastRow - 1, 1).getValues();
  var kolomMerchant = sheet.getRange(2, idxMerchant + 1, lastRow - 1, 1).getValues();
  var kolomAlamat   = idxAlamat !== -1
    ? sheet.getRange(2, idxAlamat + 1, lastRow - 1, 1).getValues()
    : null;
  var kolomMidSumber = adaDuaKolomMid
    ? sheet.getRange(2, idxMidSumber + 1, lastRow - 1, 1).getValues()
    : null;

  var kodeCari       = String(kodeCabang || '').trim();
  var merchantCari   = String(namaMerchant || '').trim();
  var alamatCari     = String(alamat || '').trim();
  var midMerchantCari = String(midMerchant || '').trim();

  // Pembeda dipakai sesuai urutan prioritas — makin ke bawah makin
  // longgar, cuma dipakai kalau yang di atasnya tidak tersedia.
  var pakaiMidMerchant = adaDuaKolomMid && midMerchantCari !== '';
  var pakaiAlamat      = !pakaiMidMerchant && idxAlamat !== -1 && alamatCari !== '';

  for (var i = 0; i < kolomKode.length; i++) {
    var kodeBaris     = String(kolomKode[i][0] || '').trim();
    var merchantBaris = String(kolomMerchant[i][0] || '').trim();
    if (kodeBaris !== kodeCari || merchantBaris !== merchantCari) continue;
    if (pakaiMidMerchant) {
      var midSumberBaris = String(kolomMidSumber[i][0] || '').trim();
      if (midSumberBaris !== midMerchantCari) continue;
    } else if (pakaiAlamat) {
      var alamatBaris = String(kolomAlamat[i][0] || '').trim();
      if (alamatBaris !== alamatCari) continue;
    }

    var row = i + 2; // +2: lewati baris header, kolom-kolom di atas 0-based mulai dari baris 2
    sheet.getRange(row, idxStatus + 1).setValue(statusBaru);
    if (idxMid !== -1 && mid && String(mid).trim()) {
      sheet.getRange(row, idxMid + 1).setValue(String(mid).trim());
    }
    if (idxAlasan !== -1 && alasan && String(alasan).trim()) {
      sheet.getRange(row, idxAlasan + 1).setValue(String(alasan).trim());
    }
    if (idxTgl !== -1) {
      sheet.getRange(row, idxTgl + 1).setValue(
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy')
      );
    }
    return true;
  }
  return false;
}

/**
 * Jalankan SEKALI dari Apps Script Editor setelah update status pipeline
 * (skema lama → Target/Done Konversi to LVM/Merchant Menolak/Merchant
 * Tutup/Perlu Kunjungan MTI), supaya baris-baris lama di sheet "Pipeline
 * EDC to LVM" yang masih pakai istilah status sebelumnya ("Proses",
 * "Sudah LVM", "DONE KONVERSI TO LVM" huruf kapital semua, "Kunjungan
 * Awal", dst.) ikut termigrasi ke istilah baru.
 * Aman dijalankan berkali-kali — baris yang sudah pakai status baru dilewati.
 */
function migrasiStatusPipeline() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET_PIPELINE + '" tidak ditemukan!');
    return;
  }

  var PETA_MIGRASI = {
    'Proses'               : 'Target',
    'Kunjungan Awal'       : 'Target',
    'Sudah LVM'            : 'Done Konversi to LVM',
    'Done Akuisisi'        : 'Done Konversi to LVM',
    'DONE KONVERSI TO LVM' : 'Done Konversi to LVM'
  };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET_PIPELINE + '" belum ada data.');
    return;
  }

  var headers   = data[0];
  var idxStatus = headers.indexOf('Status Progress');
  if (idxStatus === -1) {
    beriTahu('Kolom "Status Progress" tidak ditemukan di sheet.');
    return;
  }

  var jumlahDiubah = 0;
  for (var i = 1; i < data.length; i++) {
    var statusLama = String(data[i][idxStatus] || '').trim();
    if (PETA_MIGRASI.hasOwnProperty(statusLama)) {
      sheet.getRange(i + 1, idxStatus + 1).setValue(PETA_MIGRASI[statusLama]);
      jumlahDiubah++;
    }
  }

  beriTahu(jumlahDiubah > 0
    ? '✅ ' + jumlahDiubah + ' baris berhasil dimigrasikan ke status baru.'
    : '✅ Tidak ada baris dengan status lama — semua sudah pakai istilah baru.');
}

/**
 * Jalankan SEKALI dari Apps Script Editor untuk mengosongkan kolom
 * "Tanggal Update Terakhir" pada baris-baris sheet "Pipeline EDC to LVM"
 * yang tanggalnya SAMA dengan TANGGAL_UPLOAD di bawah (28/07/2026).
 *
 * Kenapa perlu: data konversi lama yang di-upload manual ke sheet ikut
 * kebawa kolom "Tanggal Update Terakhir" = tanggal upload, padahal
 * merchant-merchant itu sebenarnya sudah dikonversi jauh sebelum tanggal
 * itu. Akibatnya getPipelineUpdatesHariIni() salah mengira baris itu baru
 * diupdate hari itu, jadi ikut otomatis ditarik ke wording WA laporan
 * harian cabang terkait — padahal user yang submit laporan tidak
 * input/update progress apa pun.
 *
 * Baris yang MEMANG diupdate via pipeline.html pada TANGGAL_UPLOAD (bukan
 * bagian dari data upload lama) ikut kena kosongkan juga — kalau ada,
 * cukup buka pipeline.html & update ulang status merchant tsb supaya
 * tanggalnya tercatat benar lagi.
 *
 * Aman dijalankan lebih dari sekali — baris yang sudah kosong dilewati.
 */
function bersihkanTanggalUploadLama() {
  var TANGGAL_UPLOAD = '28/07/2026'; // ganti tanggal ini kalau perlu bersihkan tanggal lain

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET_PIPELINE + '" tidak ditemukan!');
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET_PIPELINE + '" belum ada data.');
    return;
  }

  var headers = data[0];
  var idxTgl = headers.indexOf('Tanggal Update Terakhir');
  if (idxTgl === -1) {
    beriTahu('Kolom "Tanggal Update Terakhir" tidak ditemukan di sheet.');
    return;
  }

  var jumlahDibersihkan = 0;
  for (var i = 1; i < data.length; i++) {
    var nilaiTgl = data[i][idxTgl];
    if (!nilaiTgl) continue; // sudah kosong, lewati
    if (formatTanggal(nilaiTgl) === TANGGAL_UPLOAD) {
      sheet.getRange(i + 1, idxTgl + 1).setValue('');
      jumlahDibersihkan++;
    }
  }

  beriTahu(jumlahDibersihkan > 0
    ? '✅ ' + jumlahDibersihkan + ' baris dengan tanggal ' + TANGGAL_UPLOAD + ' berhasil dikosongkan.'
    : '✅ Tidak ada baris dengan tanggal ' + TANGGAL_UPLOAD + ' — mungkin sudah dibersihkan sebelumnya, atau formatnya beda di sheet.');
}

/**
 * Jalankan SEKALI dari Apps Script Editor untuk menambahkan kolom
 * "Merchant EDC to LVM" dan "Status Progress EDC to LVM" ke sheet
 * "Laporan Harian" kalau belum ada. Aman dijalankan berkali-kali.
 */
function tambahKolomEDCLVM() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET);
  if (!sheet) {
    beriTahu('Sheet "' + CONFIG.NAMA_SHEET + '" tidak ditemukan!');
    return;
  }
  var kolomBaru = ['Merchant EDC to LVM', 'Status Progress EDC to LVM'];
  var ditambahkan = [];
  kolomBaru.forEach(function (nama) {
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers.indexOf(nama) !== -1) return;
    sheet.insertColumnAfter(lastCol);
    sheet.getRange(1, lastCol + 1).setValue(nama);
    sheet.getRange(1, lastCol + 1)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    ditambahkan.push(nama);
  });
  if (ditambahkan.length === 0) {
    beriTahu('✅ Kolom EDC to LVM sudah ada, tidak perlu ditambah lagi.');
  } else {
    beriTahu('✅ Kolom berhasil ditambahkan: ' + ditambahkan.join(', '));
  }
}

// ============================================================
// PIPELINE LEAKAGE TOP 100 — dipakai monitoring-top100-leakage.html.
// Data awal (722 merchant, 7 Area) diimpor manual dari file Excel
// "Pipeline Leakage Top 100" — lihat setupSheetLeakage() di bawah.
// Kolom "Cabang" (Kode Cabang) ditambahkan belakangan oleh admin
// langsung di sheet — dibaca by-name lewat headers.indexOf() di
// getLeakageAll(), jadi aman berapa pun posisi kolomnya di sheet.
// ============================================================

// Header kolom sheet "Pipeline Leakage Top 100" — dipakai getLeakageAll()
// buat tau kolom APA SAJA yang perlu dicari (posisinya dicari dinamis by
// nama, bukan by urutan, jadi aman kalau admin nyisip kolom baru di sheet
// selama namanya sama persis dengan salah satu string di bawah ini).
var HEADER_LEAKAGE = [
  'Area', 'Cabang', 'Merchant', 'No CIF', 'No Rek Set',
  'Sales Volume (Rp Jt)', 'Incoming Rekening Settlement (Rp Jt)', 'SV/Incoming Rek Settlement',
  'Leakage ke Rek Sendiri Non BMRI (Rp Jt)', 'Leakage ke Rek Pihak Ketiga Non BMRI (Rp Jt)',
  'Total Leakage (Rp Jt)', 'Leakage Ratio',
  'Keterangan Leakage', 'Potensi Winback (Rp Jt)', 'Penawaran Prima/Prima Xtra (Y/N)',
  'Penawaran Lainnya', 'Keterangan Follow Up', 'Pipeline', 'Ket Konfirmasi',
  'Tanggal Update Terakhir'
];

// 7 kolom "hasil follow up" yang diisi cabang lewat modal di
// monitoring-top100-leakage.html — dipakai bareng oleh getLeakageAll()
// (baca) & updateLeakageFollowUp() (tulis) supaya urutan field frontend
// <-> backend selalu sinkron di SATU tempat ini saja.
var FIELD_FOLLOWUP_LEAKAGE = [
  { key: 'keteranganLeakage', header: 'Keterangan Leakage',                 valid: KETERANGAN_LEAKAGE_VALID },
  { key: 'potensiWinback',    header: 'Potensi Winback (Rp Jt)',            valid: null },
  { key: 'penawaranPrima',    header: 'Penawaran Prima/Prima Xtra (Y/N)',   valid: PENAWARAN_PRIMA_VALID },
  { key: 'penawaranLainnya',  header: 'Penawaran Lainnya',                  valid: null },
  { key: 'keteranganFollowUp',header: 'Keterangan Follow Up',               valid: null },
  { key: 'sumberPipeline',    header: 'Pipeline',                          valid: SUMBER_PIPELINE_VALID },
  { key: 'ketKonfirmasi',     header: 'Ket Konfirmasi',                     valid: KET_KONFIRMASI_VALID }
];

function getOrCreateSheetLeakage(ss) {
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_LEAKAGE);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.NAMA_SHEET_LEAKAGE);
    sheet.getRange(1, 1, 1, HEADER_LEAKAGE.length).setValues([HEADER_LEAKAGE]);
    sheet.getRange(1, 1, 1, HEADER_LEAKAGE.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Jalankan SEKALI dari Apps Script Editor buat menyiapkan sheet "Pipeline
 * Leakage Top 100" (header saja). Setelah itu import file
 * "Pipeline Leakage Top 100 - Import.csv" sebagai SHEET BARU lewat menu
 * Google Sheets File > Import > Upload > "Replace current sheet" (pilih
 * sheet "Pipeline Leakage Top 100" yang baru dibuat function ini), supaya
 * 712 baris datanya masuk dengan header yang PERSIS sama seperti yang
 * dibaca kode ini.
 */
function setupSheetLeakage() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheetLeakage(ss);
  beriTahu('✅ Sheet "' + CONFIG.NAMA_SHEET_LEAKAGE + '" berhasil disiapkan! Silakan import data dari file CSV yang sudah disiapkan (lihat komentar di atas function ini).');
}

// "150 - MANADO" -> "150". Dipakai buat cocokin parameter ?area= dari
// index.html (diambil dari 3 digit pertama Kode Cabang) dan buat filter
// di getLeakageAll().
function kodeAreaDariString(area) {
  var s = String(area || '');
  var i = s.indexOf(' - ');
  return i === -1 ? s.trim() : s.substring(0, i).trim();
}

// "150 - MANADO" -> "MANADO". Buat label ringkas di UI/badge.
function labelAreaLeakage(area) {
  var s = String(area || '');
  var i = s.indexOf(' - ');
  return i === -1 ? s.trim() : s.substring(i + 3).trim();
}

/**
 * Ambil semua data Pipeline Leakage Top 100, opsional difilter per 3
 * digit kode area (mis. "150" utk Manado). Dipanggil dari
 * monitoring-top100-leakage.html lewat action=getLeakageAll.
 */
function getLeakageAll(kodeAreaFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_LEAKAGE);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0];
  var idx = {};
  HEADER_LEAKAGE.forEach(function (h) { idx[h] = headers.indexOf(h); });

  var kodeCari = String(kodeAreaFilter || '').trim();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var merchant = String(row[idx['Merchant']] || '').trim();
    var area = String(row[idx['Area']] || '').trim();
    if (!merchant || !area) continue;
    if (kodeCari && kodeAreaDariString(area) !== kodeCari) continue;

    var item = {
      area              : area,
      kodeArea          : kodeAreaDariString(area),
      labelArea         : labelAreaLeakage(area),
      kodeCabang        : String(row[idx['Cabang']] || '').trim() || '-',
      merchant          : merchant,
      noCif             : row[idx['No CIF']] || '-',
      noRekSet          : row[idx['No Rek Set']] || '-',
      salesVolume       : Number(row[idx['Sales Volume (Rp Jt)']]) || 0,
      incomingRek       : Number(row[idx['Incoming Rekening Settlement (Rp Jt)']]) || 0,
      svIncomingRatio   : Number(row[idx['SV/Incoming Rek Settlement']]) || 0,
      leakageRekSendiri : Number(row[idx['Leakage ke Rek Sendiri Non BMRI (Rp Jt)']]) || 0,
      leakageRekPihak3  : Number(row[idx['Leakage ke Rek Pihak Ketiga Non BMRI (Rp Jt)']]) || 0,
      totalLeakage      : Number(row[idx['Total Leakage (Rp Jt)']]) || 0,
      leakageRatio      : Number(row[idx['Leakage Ratio']]) || 0,
      tanggalUpdate     : row[idx['Tanggal Update Terakhir']] ? formatTanggal(row[idx['Tanggal Update Terakhir']]) : '-'
    };
    FIELD_FOLLOWUP_LEAKAGE.forEach(function (f) {
      item[f.key] = row[idx[f.header]] || '-';
    });
    hasil.push(item);
  }
  return hasil;
}

/**
 * Simpan hasil follow up satu merchant leakage. Dicari berdasarkan
 * kombinasi Area + Merchant (unik — sudah dicek waktu nyiapin data
 * awal, 1 area tidak punya 2 merchant dengan nama sama persis).
 * Menimpa baris yang ketemu (bukan nambah baris baru), sama seperti
 * updatePipelineStatus() di Pipeline EDC to LVM.
 */
function updateLeakageFollowUp(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_LEAKAGE);
  if (!sheet) return { error: 'Sheet "' + CONFIG.NAMA_SHEET_LEAKAGE + '" tidak ditemukan.' };

  var areaCari     = String(p.area || '').trim();
  var merchantCari = String(p.merchant || '').trim();
  if (!areaCari || !merchantCari) return { error: 'Area/Merchant tidak lengkap.' };

  // Validasi field dropdown — field bebas teks (valid: null) dilewati.
  for (var f = 0; f < FIELD_FOLLOWUP_LEAKAGE.length; f++) {
    var field = FIELD_FOLLOWUP_LEAKAGE[f];
    var nilai = String(p[field.key] || '').trim();
    if (field.valid && nilai && field.valid.indexOf(nilai) === -1) {
      return { error: 'Nilai "' + nilai + '" tidak valid untuk ' + field.header + '.' };
    }
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { error: 'Data Pipeline Leakage Top 100 masih kosong.' };

  var headers = data[0];
  var idxArea     = headers.indexOf('Area');
  var idxMerchant = headers.indexOf('Merchant');
  var idxTgl      = headers.indexOf('Tanggal Update Terakhir');
  if (idxArea === -1 || idxMerchant === -1) return { error: 'Header sheet tidak sesuai.' };

  for (var i = 1; i < data.length; i++) {
    var areaBaris     = String(data[i][idxArea] || '').trim();
    var merchantBaris = String(data[i][idxMerchant] || '').trim();
    if (areaBaris === areaCari && merchantBaris === merchantCari) {
      FIELD_FOLLOWUP_LEAKAGE.forEach(function (field) {
        var idxKolom = headers.indexOf(field.header);
        if (idxKolom === -1) return;
        var nilai = p[field.key];
        if (nilai !== undefined) {
          sheet.getRange(i + 1, idxKolom + 1).setValue(nilai);
        }
      });
      if (idxTgl !== -1) {
        sheet.getRange(i + 1, idxTgl + 1).setValue(
          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy')
        );
      }
      return {};
    }
  }
  return { error: 'Merchant "' + merchantCari + '" tidak ditemukan di Area "' + areaCari + '".' };
}

/**
 * Bikin/refresh SATU spreadsheet snapshot terpisah ("Export - Pipeline
 * Leakage Top 100") berisi cuma data sheet ini apa adanya (bukan seluruh
 * file utama yang punya banyak tab lain) — supaya tombol "Download Excel"
 * di monitoring-top100-leakage.html bisa langsung ngasih link export
 * .xlsx yang bersih. ID spreadsheet snapshot disimpan di Script
 * Properties biar dipakai ULANG (bukan bikin file baru terus tiap
 * diklik) — kalau file lamanya kehapus/hilang, otomatis bikin baru lagi.
 */
function exportLeakageExcelFile() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_LEAKAGE);
  if (!sheet) throw new Error('Sheet "' + CONFIG.NAMA_SHEET_LEAKAGE + '" tidak ditemukan.');

  var data = sheet.getDataRange().getValues();
  var props = PropertiesService.getScriptProperties();
  var idTersimpan = props.getProperty('LEAKAGE_EXPORT_SPREADSHEET_ID');

  var exportSs = null;
  if (idTersimpan) {
    try { exportSs = SpreadsheetApp.openById(idTersimpan); } catch (e) { exportSs = null; }
  }
  if (!exportSs) {
    exportSs = SpreadsheetApp.create('Export - Pipeline Leakage Top 100');
    props.setProperty('LEAKAGE_EXPORT_SPREADSHEET_ID', exportSs.getId());
    try {
      DriveApp.getFileById(exportSs.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) { /* kalau Drive API dibatasi domain, link tetap dibalikin — user tinggal login sendiri buat akses */ }
  }

  var exportSheet = exportSs.getSheets()[0];
  exportSheet.clear();
  if (data.length > 0) {
    exportSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    exportSheet.getRange(1, 1, 1, data[0].length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    exportSheet.setFrozenRows(1);
  }
  exportSheet.setName('Pipeline Leakage Top 100');

  return 'https://docs.google.com/spreadsheets/d/' + exportSs.getId() + '/export?format=xlsx';
}

// ============================================================
// RINGKASAN AREA — dashboard "Pipeline Akuisisi Merchant" per Area
// (dipakai monitoring-area.html), termasuk snapshot harian buat DtD.
// ============================================================

var CONFIG_RINGKASAN = {
  NAMA_SHEET_STATIS   : 'Data Statis Cabang',      // "Shifting to Optimal" diisi manual per Cabang
  NAMA_SHEET_RIWAYAT  : 'Riwayat Progress Cabang'  // snapshot Jml FU harian per Cabang, dasar hitung DtD
};

// Kategori "Hasil Follow Up" yang DIHITUNG OTOMATIS dari kolom Status
// Progress di sheet Pipeline (selain "Target", yang artinya belum
// di-follow-up sama sekali — makanya tidak masuk daftar ini, dia yang
// jadi dasar hitung Sisa Leads).
var KATEGORI_HASIL_FU_OTOMATIS = [
  'Done Konversi to LVM',
  'Tetap pakai EDC & Optimalkan SV',
  'Penawaran',
  'Merchant Menolak',
  'Merchant Tutup',
  'Perlu Kunjungan MTI'
];

// Peta Kode Cabang → Area RESMI (7 area baku Region X), sama persis dengan
// DATA_CABANG yang dipakai index.html & pipeline.html — supaya grouping per
// Area konsisten, bukan baca kolom "Area" di sheet Pipeline yang kadang
// keisi manual/data kotor. Dibangun sekali & di-cache di properti global.
var _DATA_CABANG_AREA_CACHE = null;
function getDataCabangArea() {
  if (_DATA_CABANG_AREA_CACHE) return _DATA_CABANG_AREA_CACHE;
  _DATA_CABANG_AREA_CACHE = {
    "Area Manado": ["15000","15001","15002","15003","15004","15005","15006","15007","15008","15010","15011","15012","15015","15017","15018","15019","15020","15021","15022","15075","15076","15078","15079","15081","15082","15083","15085","15086","15088","15089","15090","15091","15096","15097"],
    "Area Palu": ["15100","15101","15102","15103","15104","15105","15106","15107","15109","15110","15111","15112","15175","15177","15178","15179","15180","15181","15182","15184","15185","15186","15187","15188","15189","15190","15191"],
    "Area Makassar Kartini": ["15200","15201","15202","15203","15213","15216","15220","15221","15225","15237","15238","15242","15243","15244","15245","15246","15248","15249","15250","15251","15252","15269","15271","15291","15293","15297"],
    "Area Kendari": ["16200","16201","16202","16203","16204","16205","16206","16207","16208","16210","16275","16276","16277","16278","16279","16280","16281","16282","16283","16284","16285"],
    "Area Pare Pare": ["17000","17001","17002","17003","17004","17005","17006","17007","17008","17009","17010","17011","17015","17016","17075","17076","17077","17078","17079","17080","17082","17083","17084","17085","17087","17088","17091"],
    "Area Makassar Ratulangi": ["17400","17401","17402","17403","17404","17406","17407","17408","17409","17410","17411","17451","17452","17453","17454","17455","17456","17457","17458","17459","17462","17463"],
    "Area Maluku": ["18600","18601","18602","18603","18604","18605","18607","18608","18609","18610","18611","18612","18613","18651","18652","18654","18655","18656","18657","18658","18659","18660","18661"]
  };
  return _DATA_CABANG_AREA_CACHE;
}

var _KODE_TO_AREA_CACHE = null;
function getKodeToAreaMap() {
  if (_KODE_TO_AREA_CACHE) return _KODE_TO_AREA_CACHE;
  var dataArea = getDataCabangArea();
  var map = {};
  Object.keys(dataArea).forEach(function (area) {
    dataArea[area].forEach(function (kode) { map[kode] = area; });
  });
  _KODE_TO_AREA_CACHE = map;
  return map;
}

// Peta Kode Cabang → Nama Cabang, buat tabel per-Cabang di monitoring-area.html.
// Sama persis dengan daftar 180 cabang Region X (nama & kode resmi).
var _KODE_TO_NAMA_CACHE = null;
function getKodeToNamaMap() {
  if (_KODE_TO_NAMA_CACHE) return _KODE_TO_NAMA_CACHE;
  _KODE_TO_NAMA_CACHE = {
    "15000":"KC Manado Dotulolong Lasut","15001":"KCP Manado Toar","15002":"KC Kotamobagu","15003":"KC Bitung","15004":"KC Gorontalo","15005":"KC Tahuna","15006":"KCP Manado Sam Ratulangi","15007":"KCP Manado Sudirman","15008":"KCP Limboto","15010":"KCP Manado Boulevard","15011":"KCP Marisa","15012":"KCP Tomohon","15015":"KCP Amurang","15017":"KCP Manado Bahu","15018":"KCP Manado Pasar Calaca","15019":"KCP Girian","15020":"KCP Manado Paal Dua","15021":"KCP Airmadidi","15022":"KCP Manado Town Square 3","15075":"KCP Manado Airmadidi","15076":"KCP Manado Ratahan","15078":"KCP Gorontalo Kwandang","15079":"KCP Paguyaman","15081":"KCP Manado Paniki","15082":"KCP Manado Langowan","15083":"KCP Gorontalo Agus Salim","15085":"KCP Tilamuta","15086":"KCP Tondano","15088":"KCP Gorontalo Isimu","15089":"KCP Manado Ringroad","15090":"KCP Bone Bolango Kabila","15091":"KCP Siau","15096":"KCP Boalemo Wonosari","15097":"KCP Lolak",
    "15100":"KC Palu Sam Ratulangi","15101":"KCP Palu Imam Bonjol","15102":"KCP Palu Hasanuddin","15103":"KCP Donggala","15104":"KC Luwuk","15105":"KCP Poso","15106":"KC Toli Toli","15107":"KCP Parigi","15109":"KCP Ampana","15110":"KCP Palu Basuki Rahmat","15111":"KCP Morowali Bahodopi","15112":"KCP Tawaeli","15175":"KCP Pasangkayu","15177":"KCP Toili","15178":"KCP Palu Moh Yamin","15179":"KCP Palu Dewi Sartika","15180":"KCP Tentena","15181":"KCP Kotaraya","15182":"KCP Buol","15184":"KCP Parigi Tolai","15185":"KCP Banggai Bunta","15186":"KCP Banggai Laut","15187":"KCP Morowali Beteleme","15188":"KCP Toli Toli Soni","15189":"KCP Morowali Bungku","15190":"KCP Banggai Kepulauan","15191":"KCP Sigi",
    "15200":"KC Makassar Kartini","15201":"KCP Makassar Sulawesi","15202":"KCP Makassar Cokroaminoto","15203":"KCP Semen Tonasa","15213":"KCP Makassar Slamet Riyadi","15216":"KCP Makassar Daya","15220":"KCP Makassar Andalas","15221":"KCP Makassar Veteran","15225":"KCP Makassar R.S. Stella Maris","15237":"KCP Makassar Universitas Hasanuddin","15238":"KCP Makassar Pelabuhan Indonesia","15242":"KCP Makassar Pasar Sentral","15243":"KCP Makassar Sombaopu","15244":"KCP Pangkep Hasanuddin","15245":"KCP Makassar Pusat Grosir Daya","15246":"KCP Makassar Perintis Kemerdekaan","15248":"KCP Maros Sudirman","15249":"KCP Makassar Urip Sumoharjo","15250":"KCP Makassar Monginsidi","15251":"KCP Makassar CPI","15252":"KCP Makassar Latimojong","15269":"KCP Maros Camba","15271":"KCP Makassar Pannampu","15291":"KCP Makassar BTP","15293":"KCP Makassar Paccerakkang","15297":"KCP Makassar Sudiang",
    "16200":"KC Kendari Mesjid Agung","16201":"KCP Kendari Beach","16202":"KCP Bau Bau","16203":"KC Pomalaa","16204":"KCP Kolaka","16205":"KCP Wua Wua","16206":"KCP Kolaka Utara","16207":"KCP Kendari Bundaran Anduonohu","16208":"KCP Kendari Lepo-Lepo","16210":"KCP Morosi","16275":"KCP Kendari Andunouhu","16276":"KCP Tinanggea","16277":"KCP Pasarwajo","16278":"KCP Molawe","16279":"KCP Raha","16280":"KCP Bombana","16281":"KCP Unaaha","16282":"KCP Buton Lombe","16283":"KCP Baubau Wolio","16284":"KCP Rate Rate","16285":"KCP Lapai (dh Lapaik)",
    "17000":"KC Pare Pare","17001":"KCP Pinrang","17002":"KCP Polewali Mandar","17003":"KCP Pare-Pare Pattompo","17004":"KCP Mamuju","17005":"KC Watampone","17006":"KCP Sengkang","17007":"KC Palopo","17008":"KCP Rantepao","17009":"KC Soroako","17010":"KCP Sidrap Sudirman","17011":"KCP Sinjai Pasar Sentral","17015":"KCP Kahu","17016":"KCP Majene","17075":"KCP Topoyo","17076":"KCP Sidrap","17077":"KCP Siwa","17078":"KCP Soppeng","17079":"KCP Belopa","17080":"KCP Masamba","17082":"KCP Barru","17083":"KCP Tomoni","17084":"KCP Malili","17085":"KCP Polewali Mandar Pekkabata","17087":"KCP Toraja Makale","17088":"KCP Enrekang Pasar Sudu","17091":"KCP Bone Bone",
    "17400":"KC Makassar Sam Ratulangi","17401":"KCP Makassar Panakkukang","17402":"KCP Makassar Hertasning","17403":"KCP Sungguminasa","17404":"KCP Makassar Cendrawasih","17406":"KCP Makassar Toddopuli","17407":"KCP Bulukumba","17408":"KCP Makassar Pettarani","17409":"KCP Makassar Veteran Selatan","17410":"KCP Bantaeng Pahlawan","17411":"KCP Jeneponto","17451":"KCP Selayar","17452":"KCP Gowa Malino","17453":"KCP Bantaeng Tanetea","17454":"KCP Takalar","17455":"KCP Limbung","17456":"KCP Hartaco","17457":"KCP Antang","17458":"KCP Gowa Pallangga","17459":"KCP Gowa Balang - Balang","17462":"KCP Bulukumba Tanete","17463":"KCP Barombong",
    "18600":"KC Ambon Pantai Mardika","18601":"KC Ambon Pattimura","18602":"KC Ternate","18603":"KCP Ambon Paso","18604":"KCP Ambon Masohi","18605":"KCP Ambon Namlea","18607":"KCP Maluku Tenggara","18608":"KCP Ternate Pahlawan Revolusi","18609":"KCP Buli","18610":"KCP Tobelo","18611":"KCP Sofifi","18612":"KCP Ambon Universitas Pattimura","18613":"KCP Indonesia Weda Bay Industrial Park","18651":"KCP Bula","18652":"KCP Gemba","18654":"KCP Waeapo","18655":"KCP Wayame","18656":"KCP Saumlaki","18657":"KCP Ternate Bastiong","18658":"KCP Labuha","18659":"KCP Weda","18660":"KCP Jailolo","18661":"KCP Obi"
  };
  return _KODE_TO_NAMA_CACHE;
}

// 7 Area resmi Region X, urutan tetap dipakai di tabel ringkasan.
var AREA_RESMI_URUT = [
  "Area Manado", "Area Palu", "Area Makassar Kartini", "Area Kendari",
  "Area Pare Pare", "Area Makassar Ratulangi", "Area Maluku"
];

// Daftar semua 180 cabang { kode, nama, area }, urut per Area (AREA_RESMI_URUT)
// lalu urutan asli per Area — dipakai buat bikin baris per-Cabang di tabel
// ringkasan & sheet "Data Statis Cabang" / "Riwayat Progress Cabang".
function getSemuaCabang() {
  var dataArea = getDataCabangArea();
  var namaMap  = getKodeToNamaMap();
  var hasil = [];
  AREA_RESMI_URUT.forEach(function (area) {
    (dataArea[area] || []).forEach(function (kode) {
      hasil.push({ kode: kode, nama: namaMap[kode] || ('Cabang ' + kode), area: area });
    });
  });
  return hasil;
}

/**
 * Ambil / buat sheet "Data Statis Cabang" — tempat admin isi MANUAL angka
 * "Shifting to Optimal" PER CABANG (data ini tidak berasal dari status
 * merchant per-baris di Pipeline, makanya tidak bisa dihitung otomatis).
 * Dibuat otomatis dengan 180 baris (satu per cabang, dikelompokkan per
 * Area) & nilai 0, tinggal diedit langsung di Google Sheets.
 *
 * CATATAN MIGRASI: sebelumnya sheet ini ("Data Statis Area") cuma 7 baris
 * per-Area. Sekarang dipecah per-Cabang biar tabel per-Cabang di
 * monitoring-area.html konsisten sama tabel per-Area (Area = jumlah semua
 * cabang di dalamnya). Kalau sheet lama "Data Statis Area" masih ada &
 * sudah keisi, boleh dihapus manual — sudah tidak dipakai lagi.
 */
function getOrCreateSheetStatis(ss) {
  var sheet = ss.getSheetByName(CONFIG_RINGKASAN.NAMA_SHEET_STATIS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_RINGKASAN.NAMA_SHEET_STATIS);
    var headers = ['Kode Cabang', 'Nama Cabang', 'Area', 'Shifting to Optimal'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    var rows = getSemuaCabang().map(function (c) { return [c.kode, c.nama, c.area, 0]; });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function setupSheetDataStatisArea() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreateSheetStatis(ss);
  // Pakai Logger.log, BUKAN beriTahu()/getUi().alert() — fungsi ini
  // dijalankan manual dari Apps Script Editor, dan alert() lewat getUi()
  // butuh diklik langsung di tab Google Sheets. Kalau tab itu nggak lagi
  // dibuka/dilihat, script nge-block nunggu klik OK sampai kena timeout
  // 6 menit. Cek hasilnya lewat Execution log, bukan popup.
  Logger.log('✅ Sheet "' + CONFIG_RINGKASAN.NAMA_SHEET_STATIS + '" siap (180 baris per Cabang). Silakan isi angka "Shifting to Optimal" per Cabang secara manual di sana.');
}

function getDataStatisCabang() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheetStatis(ss);
  var lastRow = sheet.getLastRow();
  var hasil = {};
  if (lastRow < 2) return hasil;
  var data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  for (var i = 0; i < data.length; i++) {
    var kode = String(data[i][0] || '').trim();
    if (kode) hasil[kode] = Number(data[i][3]) || 0;
  }
  return hasil;
}

/**
 * Hitung ringkasan progress per Area (Jumlah Leads, Jml FU, Sisa Leads,
 * %FU, breakdown 7 kategori Hasil Follow Up) + baris Total Region, plus
 * "Progress DtD" (selisih Jml FU hari ini vs snapshot kemarin).
 * Dipanggil dari doGet (action=getRingkasanArea) buat monitoring-area.html.
 */
function getRingkasanArea() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  var semuaCabang  = getSemuaCabang();          // [{kode, nama, area}, ...] urut per Area
  var shiftingKode = getDataStatisCabang();      // { kode: angka manual }
  var kemarinInfo  = getJmlFuKemarinSemuaCabang(); // { nilai: {kode: Jml FU kemarin}, waktu: label jam }
  var kemarinKode  = kemarinInfo.nilai;

  // Siapkan akumulator per KODE CABANG (bukan per Area lagi) — rollup ke
  // Area dilakukan belakangan dengan menjumlah cabang-cabangnya, jadi
  // angka Area dijamin selalu = total dari cabang di dalamnya.
  var akum = {};
  semuaCabang.forEach(function (c) {
    akum[c.kode] = { jumlahLeads: 0, target: 0 };
    KATEGORI_HASIL_FU_OTOMATIS.forEach(function (k) { akum[c.kode][k] = 0; });
  });

  if (sheet) {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow >= 2) {
      var headers     = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var idxKode     = headers.indexOf('Kode Cabang');
      // "Nama Merchant" pernah berganti nama jadi "dbaname" di sheet sumber —
      // ini PERSIS penyebab bug "semua angka jadi 0" yang pernah kejadian:
      // kalau kolomnya gak ketemu (-1), if di bawah gagal & seluruh blok
      // pembacaan Pipeline di-skip diam-diam tanpa error apapun.
      var idxMerchant = _cariIndexHeader(headers, ['Nama Merchant', 'dbaname']);
      var idxStatus   = headers.indexOf('Status Progress');

      if (idxKode !== -1 && idxMerchant !== -1 && idxStatus !== -1) {
        // Cuma baca 3 kolom yang dibutuhkan (bukan getDataRange penuh) —
        // sheet ini > 4000 baris, baca semua kolom (termasuk Alamat/Catatan
        // yang teksnya panjang) bikin lambat & rawan timeout.
        var kolomKode     = sheet.getRange(2, idxKode + 1, lastRow - 1, 1).getValues();
        var kolomMerchant = sheet.getRange(2, idxMerchant + 1, lastRow - 1, 1).getValues();
        var kolomStatus   = sheet.getRange(2, idxStatus + 1, lastRow - 1, 1).getValues();

        for (var i = 0; i < kolomKode.length; i++) {
          var kode = String(kolomKode[i][0] || '').trim();
          var merchant = String(kolomMerchant[i][0] || '').trim();
          if (!kode || !merchant) continue;
          if (!akum.hasOwnProperty(kode)) continue; // kode tidak dikenal, lewati

          akum[kode].jumlahLeads++;

          var status = String(kolomStatus[i][0] || '').trim();
          if (!status || status === 'Target') {
            akum[kode].target++;
          } else if (KATEGORI_HASIL_FU_OTOMATIS.indexOf(status) !== -1) {
            akum[kode][status]++;
          }
          // Status lain yang tidak dikenal (mis. sisa data kotor) sengaja
          // tidak dihitung ke kategori manapun, supaya tidak salah kelompok.
        }
      }
    }
  }

  // Susun baris per Cabang dulu, dikelompokkan per Area
  var cabangByArea = {};
  AREA_RESMI_URUT.forEach(function (a) { cabangByArea[a] = []; });

  semuaCabang.forEach(function (c) {
    var a = akum[c.kode];

    // Jumlah dari 6 kategori otomatis (status Progress asli di Pipeline —
    // ini ANGKA GROUND-TRUTH, tidak pernah diubah/dipotong).
    var autoSum = 0;
    KATEGORI_HASIL_FU_OTOMATIS.forEach(function (k) { autoSum += a[k]; });

    // Jumlah Leads = jumlah baris merchant di sheet Pipeline apa adanya.
    // TIDAK PERNAH diotak-atik — harus selalu persis sama dengan jumlah
    // baris merchant di sheet Pipeline (supaya Total Region selalu persis
    // 4.174, sama dengan total baris asli di Google Sheets).
    var jumlahLeads = a.jumlahLeads;

    // PENGAMAN %FU + konsistensi breakdown: di beberapa Cabang, angka
    // "Shifting to Optimal" manual di sheet "Data Statis Cabang" ternyata
    // TIDAK sepenuhnya cocok 1:1 dengan baris di Pipeline Cabang itu (mis.
    // sebagian merchant yang di-shift sudah lebih dulu punya status di
    // salah satu 6 kategori otomatis JUGA) — akibatnya shifting + autoSum
    // bisa lebih besar dari Jumlah Leads mentahnya (%FU > 100%, kejadian di
    // Area Maluku dkk).
    //
    // Sebelumnya yang di-floor itu Jml FU-nya SETELAH dijumlah — akibatnya
    // kolom breakdown "Hasil Follow Up" (termasuk Shifting to Optimal) yang
    // ditampilkan di tabel jadi kalau dijumlah manual LEBIH BESAR dari Jml
    // FU yang tampil (bikin bingung, kelihatan kayak salah hitung).
    //
    // Sekarang yang di-floor adalah "shifting"-nya SEBELUM dijumlah ke Jml
    // FU — supaya kolom Shifting to Optimal yang ditampilkan itu PERSIS
    // porsi yang belum kepakai sama 6 kategori otomatis (gak dobel-hitung),
    // dan breakdown 7 kategori selalu pas persis sama Jml FU total. Kalau
    // autoSum sendirian sudah >= jumlahLeads (jarang, kondisi data lain
    // yang lebih parah), shifting jadi 0 — 6 kategori otomatis tidak pernah
    // ikut dipotong karena itu ground-truth dari Status Progress asli.
    var shiftingMentah = shiftingKode[c.kode] || 0;
    var shifting = Math.max(0, Math.min(shiftingMentah, jumlahLeads - autoSum));
    var jmlFU = shifting + autoSum;

    var sisaLeads   = jmlFU - jumlahLeads;
    var persenFU    = jumlahLeads > 0 ? (jmlFU / jumlahLeads * 100) : 0;
    var dtd         = jmlFU - (kemarinKode[c.kode] || 0);

    var row = {
      kode              : c.kode,
      nama              : c.nama,
      jumlahLeads       : jumlahLeads,
      jmlFU             : jmlFU,
      sisaLeads         : sisaLeads,
      persenFU          : Math.round(persenFU * 10) / 10,
      shiftingToOptimal : shifting,
      dtd               : dtd
    };
    KATEGORI_HASIL_FU_OTOMATIS.forEach(function (k) { row[k] = a[k]; });
    cabangByArea[c.area].push(row);
  });

  // Rollup per Area = jumlah semua cabang di dalamnya (bottom-up, bukan
  // dihitung ulang terpisah) — jadi Area selalu konsisten sama Cabang.
  var totalRegion = { jumlahLeads: 0, jmlFU: 0, shiftingToOptimal: 0, dtd: 0 };
  KATEGORI_HASIL_FU_OTOMATIS.forEach(function (k) { totalRegion[k] = 0; });

  var hasilArea = AREA_RESMI_URUT.map(function (area) {
    var daftarCabang = cabangByArea[area];
    var agg = { jumlahLeads: 0, jmlFU: 0, shiftingToOptimal: 0, dtd: 0 };
    KATEGORI_HASIL_FU_OTOMATIS.forEach(function (k) { agg[k] = 0; });

    daftarCabang.forEach(function (c) {
      agg.jumlahLeads       += c.jumlahLeads;
      agg.jmlFU             += c.jmlFU;
      agg.shiftingToOptimal += c.shiftingToOptimal;
      agg.dtd               += c.dtd;
      KATEGORI_HASIL_FU_OTOMATIS.forEach(function (k) { agg[k] += c[k]; });
    });

    var sisaLeads = agg.jmlFU - agg.jumlahLeads;
    var persenFU  = agg.jumlahLeads > 0 ? (agg.jmlFU / agg.jumlahLeads * 100) : 0;

    totalRegion.jumlahLeads       += agg.jumlahLeads;
    totalRegion.jmlFU             += agg.jmlFU;
    totalRegion.shiftingToOptimal += agg.shiftingToOptimal;
    totalRegion.dtd               += agg.dtd;
    KATEGORI_HASIL_FU_OTOMATIS.forEach(function (k) { totalRegion[k] += agg[k]; });

    var row = {
      area              : area,
      jumlahLeads       : agg.jumlahLeads,
      jmlFU             : agg.jmlFU,
      sisaLeads         : sisaLeads,
      persenFU          : Math.round(persenFU * 10) / 10,
      shiftingToOptimal : agg.shiftingToOptimal,
      dtd               : agg.dtd,
      cabang            : daftarCabang
    };
    KATEGORI_HASIL_FU_OTOMATIS.forEach(function (k) { row[k] = agg[k]; });
    return row;
  });

  var totalSisaLeads = totalRegion.jmlFU - totalRegion.jumlahLeads;
  var totalPersenFU  = totalRegion.jumlahLeads > 0 ? (totalRegion.jmlFU / totalRegion.jumlahLeads * 100) : 0;

  var rowTotal = {
    area              : 'Region X / Sulawesi & Maluku',
    jumlahLeads       : totalRegion.jumlahLeads,
    jmlFU             : totalRegion.jmlFU,
    sisaLeads         : totalSisaLeads,
    persenFU          : Math.round(totalPersenFU * 10) / 10,
    shiftingToOptimal : totalRegion.shiftingToOptimal,
    dtd               : totalRegion.dtd
  };
  KATEGORI_HASIL_FU_OTOMATIS.forEach(function (k) { rowTotal[k] = totalRegion[k]; });

  var sekarang = new Date();
  return {
    diperbarui   : Utilities.formatDate(sekarang, Session.getScriptTimeZone(), "dd MMMM yyyy, HH:mm 'WITA'"),
    // Label jam ringkas (mis. "31 Jul 10.30 WITA") & jam kemarin (mis.
    // "30 Jul 20.30 WITA", kosong kalau belum ada snapshot kemarin) —
    // dipakai susunWaLaporanRegion()/susunWaLaporanArea() buat kutip
    // rentang waktu DtD di wording WA.
    waktuSekarang : _labelWaktuSingkat(sekarang),
    waktuKemarin  : kemarinInfo.waktu || '',
    area       : hasilArea,
    total      : rowTotal
  };
}

// ============================================================
// REKAP FORMULA — versi RUMUS SHEETS ASLI (COUNTIFS/SUMIF/VLOOKUP) dari
// tabel yang sama persis kayak monitoring-area.html, TAPI dihitung
// langsung sama Google Sheets (bukan angka jadi dari Apps Script) —
// jadi bisa di-cross-check manual sel per sel, klik & lihat rumusnya.
// Dipanggil lewat menu "📊 Rekap Cabang → Buat/Refresh Rekap Formula".
// ============================================================

// Ubah nomor kolom (1-based) jadi huruf kolom A1 notation, mis. 1->"A",
// 27->"AA". Dipakai buat nyusun rumus yang nunjuk ke posisi kolom ASLI di
// sheet Pipeline/Data Statis Cabang SAAT fungsi ini dijalankan (posisi
// kolomnya dibaca dari header, bukan ditebak/di-hardcode).
function _kolomHuruf(n) {
  var s = '';
  while (n > 0) {
    var sisa = (n - 1) % 26;
    s = String.fromCharCode(65 + sisa) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Bikin (atau refresh total) 2 sheet:
 *   - "Rekap Cabang (Formula)" -> 1 baris per Cabang (180 baris), rumus
 *     COUNTIF/COUNTIFS langsung ke sheet "Pipeline EDC to LVM" + VLOOKUP
 *     ke "Data Statis Cabang" buat Shifting to Optimal (dengan pengaman
 *     MAX/MIN yang sama kayak getRingkasanArea(), biar %FU gak pernah
 *     >100% dan breakdown selalu pas sama Jml FU — lihat komentar di
 *     getRingkasanArea() soal kenapa perlu di-cap).
 *   - "Rekap Area (Formula)" -> 1 baris per Area (7 baris) + Total
 *     Region, dihitung SUMIF dari sheet Cabang di atas berdasarkan kolom
 *     Area (dari peta resmi getSemuaCabang(), BUKAN kolom "Area" di
 *     Pipeline yang kadang keisi manual/kotor) — sama persis logic
 *     rollup Area = jumlah Cabang di dalamnya kayak di getRingkasanArea().
 *
 * "Progress DtD" SENGAJA tidak diikutkan di sini (butuh baca histori
 * "Riwayat Progress Cabang" tiap hari, beda logic dari rekap statis
 * ini) — kalau perlu itu juga, pakai angka dari monitoring-area.html
 * langsung.
 *
 * Aman dijalankan berkali-kali (isi sheet ditimpa ulang tiap run). WAJIB
 * dijalankan ULANG kalau kolom di sheet "Pipeline EDC to LVM" atau "Data
 * Statis Cabang" ditambah/dihapus/digeser — rumus di sini "membeku" pada
 * posisi kolom SAAT fungsi ini dijalankan.
 */
function buatRekapFormula() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---------- Cari posisi kolom ASLI di sheet Pipeline ----------
  var shPipeline = ss.getSheetByName(CONFIG.NAMA_SHEET_PIPELINE);
  if (!shPipeline) { beriTahu('Sheet "' + CONFIG.NAMA_SHEET_PIPELINE + '" tidak ditemukan!'); return; }
  var headerPipeline = shPipeline.getRange(1, 1, 1, shPipeline.getLastColumn()).getValues()[0];
  var idxKodeP   = headerPipeline.indexOf('Kode Cabang');
  var idxStatusP = headerPipeline.indexOf('Status Progress');
  if (idxKodeP === -1 || idxStatusP === -1) {
    beriTahu('Kolom "Kode Cabang" atau "Status Progress" tidak ketemu di header sheet "' + CONFIG.NAMA_SHEET_PIPELINE + '".');
    return;
  }
  var namaPipeline = CONFIG.NAMA_SHEET_PIPELINE;
  var refKodeP   = "'" + namaPipeline + "'!$" + _kolomHuruf(idxKodeP + 1)   + ':$' + _kolomHuruf(idxKodeP + 1);
  var refStatusP = "'" + namaPipeline + "'!$" + _kolomHuruf(idxStatusP + 1) + ':$' + _kolomHuruf(idxStatusP + 1);

  // ---------- Cari posisi kolom ASLI di sheet Data Statis Cabang ----------
  var shStatis = getOrCreateSheetStatis(ss);
  var headerStatis = shStatis.getRange(1, 1, 1, shStatis.getLastColumn()).getValues()[0];
  var idxKodeS  = headerStatis.indexOf('Kode Cabang');
  var idxShiftS = headerStatis.indexOf('Shifting to Optimal');
  if (idxKodeS === -1 || idxShiftS === -1) {
    beriTahu('Kolom "Kode Cabang" atau "Shifting to Optimal" tidak ketemu di header sheet "' + CONFIG_RINGKASAN.NAMA_SHEET_STATIS + '".');
    return;
  }
  var namaStatis  = CONFIG_RINGKASAN.NAMA_SHEET_STATIS;
  var kolomKiriS  = Math.min(idxKodeS, idxShiftS) + 1;
  var kolomKananS = Math.max(idxKodeS, idxShiftS) + 1;
  var offsetShiftS = (idxShiftS + 1) - kolomKiriS + 1; // kolom ke berapa DI DALAM range VLOOKUP
  var refStatis = "'" + namaStatis + "'!$" + _kolomHuruf(kolomKiriS) + '$2:$' + _kolomHuruf(kolomKananS);

  // ================= Sheet "Rekap Cabang (Formula)" =================
  var namaSheetCabang = 'Rekap Cabang (Formula)';
  var shCabang = ss.getSheetByName(namaSheetCabang);
  if (shCabang) { shCabang.clear(); } else { shCabang = ss.insertSheet(namaSheetCabang); }

  var headerCabang = [
    'Area', 'Kode Cabang', 'Nama Cabang', 'Jumlah Leads',
    'Done Konversi to LVM', 'Tetap pakai EDC & Optimalkan SV', 'Penawaran',
    'Merchant Menolak', 'Merchant Tutup', 'Perlu Kunjungan MTI',
    'Auto Sum (6 Kategori)', 'Shifting Mentah', 'Shifting to Optimal (Capped)',
    'Jml FU', 'Sisa Leads', '% FU'
  ]; // A..P

  var semuaCabang = getSemuaCabang();
  var rowsCabang = semuaCabang.map(function (c, i) {
    var r = i + 2; // nomor baris di sheet (baris 1 = header)
    var kode = 'B' + r;
    return [
      c.area, c.kode, c.nama,
      '=COUNTIF(' + refKodeP + ',' + kode + ')',
      '=COUNTIFS(' + refKodeP + ',' + kode + ',' + refStatusP + ',"Done Konversi to LVM")',
      '=COUNTIFS(' + refKodeP + ',' + kode + ',' + refStatusP + ',"Tetap pakai EDC & Optimalkan SV")',
      '=COUNTIFS(' + refKodeP + ',' + kode + ',' + refStatusP + ',"Penawaran")',
      '=COUNTIFS(' + refKodeP + ',' + kode + ',' + refStatusP + ',"Merchant Menolak")',
      '=COUNTIFS(' + refKodeP + ',' + kode + ',' + refStatusP + ',"Merchant Tutup")',
      '=COUNTIFS(' + refKodeP + ',' + kode + ',' + refStatusP + ',"Perlu Kunjungan MTI")',
      '=SUM(E' + r + ':J' + r + ')',                                   // Auto Sum 6 kategori
      '=IFERROR(VLOOKUP(' + kode + ',' + refStatis + ',' + offsetShiftS + ',FALSE),0)', // Shifting Mentah
      '=MAX(0,MIN(L' + r + ',D' + r + '-K' + r + '))',                 // Shifting to Optimal (Capped)
      '=M' + r + '+K' + r,                                             // Jml FU
      '=N' + r + '-D' + r,                                             // Sisa Leads
      '=IF(D' + r + '>0,N' + r + '/D' + r + '*100,0)'                  // % FU
    ];
  });

  shCabang.getRange(1, 1, 1, headerCabang.length).setValues([headerCabang]);
  shCabang.getRange(1, 1, 1, headerCabang.length)
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  shCabang.getRange(2, 1, rowsCabang.length, headerCabang.length).setValues(rowsCabang);
  shCabang.getRange(2, 16, rowsCabang.length, 1).setNumberFormat('0.0"%"'); // kolom P (% FU)
  shCabang.setFrozenRows(1);
  shCabang.autoResizeColumns(1, headerCabang.length);

  // ================= Sheet "Rekap Area (Formula)" =================
  var namaSheetArea = 'Rekap Area (Formula)';
  var shArea = ss.getSheetByName(namaSheetArea);
  if (shArea) { shArea.clear(); } else { shArea = ss.insertSheet(namaSheetArea); }

  var headerArea = [
    'Area', 'Jumlah Leads', 'Jml FU', 'Sisa Leads', '% FU',
    'Shifting to Optimal', 'Done Konversi to LVM', 'Tetap pakai EDC & Optimalkan SV',
    'Penawaran', 'Merchant Menolak', 'Merchant Tutup', 'Perlu Kunjungan MTI'
  ]; // A..L

  var nBarisCabang = rowsCabang.length; // 180
  var akhirCabang = nBarisCabang + 1;   // baris terakhir di sheet Cabang (+1 krn header)
  function refCabangKolom(huruf) {
    return "'" + namaSheetCabang + "'!$" + huruf + '$2:$' + huruf + '$' + akhirCabang;
  }
  var refAreaKolom     = refCabangKolom('A');
  var refJumlahLeadsC  = refCabangKolom('D');
  var refDoneC         = refCabangKolom('E');
  var refTetapC        = refCabangKolom('F');
  var refPenawaranC    = refCabangKolom('G');
  var refMenolakC      = refCabangKolom('H');
  var refTutupC        = refCabangKolom('I');
  var refMtiC          = refCabangKolom('J');
  var refShiftC        = refCabangKolom('M');
  var refJmlFuC        = refCabangKolom('N');

  var jumlahArea = AREA_RESMI_URUT.length; // 7
  var barisTotal = jumlahArea + 2;         // baris "Total Region"

  var rowsArea = AREA_RESMI_URUT.map(function (area, i) {
    var r = i + 2;
    var kriteria = 'A' + r;
    return [
      area,
      '=SUMIF(' + refAreaKolom + ',' + kriteria + ',' + refJumlahLeadsC + ')',
      '=SUMIF(' + refAreaKolom + ',' + kriteria + ',' + refJmlFuC + ')',
      '=C' + r + '-B' + r,
      '=IF(B' + r + '>0,C' + r + '/B' + r + '*100,0)',
      '=SUMIF(' + refAreaKolom + ',' + kriteria + ',' + refShiftC + ')',
      '=SUMIF(' + refAreaKolom + ',' + kriteria + ',' + refDoneC + ')',
      '=SUMIF(' + refAreaKolom + ',' + kriteria + ',' + refTetapC + ')',
      '=SUMIF(' + refAreaKolom + ',' + kriteria + ',' + refPenawaranC + ')',
      '=SUMIF(' + refAreaKolom + ',' + kriteria + ',' + refMenolakC + ')',
      '=SUMIF(' + refAreaKolom + ',' + kriteria + ',' + refTutupC + ')',
      '=SUMIF(' + refAreaKolom + ',' + kriteria + ',' + refMtiC + ')'
    ];
  });

  var akhirArea = jumlahArea + 1; // baris Area terakhir (sebelum Total)
  rowsArea.push([
    'Region X / Sulawesi & Maluku',
    '=SUM(B2:B' + akhirArea + ')',
    '=SUM(C2:C' + akhirArea + ')',
    '=C' + barisTotal + '-B' + barisTotal, // Sisa Leads Total
    '=IF(B' + barisTotal + '>0,C' + barisTotal + '/B' + barisTotal + '*100,0)',
    '=SUM(F2:F' + akhirArea + ')',
    '=SUM(G2:G' + akhirArea + ')',
    '=SUM(H2:H' + akhirArea + ')',
    '=SUM(I2:I' + akhirArea + ')',
    '=SUM(J2:J' + akhirArea + ')',
    '=SUM(K2:K' + akhirArea + ')',
    '=SUM(L2:L' + akhirArea + ')'
  ]);

  shArea.getRange(1, 1, 1, headerArea.length).setValues([headerArea]);
  shArea.getRange(1, 1, 1, headerArea.length)
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  shArea.getRange(2, 1, rowsArea.length, headerArea.length).setValues(rowsArea);
  shArea.getRange(2, 5, rowsArea.length, 1).setNumberFormat('0.0"%"'); // kolom E (% FU)
  shArea.getRange(barisTotal, 1, 1, headerArea.length).setFontWeight('bold').setBackground('#dbe4fb');
  shArea.setFrozenRows(1);
  shArea.autoResizeColumns(1, headerArea.length);

  beriTahu('✅ Rekap Formula siap!\n\n- "' + namaSheetCabang + '" (' + nBarisCabang + ' baris per Cabang)\n- "' + namaSheetArea + '" (' + jumlahArea + ' Area + Total Region)\n\nSemua angka dihitung pakai rumus COUNTIFS/SUMIF/VLOOKUP asli Google Sheets, bisa diklik & dicek satu-satu. Jalankan ulang menu ini kalau kolom di sheet Pipeline/Data Statis Cabang berubah.');
}

/**
 * Simpan snapshot Jml FU hari ini PER CABANG (180 baris) ke sheet
 * "Riwayat Progress Cabang" — dasar hitung "Progress DtD" (Day to Day)
 * besok, baik di level Cabang maupun Area (Area = jumlah cabangnya).
 * Aman dijalankan berkali-kali di hari yang sama (baris hari ini ditimpa,
 * bukan dobel). Dipasang sebagai time-driven trigger harian lewat
 * pasangTriggerSnapshotHarian().
 *
 * PENTING: semua baca/tulis di sini di-batch (1 setValues() buat semua
 * baris), BUKAN loop per-baris — dengan 180 cabang, nulis satu-satu lewat
 * getRange().setValue() di dalam loop bisa bikin ratusan API call terpisah
 * & berisiko timeout 6 menit (persis kayak bug yang pernah kejadian di
 * buatSheetRekap()).
 */
// Bikin key tanggal yang AMAN buat dibandingin (yyyy-MM-dd) — SENGAJA
// tidak lewat `new Date(teksnya)` buat re-parse ulang. Alasannya: kalau sel
// tanggalnya ternyata masih berupa teks "dd/MM/yyyy" (bukan ke-convert jadi
// Date beneran sama Google Sheets, misalnya kolomnya diformat Plain Text),
// `new Date("02/08/2026")` dibaca JS ala Amerika (MM/dd/yyyy) — jadi
// 8 Februari, padahal maksudnya 2 Agustus. Salah baca tanggal begini bikin
// DtD salah hitung & snapshot harian bisa kesimpen dobel tanpa ketahuan.
// Fungsi ini format langsung kalau selnya sudah Date, atau parse manual
// pola dd/MM/yyyy kalau masih teks — tidak pernah lewat `new Date(string)`.
function _tglKey(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // dd/MM/yyyy
  if (m) {
    var dd = ('0' + m[1]).slice(-2);
    var mm = ('0' + m[2]).slice(-2);
    return m[3] + '-' + mm + '-' + dd;
  }
  return s;
}

// Nama bulan pendek ala Indonesia (dipakai wording WA & label waktu snapshot
// ringkas, mis. "30 Jul 20.30 WITA") — beda dari format panjang "dd MMMM
// yyyy" (mis. "31 Juli 2026") yang dipakai di header laporan.
var BULAN_PENDEK_ID = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
function _labelWaktuSingkat(d) {
  var tz = Session.getScriptTimeZone();
  var tgl  = d.getDate() + ' ' + BULAN_PENDEK_ID[d.getMonth()];
  var jam  = Utilities.formatDate(d, tz, 'HH.mm');
  return tgl + ' ' + jam + ' WITA';
}

function simpanSnapshotHarian() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_RINGKASAN.NAMA_SHEET_RIWAYAT);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_RINGKASAN.NAMA_SHEET_RIWAYAT);
    sheet.getRange(1, 1, 1, 5).setValues([['Tanggal', 'Kode Cabang', 'Nama Cabang', 'Jml FU', 'Waktu Snapshot']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  // Migrasi: sheet yang dibuat SEBELUM kolom "Waktu Snapshot" ada cuma
  // punya 4 kolom — tambahkan kolom ke-5 supaya baris baru & lama sama-sama
  // punya tempat buat catat jam snapshot (dipakai wording WA buat kutip
  // "30 Jul 20.30 WITA → 31 Jul 10.30 WITA"). Baris lama yang sudah kadung
  // tidak punya jam tercatat, dibiarkan kosong (fallback ditangani di
  // getJmlFuKemarinSemuaCabang()/susunWaLaporan*()).
  var headerSekarang = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headerSekarang.indexOf('Waktu Snapshot') === -1) {
    var kolomBaru = sheet.getLastColumn() + 1;
    sheet.getRange(1, kolomBaru).setValue('Waktu Snapshot');
    sheet.getRange(1, kolomBaru).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  }

  var ringkasan = getRingkasanArea();
  var semuaCabang = [];
  ringkasan.area.forEach(function (a) {
    a.cabang.forEach(function (c) { semuaCabang.push(c); });
  });

  var sekarang     = new Date();
  var tglHariIni    = Utilities.formatDate(sekarang, Session.getScriptTimeZone(), 'dd/MM/yyyy'); // buat ditulis, biar gampang dibaca manual
  var tglHariIniKey = Utilities.formatDate(sekarang, Session.getScriptTimeZone(), 'yyyy-MM-dd');  // buat dibandingin, aman
  var waktuSnapshot = _labelWaktuSingkat(sekarang); // mis. "31 Jul 10.30 WITA", sama buat semua baris (1x eksekusi)
  var rows = semuaCabang.map(function (c) { return [tglHariIni, c.kode, c.nama, c.jmlFU, waktuSnapshot]; });

  // Cek apakah baris utk hari ini sudah pernah disimpan (biar bisa ditimpa,
  // bukan dobel, kalau trigger/manual dijalankan berkali-kali di hari yang
  // sama). Baris hari ini selalu ditulis berurutan dalam 1 batch, jadi
  // cukup cari baris awalnya & berapa banyak baris yang match tanggal ini.
  var lastRow = sheet.getLastRow();
  var barisMulaiHariIni = -1, jumlahBarisHariIni = 0;
  if (lastRow >= 2) {
    var tglSaja = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < tglSaja.length; i++) {
      if (_tglKey(tglSaja[i][0]) === tglHariIniKey) {
        if (barisMulaiHariIni === -1) barisMulaiHariIni = i + 2; // +1 header, +1 1-based
        jumlahBarisHariIni++;
      }
    }
  }

  if (barisMulaiHariIni !== -1 && jumlahBarisHariIni === rows.length) {
    // Sudah ada snapshot hari ini & jumlah cabangnya sama persis -> timpa
    sheet.getRange(barisMulaiHariIni, 1, rows.length, 5).setValues(rows);
  } else {
    // Belum ada snapshot hari ini (atau jumlah cabang berubah) -> tambah
    // baris baru di akhir, sekali batch buat semua 180 cabang.
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  }

  Logger.log('Snapshot progress cabang tersimpan: ' + tglHariIni + ' (' + rows.length + ' cabang, jam ' + waktuSnapshot + ')');
}

/**
 * Ambil Jml FU KEMARIN per Kode Cabang dari sheet riwayat snapshot.
 * Selisihnya terhadap Jml FU HARI INI (dihitung di getRingkasanArea) itu
 * yang jadi angka "Progress DtD" — baik di level Cabang maupun Area
 * (Area DtD = jumlah DtD semua cabangnya). Kalau belum ada snapshot
 * kemarin (misal trigger baru dipasang / hari pertama), dianggap 0 — jadi
 * DtD = Jml FU hari ini penuh (wajar, karena belum ada pembanding).
 */
/**
 * Ambil Jml FU KEMARIN per Kode Cabang, PLUS label jam snapshot kemarin
 * (mis. "30 Jul 20.30 WITA", diambil dari kolom "Waktu Snapshot" — lihat
 * simpanSnapshotHarian()). Dipakai getRingkasanArea() buat hitung DtD, dan
 * susunWaLaporanRegion()/susunWaLaporanArea() buat kutip jam pembanding di
 * wording WA ("DtD: +233 (30 Jul 20.30 WITA: 2.654 → ...)").
 * Return: { nilai: {kode: jmlFU}, waktu: 'label jam' atau '' kalau tidak ada
 * (snapshot kemarin belum pernah tersimpan, atau baris lama sebelum kolom
 * "Waktu Snapshot" ditambahkan).
 */
function getJmlFuKemarinSemuaCabang() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_RINGKASAN.NAMA_SHEET_RIWAYAT);
  var hasil = {};
  if (!sheet) return { nilai: hasil, waktu: '' };

  var kemarin = new Date();
  kemarin.setDate(kemarin.getDate() - 1);
  var tglKemarinKey = Utilities.formatDate(kemarin, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { nilai: hasil, waktu: '' };
  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, 5)).getValues();
  var waktu = '';
  for (var i = 0; i < data.length; i++) {
    if (_tglKey(data[i][0]) === tglKemarinKey) {
      hasil[String(data[i][1])] = Number(data[i][3]) || 0;
      if (!waktu && data[i][4]) waktu = String(data[i][4]);
    }
  }
  return { nilai: hasil, waktu: waktu };
}

/**
 * Jalankan SEKALI dari Apps Script Editor untuk memasang trigger harian
 * yang otomatis menjalankan simpanSnapshotHarian() tiap jam yang
 * ditentukan (default jam 23 / 11 malam, mendekati akhir hari kerja).
 * Aman dijalankan berkali-kali — trigger lama utk fungsi ini dihapus dulu
 * sebelum bikin yang baru, jadi tidak dobel-dobel.
 */
function pasangTriggerSnapshotHarian() {
  var JAM_TRIGGER = 23; // ganti sesuai kebutuhan, 0-23

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'simpanSnapshotHarian') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('simpanSnapshotHarian')
    .timeBased()
    .everyDays(1)
    .atHour(JAM_TRIGGER)
    .create();

  // Logger.log, bukan beriTahu()/getUi().alert() — sama alasannya kayak
  // setupSheetDataStatisArea(): dijalankan manual dari Editor, alert()
  // bisa nge-block nunggu diklik di tab Sheets sampai timeout.
  Logger.log('✅ Trigger harian terpasang — simpanSnapshotHarian() akan otomatis jalan tiap hari sekitar jam ' + JAM_TRIGGER + ':00.');
}

// ============================================================
// WORDING LAPORAN WHATSAPP — teks siap kirim (Region & per Area),
// dipakai tombol "Kirim Laporan" di monitoring-area.html. Dibuka lewat
// wa.me/?text=... TANPA nomor tujuan, jadi user pilih sendiri kontak/
// grup WA yang mau dikirimin (WhatsApp tidak bisa deep-link langsung ke
// grup).
// ============================================================

// Kepala Area per Area resmi — dipakai buat sapaan di wording laporan
// per-Area ("Semangat pagi Pak Daru & Area Manado, ..."). Ganti nama di
// sini kalau ada pergantian Kepala Area (lalu Deploy ulang Web App).
var AREA_HEAD = {
  "Area Manado"              : "Daru",
  "Area Palu"                : "Sibly",
  "Area Makassar Kartini"    : "Ferdinand",
  "Area Kendari"              : "Joko",
  "Area Pare Pare"            : "Said",
  "Area Makassar Ratulangi"  : "Iwa",
  "Area Maluku"               : "Tito"
};

// Sapaan tetap buat wording laporan Region — ganti di sini kalau ada
// pergantian pejabat.
var SAPAAN_REGION = ['Yth. Pak Nunung,', 'Pak Djati,', 'Bapak/Ibu AH,', 'dan rekan-rekan Cabang Region X,'];

// "Area Manado" -> "Manado" (buang prefix "Area ")
function labelAreaSingkat(area) {
  return String(area || '').replace(/^Area\s+/, '');
}

// "KC Bitung" -> "Bitung", "KCP Manado Sam Ratulangi" -> "Manado Sam Ratulangi"
function labelCabangSingkat(nama) {
  return String(nama || '').replace(/^KCP?\s+/, '');
}

// Format angka ala Indonesia: pemisah ribuan titik, mis. 4174 -> "4.174"
function formatAngkaID(n) {
  return Number(n || 0).toLocaleString('id-ID');
}

// Format persen ala Indonesia: koma buat desimal, mis. 69.2 -> "69,2"
function formatPersenID(n) {
  return (Math.round(Number(n || 0) * 10) / 10).toFixed(1).replace('.', ',');
}

// DtD selalu pakai tanda +/- eksplisit, mis. +233, -12, +0
function formatDtdSigned(n) {
  var num = Number(n || 0);
  var abs = formatAngkaID(Math.abs(num));
  return (num < 0 ? '-' : '+') + abs;
}

// Baris "Hasil Follow Up" — urutan & label sesuai wording yang diminta,
// beda sedikit dari label kolom tabel monitoring-area.html.
var LABEL_HASIL_FU_WA = [
  { key: 'shiftingToOptimal',                  label: 'Shifting to Optimal' },
  { key: 'Done Konversi to LVM',               label: 'Konversi ke LVM' },
  { key: 'Tetap pakai EDC & Optimalkan SV',     label: 'Tetap Pakai EDC (Komit Optimalkan SV)' },
  { key: 'Penawaran',                          label: 'Penawaran' },
  { key: 'Merchant Menolak',                   label: 'Merchant Menolak' },
  { key: 'Merchant Tutup',                     label: 'Merchant Tutup' },
  { key: 'Perlu Kunjungan MTI',                label: 'Perlu Kunjungan MTI' }
];

function _blokHasilFU(row) {
  return LABEL_HASIL_FU_WA.map(function (item) {
    return '- ' + item.label + ': ' + formatAngkaID(row[item.key]);
  }).join('\n');
}

// Baris "DtD: +233 (30 Jul 20.30 WITA: 2.654 → 31 Jul 10.30 WITA: 2.887)".
// Kalau belum ada snapshot kemarin (waktuKemarin kosong), tampilkan versi
// singkat tanpa breakdown jam (belum ada pembanding).
function _barisDtd(dtd, jmlFuSekarang, jmlFuKemarin, waktuKemarin, waktuSekarang) {
  if (!waktuKemarin) {
    return 'DtD: ' + formatDtdSigned(dtd) + ' (belum ada snapshot hari sebelumnya utk pembanding)';
  }
  return 'DtD: ' + formatDtdSigned(dtd) + ' (' + waktuKemarin + ': ' + formatAngkaID(jmlFuKemarin)
    + ' → ' + waktuSekarang + ': ' + formatAngkaID(jmlFuSekarang) + ')';
}

/**
 * Susun wording laporan WhatsApp tingkat REGION — dikirim ke Pak Nunung,
 * Pak Djati, Bapak/Ibu AH & rekan-rekan Cabang Region X. Berisi ringkasan
 * total Region + breakdown 7 kategori Hasil Follow Up + daftar per Area
 * diurut %FU terendah -> tertinggi.
 */
function susunWaLaporanRegion() {
  var r = getRingkasanArea();
  var t = r.total;
  // DtD kemarin dihitung mundur dari dtd & nilai sekarang (dtd = sekarang - kemarin)
  var jmlFuKemarin = t.jmlFU - t.dtd;

  var lines = [];
  lines = lines.concat(SAPAAN_REGION);
  lines.push('Terlampir kami sampaikan Progress FU Optimalisasi Merchant – Region X');
  lines.push('YTD ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy') + ' (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH.mm') + ' WITA)');
  lines.push('Total Leads: ' + formatAngkaID(t.jumlahLeads));
  lines.push('Jml FU: ' + formatAngkaID(t.jmlFU) + ' (' + formatPersenID(t.persenFU) + '%)');
  lines.push('Sisa Leads: ' + formatAngkaID(t.jumlahLeads - t.jmlFU));
  lines.push(_barisDtd(t.dtd, t.jmlFU, jmlFuKemarin, r.waktuKemarin, r.waktuSekarang));
  lines.push('Hasil Follow Up:');
  lines.push(_blokHasilFU(t));
  lines.push('Per Area (% FU terendah → tertinggi):');

  var areaUrut = r.area.slice().sort(function (a, b) { return a.persenFU - b.persenFU; });
  areaUrut.forEach(function (a, idx) {
    lines.push((idx + 1) + '. ' + labelAreaSingkat(a.area) + ' – ' + formatAngkaID(a.jmlFU)
      + ' (' + formatPersenID(a.persenFU) + '%), DtD ' + formatDtdSigned(a.dtd));
  });

  lines.push('Demikian, terima kasih.');
  return lines.join('\n');
}

/**
 * Susun wording laporan WhatsApp per AREA — dikirim ke Kepala Area & rekan
 * cabang di Area tsb. Berisi ringkasan Area + breakdown 7 kategori Hasil
 * Follow Up + daftar Cabang yang belum ada progress FU hari ini (DtD = 0),
 * diurut %FU (kumulatif) terendah -> tertinggi, + ajakan fokus ke 3 cabang
 * %FU terendah.
 */
function susunWaLaporanArea(namaArea) {
  var r = getRingkasanArea();
  var a = r.area.filter(function (x) { return x.area === namaArea; })[0];
  if (!a) throw new Error('Area tidak ditemukan: ' + namaArea);

  var jmlFuKemarin = a.jmlFU - a.dtd;
  var kepalaArea = AREA_HEAD[namaArea] || '';
  var labelArea  = labelAreaSingkat(namaArea);

  var lines = [];
  lines.push('Semangat pagi' + (kepalaArea ? ' Pak ' + kepalaArea : '') + ' & Area ' + labelArea + ',');
  lines.push('Terlampir kami sampaikan Progress FU Optimalisasi Merchant – Area ' + labelArea);
  lines.push('YTD ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy') + ' (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH.mm') + ' WITA)');
  lines.push('Total Leads: ' + formatAngkaID(a.jumlahLeads));
  lines.push('Jml FU: ' + formatAngkaID(a.jmlFU) + ' (' + formatPersenID(a.persenFU) + '%)');
  lines.push('Sisa Leads: ' + formatAngkaID(a.jumlahLeads - a.jmlFU));
  lines.push(_barisDtd(a.dtd, a.jmlFU, jmlFuKemarin, r.waktuKemarin, r.waktuSekarang));
  lines.push('Hasil Follow Up:');
  lines.push(_blokHasilFU(a));

  // Cabang yang belum ada progress FU hari ini (DtD = 0), urut %FU kumulatif
  // terendah -> tertinggi — supaya cabang paling butuh perhatian nongol duluan.
  // Dua kondisi sengaja DIKECUALIKAN dari daftar ini, walau DtD-nya 0:
  //   1. Cabang tanpa leads sama sekali (jumlahLeads = 0) — DtD-nya akan
  //      selalu 0 (tidak ada apa pun buat di-follow-up), jadi kalau ikut
  //      dimasukkan akan terus-menerus muncul di daftar ini tanpa makna.
  //   2. Cabang yang %FU-nya SUDAH 100% — semua leads-nya sudah di-follow-up
  //      tuntas, jadi wajar DtD-nya 0 terus (memang sudah tidak ada sisa
  //      leads buat di-follow-up lagi). Ini bukan cabang yang "belum
  //      progress", justru sudah selesai — kalau ikut di-mention malah
  //      salah kaprah/nyalahin cabang yang udah kerja tuntas.
  var belumProgress = a.cabang.filter(function (c) { return c.dtd === 0 && c.jumlahLeads > 0 && c.persenFU < 100; })
    .sort(function (x, y) { return x.persenFU - y.persenFU; });

  if (belumProgress.length > 0) {
    lines.push('Cabang yang belum ada progress FU (DtD = 0):');
    belumProgress.forEach(function (c, idx) {
      lines.push((idx + 1) + '. ' + labelCabangSingkat(c.nama) + ' – ' + formatAngkaID(c.jmlFU)
        + '/' + formatAngkaID(c.jumlahLeads) + ' (' + formatPersenID(c.persenFU) + '%)');
    });
    var top3 = belumProgress.slice(0, 3).map(function (c) { return labelCabangSingkat(c.nama); }).join(', ');
    lines.push('Mohon bantuan rekan-rekan cabang di atas untuk segera melanjutkan FU ke leads yang masih tersisa, khususnya di cabang dengan %FU rendah (' + top3 + ').');
  } else {
    lines.push('Seluruh cabang di Area ini sudah menunjukkan progress FU hari ini. Terima kasih atas kerja samanya, mohon dipertahankan.');
  }

  lines.push('Demikian, terima kasih.');
  return lines.join('\n');
}

// ============================================================
// MONITORING LIVIN' FOOD — Sheet baru otomatis tiap minggu
// ============================================================

// Data merchant tetap (urutan = urutan baris di sheet)
var LIVIN_MERCHANTS = [
  { region: 'REGION X/SULAWESI & MALUKU', nama: 'Coto Dg Tayang', outlet: 'Coto Daeng Tayang' },
  { region: 'REGION X/SULAWESI & MALUKU', nama: 'Dari kopi',      outlet: 'Dari kopi'          },
  { region: 'REGION X/SULAWESI & MALUKU', nama: 'Eksposed',       outlet: 'EKSPOSED SIGNATURE' },
  { region: 'REGION X/SULAWESI & MALUKU', nama: 'Konijiwa',       outlet: 'Konijiwa'           },
  { region: 'REGION X/SULAWESI & MALUKU', nama: 'Mas Daeng',      outlet: 'Mas Daeng Kuliner'  }
];

// Peta hari (JS getDay()) → offset kolom dari G (col 7)
// Setiap hari = 2 kolom: Kode Cabang & #order
// 0=Minggu→6, 1=Senin→0, 2=Selasa→1, ..., 6=Sabtu→5
var HARI_OFFSET = { 1:0, 2:1, 3:2, 4:3, 5:4, 6:5, 0:6 };

/**
 * Hitung nomor minggu ISO dari sebuah tanggal.
 * Minggu ISO dimulai Senin, minggu pertama = minggu yang mengandung Kamis pertama tahun tsb.
 */
function getISOWeek(date) {
  var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var day = d.getUTCDay() || 7; // jadikan Minggu = 7
  d.setUTCDate(d.getUTCDate() + 4 - day);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return { week: Math.ceil(((d - yearStart) / 86400000 + 1) / 7), year: d.getUTCFullYear() };
}

/**
 * Dapatkan tanggal Senin dan Minggu dari sebuah tanggal (untuk label header sheet).
 */
function getRangeMinggu(date) {
  var d = new Date(date);
  var day = d.getDay() || 7; // Minggu = 7
  var senin = new Date(d); senin.setDate(d.getDate() - day + 1);
  var minggu = new Date(senin); minggu.setDate(senin.getDate() + 6);
  var bulan = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  return (senin.getDate() + ' ' + bulan[senin.getMonth()])
       + ' – '
       + (minggu.getDate() + ' ' + bulan[minggu.getMonth()] + ' ' + minggu.getFullYear());
}

/**
 * Update monitoring Livin' Food di sheet minggu yang sesuai.
 * Sheet dibuat otomatis jika belum ada untuk minggu tersebut.
 */
function updateMonitoringLivinFood(ss, tanggal, kodeCabang, namaMerchant, jumlahOrder) {
  var tglObj = new Date(tanggal);
  var wk     = getISOWeek(tglObj);
  var sheetName = "Livin' Food W" + wk.week + ' ' + wk.year;

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = buatSheetLivinMinggu(ss, sheetName, tglObj);

  // Tentukan kolom hari
  var dayJS    = tglObj.getDay();
  var offset   = HARI_OFFSET[dayJS];
  var colKode  = 7 + offset * 2;
  var colOrder = 7 + offset * 2 + 1;

  // Cari baris merchant (mulai baris 5 karena baris 1 = title, 2-3 = header, 4 = kosong/range)
  var DATA_START_ROW = 5;
  var merchantNorm = (namaMerchant || '').toLowerCase().trim();
  var targetRow = -1;

  for (var i = 0; i < LIVIN_MERCHANTS.length; i++) {
    if (LIVIN_MERCHANTS[i].nama.toLowerCase() === merchantNorm) {
      targetRow = DATA_START_ROW + i; break;
    }
    var sheetNama = sheet.getRange(DATA_START_ROW + i, 2).getValue();
    if (String(sheetNama).toLowerCase().trim() === merchantNorm) {
      targetRow = DATA_START_ROW + i; break;
    }
  }
  if (targetRow === -1) {
    Logger.log('Merchant tidak ditemukan: ' + namaMerchant); return;
  }

  // Akumulasi #order; gabung kode cabang jika berbeda
  var existingKode  = sheet.getRange(targetRow, colKode).getValue();
  var existingOrder = Number(sheet.getRange(targetRow, colOrder).getValue()) || 0;
  var newOrder = existingOrder + jumlahOrder;
  var newKode  = (!existingKode || existingKode === '')
    ? kodeCabang
    : (String(existingKode).indexOf(kodeCabang) === -1 ? existingKode + ', ' + kodeCabang : existingKode);

  sheet.getRange(targetRow, colKode).setValue(newKode);
  sheet.getRange(targetRow, colOrder).setValue(newOrder);
  Logger.log('Monitoring updated [' + sheetName + ']: ' + namaMerchant + ' day=' + dayJS + ' order=' + newOrder);
}

/**
 * Buat sheet monitoring baru untuk satu minggu.
 * Nama sheet: "Livin' Food W27 2026"
 */
function buatSheetLivinMinggu(ss, sheetName, tglObj) {
  var sheet    = ss.insertSheet(sheetName);
  var hdrFill  = '#002060';
  var hdrFont  = '#FFFFFF';
  var days     = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];
  var rangeMgg = getRangeMinggu(tglObj);

  // Baris 1: periode minggu (A1:F1) + judul Cabang Pemesan (G1:T1)
  sheet.getRange('A1:F1').merge();
  sheet.getRange('A1').setValue("Periode: " + rangeMgg);
  styleHeader(sheet.getRange('A1'), hdrFill, hdrFont, true, 'left');

  sheet.getRange('G1:T1').merge();
  sheet.getRange('G1').setValue('Cabang Pemesan (min order 10 per hari)');
  styleHeader(sheet.getRange('G1'), hdrFill, hdrFont, true, 'center');

  // Baris 2–3: header kolom tetap (merge rows 2:3 per kolom)
  var fixedHeaders = ['Region Merchant', 'Nama Merchant', 'Nama Outlet', 'Area',
                      'Nama PIC Area/ Jabatan', 'No HP  PIC Area'];
  for (var f = 0; f < fixedHeaders.length; f++) {
    var r = sheet.getRange(2, f + 1, 2, 1);
    r.merge(); r.setValue(fixedHeaders[f]);
    styleHeader(r, hdrFill, hdrFont, true, 'center');
  }

  // Baris 2–3: header hari
  for (var d = 0; d < days.length; d++) {
    var col = 7 + d * 2;
    var rng2 = sheet.getRange(2, col, 1, 2);
    rng2.merge(); rng2.setValue(days[d]);
    styleHeader(rng2, hdrFill, hdrFont, true, 'center');

    styleHeader(sheet.getRange(3, col),     hdrFill, hdrFont, true, 'center').setValue('Kode Cabang');
    styleHeader(sheet.getRange(3, col + 1), hdrFill, hdrFont, true, 'center').setValue('#order');
  }

  // Baris 4: kosong (spacer) — baris data merchant mulai baris 5
  sheet.setRowHeight(3, 32);

  // Lebar kolom
  var widths = [200, 130, 160, 85, 95, 95, 65, 55, 65, 55, 65, 55, 65, 55, 65, 55, 65, 55, 65, 55];
  for (var w = 0; w < widths.length; w++) {
    sheet.setColumnWidth(w + 1, widths[w]);
  }

  // Data rows merchant (mulai baris 5)
  for (var m = 0; m < LIVIN_MERCHANTS.length; m++) {
    var row = 5 + m;
    sheet.getRange(row, 1).setValue(LIVIN_MERCHANTS[m].region);
    sheet.getRange(row, 2).setValue(LIVIN_MERCHANTS[m].nama);
    sheet.getRange(row, 3).setValue(LIVIN_MERCHANTS[m].outlet);
  }

  // Border seluruh tabel (baris 1–(5+jumlah merchant-1), kolom A–T)
  sheet.getRange(1, 1, 4 + LIVIN_MERCHANTS.length, 20)
    .setBorder(true, true, true, true, true, true);

  sheet.setFrozenRows(3);
  Logger.log('Sheet baru dibuat: ' + sheetName + ' (' + rangeMgg + ')');
  return sheet;
}

function styleHeader(range, bgHex, fontHex, bold, halign) {
  range.setBackground(bgHex)
       .setFontColor(fontHex)
       .setFontWeight(bold ? 'bold' : 'normal')
       .setHorizontalAlignment(halign || 'center')
       .setVerticalAlignment('middle')
       .setWrap(true);
  return range;
}

/**
 * Test manual: buat sheet monitoring untuk minggu ini.
 * Jalankan dari Apps Script Editor untuk cek hasilnya.
 */
function testBuatSheetLivinMingguIni() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tgl = new Date();
  var wk  = getISOWeek(tgl);
  var sheetName = "Livin' Food W" + wk.week + ' ' + wk.year;
  var existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  buatSheetLivinMinggu(ss, sheetName, tgl);
  beriTahu('✅ Sheet "' + sheetName + '" berhasil dibuat!');
}

/**
 * Test: cek apakah API berjalan dengan benar.
 * Jalankan dari Apps Script Editor, lihat hasilnya di Logs (View > Logs).
 */
function testAPI() {
  Logger.log('URL Web App: ' + ScriptApp.getService().getUrl());
  Logger.log('Sheet: ' + CONFIG.NAMA_SHEET);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NAMA_SHEET);
  Logger.log('Sheet ada: ' + (sheet ? 'YA' : 'BELUM — jalankan setupSheet()'));
}


// ============================================================
// BAGIAN B — REKAP CABANG (menu & sidebar di Google Sheets)
// ============================================================

const SHEET_LAPORAN = "Laporan Harian";
const SHEET_MASTER  = "Master Cabang"; // opsional, lihat keterangan di bawah

// Nama kolom di sheet "Laporan Harian" — HARUS sama persis dengan tulisan
// header di baris 1 sheet tersebut. Kalau nanti ada kolom baru, cukup
// tambahkan barisnya di sini, tidak perlu ubah bagian lain.
const KOLOM = {
  TANGGAL     : "Tanggal Laporan",
  AREA        : "Area",
  CABANG      : "Nama Cabang",
  KODE        : "Kode Cabang",
  LVM         : "Jumlah Akuisisi LVM",
  EDC         : "Jumlah Akuisisi EDC",
  EDC_POT     : "Jumlah Akuisisi EDC POT",
  PASANG_LVM  : "Jumlah Pemasangan LVM",
  RETENSI_EDC : "Jumlah Retensi EDC",
  TOTAL       : "Total Akuisisi"
};

// Baca header baris 1 dari sebuah sheet, lalu cocokkan dengan KOLOM di atas.
// Hasilnya: object berisi index kolom (0-based). -1 berarti kolom tidak ditemukan.
function getColMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  for (const key in KOLOM) {
    map[key] = headers.indexOf(KOLOM[key]);
  }
  return map;
}

// Ambil angka dari sebuah baris dengan aman. Kalau kolomnya tidak ada
// (index -1, misal header belum dibuat di sheet), otomatis dianggap 0.
function ambilAngka(row, idx) {
  if (idx < 0) return 0;
  return Number(row[idx]) || 0;
}

// =====================================================================
// 1. MENU - otomatis muncul saat file dibuka
// =====================================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 Rekap Cabang")
    .addItem("Buka Panel Rekap", "bukaPanel")
    .addItem("Buat Sheet Rekap Hari Ini", "buatSheetRekapHariIni")
    .addSeparator()
    .addItem("🧮 Buat/Refresh Rekap Formula (COUNTIFS)", "buatRekapFormula")
    .addToUi();

  ui.createMenu("🔧 Maintenance")
    .addItem("🔴 Aktifkan Mode Maintenance", "aktifkanMaintenance")
    .addItem("🟢 Matikan Mode Maintenance", "matikanMaintenance")
    .addSeparator()
    .addItem("Cek Status Sekarang", "cekStatusMaintenanceUI")
    .addToUi();
}

// =====================================================================
// MODE MAINTENANCE — blokir sementara semua akses dari GitHub Pages
// (form/pipeline/laporan/monitoring) selagi database lagi dibenerin
// manual. Statusnya disimpan di Script Properties (bukan sheet), jadi
// gak nambah baris/kolom di spreadsheet. Toggle lewat menu "🔧
// Maintenance" di atas — TIDAK PERLU buka Apps Script Editor.
// =====================================================================
function isMaintenanceAktif() {
  return PropertiesService.getScriptProperties().getProperty('MAINTENANCE_MODE') === 'ON';
}

function aktifkanMaintenance() {
  PropertiesService.getScriptProperties().setProperty('MAINTENANCE_MODE', 'ON');
  SpreadsheetApp.getUi().alert(
    '🔴 Mode Maintenance AKTIF.\n\n' +
    'Form pelaporan, Pipeline, dan Monitoring Area sekarang nampilin halaman ' +
    '"Sedang Pemeliharaan" ke siapapun yang buka — aman buat kamu benerin data ' +
    'langsung di sheet. Jangan lupa matiin lagi lewat menu ini kalau udah selesai.'
  );
}

function matikanMaintenance() {
  PropertiesService.getScriptProperties().deleteProperty('MAINTENANCE_MODE');
  SpreadsheetApp.getUi().alert('🟢 Mode Maintenance DIMATIKAN. Semua halaman kembali normal.');
}

function cekStatusMaintenanceUI() {
  var aktif = isMaintenanceAktif();
  SpreadsheetApp.getUi().alert(aktif
    ? '🔴 Mode Maintenance sedang AKTIF. Semua halaman GitHub Pages lagi diblokir.'
    : '🟢 Mode Maintenance sedang NONAKTIF. Semua halaman berjalan normal.');
}

// =====================================================================
// 2. BUKA SIDEBAR
// =====================================================================
function bukaPanel() {
  const html = HtmlService.createHtmlOutput(getHtmlSidebar())
    .setTitle("Rekap Laporan Cabang")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// =====================================================================
// 3. AMBIL DAFTAR TANGGAL (dipanggil dari HTML)
// =====================================================================
function getDaftarTanggal() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LAPORAN);
  if (!sheet) return [];

  const col         = getColMap(sheet);
  const data        = sheet.getDataRange().getValues();
  const tanggalSet  = new Set();

  for (let i = 1; i < data.length; i++) {
    const tgl = data[i][col.TANGGAL];
    if (tgl instanceof Date && !isNaN(tgl)) {
      tanggalSet.add(formatTanggalISO(tgl));
    }
  }

  // Urutkan dari terbaru
  return Array.from(tanggalSet).sort((a, b) => b.localeCompare(a));
}

// =====================================================================
// 4. AMBIL SEMUA CABANG DARI MASTER CABANG
// Prioritas utama: sheet "Master Cabang"
// Fallback: ambil dari data unik di "Laporan Harian" (jika master kosong)
// =====================================================================
function getAllCabang() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const master      = {};
  const sheetMaster = ss.getSheetByName(SHEET_MASTER);

  // ✅ Prioritas utama: baca dari sheet Master Cabang
  if (sheetMaster && sheetMaster.getLastRow() > 1) {
    const dataMaster = sheetMaster.getDataRange().getValues();
    for (let i = 1; i < dataMaster.length; i++) {
      const kode = String(dataMaster[i][0]).trim();
      const nama = String(dataMaster[i][1]).trim();
      const area = String(dataMaster[i][2] || "-").trim();
      if (kode && nama && kode !== "" && nama !== "") {
        master[kode] = { nama: nama, area: area };
      }
    }
    return master;
  }

  // ⚠️ Fallback: ambil cabang unik dari Laporan Harian jika master belum diisi
  const sheetLaporan = ss.getSheetByName(SHEET_LAPORAN);
  if (!sheetLaporan) return {};

  const col  = getColMap(sheetLaporan);
  const data = sheetLaporan.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const kode = String(data[i][col.KODE]).trim();
    const nama = String(data[i][col.CABANG]).trim();
    const area = String(data[i][col.AREA] || "-").trim();
    if (kode && nama) {
      master[kode] = { nama: nama, area: area };
    }
  }

  return master;
}

// =====================================================================
// 5. REKAP BERDASARKAN TANGGAL (dipanggil dari HTML)
// =====================================================================
function getRekap(tanggalStr) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LAPORAN);
  if (!sheet) return { sudahLapor: [], belumLapor: [], error: "Sheet tidak ditemukan!" };

  const col         = getColMap(sheet);
  const data        = sheet.getDataRange().getValues();
  const allCabang   = getAllCabang();
  const sudahLapor  = {};

  // Cari cabang yang sudah lapor di tanggal tsb
  for (let i = 1; i < data.length; i++) {
    const tgl  = data[i][col.TANGGAL];
    const kode = data[i][col.KODE];
    const nama = data[i][col.CABANG];
    const area = data[i][col.AREA];

    if (tgl instanceof Date && formatTanggalISO(tgl) === tanggalStr && kode) {
      if (!sudahLapor[kode]) {
        sudahLapor[kode] = {
          nama          : nama,
          area          : area || "-",
          jumlahLaporan : 0,
          lvm           : 0,
          edc           : 0,
          edcPot        : 0,
          pasangLvm     : 0,
          retensiEdc    : 0,
          total         : 0
        };
      }
      sudahLapor[kode].jumlahLaporan++;
      sudahLapor[kode].lvm        += ambilAngka(data[i], col.LVM);
      sudahLapor[kode].edc        += ambilAngka(data[i], col.EDC);
      sudahLapor[kode].edcPot     += ambilAngka(data[i], col.EDC_POT);
      sudahLapor[kode].pasangLvm  += ambilAngka(data[i], col.PASANG_LVM);
      sudahLapor[kode].retensiEdc += ambilAngka(data[i], col.RETENSI_EDC);
      sudahLapor[kode].total      += ambilAngka(data[i], col.TOTAL);
    }
  }

  // Cabang yang belum lapor
  const belumLapor = [];
  for (const kode in allCabang) {
    if (!sudahLapor[kode]) {
      belumLapor.push({
        kode : kode,
        nama : allCabang[kode].nama,
        area : allCabang[kode].area
      });
    }
  }

  // Format hasil sudah lapor
  const sudahLaporArr = Object.entries(sudahLapor).map(([kode, info]) => ({
    kode          : kode,
    nama          : info.nama,
    area          : info.area,
    jumlahLaporan : info.jumlahLaporan,
    lvm           : info.lvm,
    edc           : info.edc,
    edcPot        : info.edcPot,
    pasangLvm     : info.pasangLvm,
    retensiEdc    : info.retensiEdc,
    total         : info.total
  }));

  // Urutkan berdasarkan area lalu nama
  const sortByAreaNama = (a, b) => a.area.localeCompare(b.area) || a.nama.localeCompare(b.nama);
  sudahLaporArr.sort(sortByAreaNama);
  belumLapor.sort(sortByAreaNama);

  return {
    tanggal      : tanggalStr,
    totalCabang  : Object.keys(allCabang).length,
    sudahLapor   : sudahLaporArr,
    belumLapor   : belumLapor
  };
}

// =====================================================================
// 6. BUAT SHEET REKAP (tombol ekspor)
// =====================================================================
function buatSheetRekapHariIni() {
  const tanggal = formatTanggalISO(new Date());
  buatSheetRekap(tanggal);
}

function buatSheetRekap(tanggalStr) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const rekap  = getRekap(tanggalStr);
  const namaSheet = "Rekap " + tanggalStr;

  // Hapus sheet lama jika sudah ada
  const existing = ss.getSheetByName(namaSheet);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(namaSheet);

  // Header utama
  sheet.getRange("A1").setValue("REKAP LAPORAN CABANG");
  sheet.getRange("A2").setValue("Tanggal: " + tanggalStr);
  sheet.getRange("A3").setValue(
    `Total: ${rekap.sudahLapor.length} sudah lapor, ${rekap.belumLapor.length} belum lapor dari ${rekap.totalCabang} cabang`
  );

  // ---- Tabel SUDAH LAPOR ----
  const headerSudah = ["No", "Kode", "Area", "Nama Cabang", "Akuisisi LVM", "Akuisisi EDC", "Akuisisi EDC POT", "Pemasangan LVM", "Retensi EDC", "Total Akuisisi", "Jml Laporan"];
  const JML_KOLOM = headerSudah.length;
  sheet.getRange(5, 1, 1, JML_KOLOM).setValues([headerSudah]);
  sheet.getRange(5, 1, 1, JML_KOLOM)
    .setBackground("#34a853").setFontColor("white").setFontWeight("bold");

  // Susun semua baris jadi satu array dulu, baru ditulis SEKALI lewat
  // setValues() — nulis satu-satu per baris via getRange().setValues()
  // di dalam loop bikin ratusan API call terpisah & bisa timeout (6 menit)
  // kalau jumlah cabang banyak.
  if (rekap.sudahLapor.length > 0) {
    const dataSudah = rekap.sudahLapor.map((c, i) => [
      i + 1, c.kode, c.area, c.nama,
      c.lvm, c.edc, c.edcPot, c.pasangLvm, c.retensiEdc, c.total,
      c.jumlahLaporan
    ]);
    sheet.getRange(6, 1, dataSudah.length, JML_KOLOM).setValues(dataSudah);

    const bgSudah = rekap.sudahLapor.map((c, i) =>
      Array(JML_KOLOM).fill(i % 2 === 0 ? "#d9ead3" : null)
    );
    sheet.getRange(6, 1, bgSudah.length, JML_KOLOM).setBackgrounds(bgSudah);
  }

  // Baris total
  const totalRow = 6 + rekap.sudahLapor.length;
  const totals = rekap.sudahLapor.reduce((acc, c) => {
    acc.lvm        += c.lvm;
    acc.edc        += c.edc;
    acc.edcPot     += c.edcPot;
    acc.pasangLvm  += c.pasangLvm;
    acc.retensiEdc += c.retensiEdc;
    acc.total      += c.total;
    return acc;
  }, { lvm: 0, edc: 0, edcPot: 0, pasangLvm: 0, retensiEdc: 0, total: 0 });

  sheet.getRange(totalRow, 1, 1, JML_KOLOM).setValues([[
    "", "", "", "TOTAL",
    totals.lvm, totals.edc, totals.edcPot, totals.pasangLvm, totals.retensiEdc, totals.total, ""
  ]]);
  sheet.getRange(totalRow, 1, 1, JML_KOLOM).setBackground("#b6d7a8").setFontWeight("bold");

  // ---- Tabel BELUM LAPOR ----
  const startRow = 6 + rekap.sudahLapor.length + 3;
  sheet.getRange(startRow, 1).setValue("BELUM LAPOR");
  const headerBelum = ["No", "Kode", "Area", "Nama Cabang"];
  sheet.getRange(startRow + 1, 1, 1, headerBelum.length).setValues([headerBelum]);
  sheet.getRange(startRow + 1, 1, 1, headerBelum.length)
    .setBackground("#ea4335").setFontColor("white").setFontWeight("bold");

  if (rekap.belumLapor.length > 0) {
    const dataBelum = rekap.belumLapor.map((c, i) => [i + 1, c.kode, c.area, c.nama]);
    sheet.getRange(startRow + 2, 1, dataBelum.length, 4).setValues(dataBelum);

    const bgBelum = rekap.belumLapor.map((c, i) =>
      Array(4).fill(i % 2 === 0 ? "#fce8e6" : null)
    );
    sheet.getRange(startRow + 2, 1, bgBelum.length, 4).setBackgrounds(bgBelum);
  }

  sheet.autoResizeColumns(1, JML_KOLOM);
  SpreadsheetApp.getUi().alert(`Sheet rekap "${namaSheet}" berhasil dibuat!`);
  return namaSheet;
}

// =====================================================================
// 7. HELPER (BAGIAN B)
// Catatan: nama fungsi ini "formatTanggalISO" (bukan "formatTanggal")
// karena nama "formatTanggal" sudah dipakai di Bagian A untuk keperluan
// berbeda (format dd/MM/yyyy untuk teks WA & tampilan laporan).
// =====================================================================
function formatTanggalISO(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

// =====================================================================
// 8. HTML SIDEBAR
// =====================================================================
function getHtmlSidebar() {
  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, sans-serif; font-size: 13px; background: #f8f9fa; }

  .header {
    background: linear-gradient(135deg, #1a73e8, #0d47a1);
    color: white; padding: 14px 16px; text-align: center;
  }
  .header h2 { font-size: 15px; margin-bottom: 3px; }
  .header p  { font-size: 11px; opacity: 0.85; }

  .content { padding: 14px; }

  .form-group { margin-bottom: 12px; }
  label { display: block; font-weight: bold; margin-bottom: 5px; color: #333; }
  select {
    width: 100%; padding: 8px 10px; border: 1px solid #ccc;
    border-radius: 6px; font-size: 13px; background: white;
  }

  .btn {
    width: 100%; padding: 9px; border: none; border-radius: 6px;
    font-size: 13px; font-weight: bold; cursor: pointer; margin-bottom: 6px;
  }
  .btn-primary { background: #1a73e8; color: white; }
  .btn-primary:hover { background: #1558b0; }
  .btn-export  { background: #34a853; color: white; }
  .btn-export:hover { background: #2d8f47; }

  .summary {
    display: flex; gap: 8px; margin-bottom: 14px;
  }
  .stat-box {
    flex: 1; text-align: center; padding: 10px 6px;
    border-radius: 8px; font-weight: bold;
  }
  .stat-box .num  { font-size: 22px; }
  .stat-box .lbl  { font-size: 10px; margin-top: 2px; }
  .stat-sudah { background: #e6f4ea; color: #1e7e34; border: 1px solid #b7dfbc; }
  .stat-belum { background: #fce8e6; color: #c0392b; border: 1px solid #f5b7b1; }
  .stat-total { background: #e8f0fe; color: #1a73e8; border: 1px solid #aecbfa; }

  .tab-header { display: flex; margin-bottom: 8px; border-bottom: 2px solid #ddd; }
  .tab-btn {
    flex: 1; padding: 8px; background: none; border: none;
    cursor: pointer; font-size: 12px; font-weight: bold; color: #666;
  }
  .tab-btn.active { color: #1a73e8; border-bottom: 3px solid #1a73e8; margin-bottom: -2px; }

  .tab-content { display: none; }
  .tab-content.active { display: block; }

  .search-box {
    width: 100%; padding: 6px 10px; margin-bottom: 8px;
    border: 1px solid #ccc; border-radius: 6px; font-size: 12px;
  }

  .list-item {
    padding: 7px 10px; border-radius: 6px; margin-bottom: 4px;
    border-left: 4px solid;
  }
  .item-sudah { background: #f0faf2; border-color: #34a853; }
  .item-belum { background: #fff5f5; border-color: #ea4335; }
  .item-nama  { font-weight: bold; font-size: 12px; color: #333; }
  .item-info  { font-size: 11px; color: #666; margin-top: 2px; }
  .item-stats { display: flex; flex-wrap: wrap; gap: 4px; margin-top: 5px; }
  .stat-pill  {
    font-size: 10px; padding: 2px 6px; border-radius: 10px;
    background: #e8f5e9; color: #2e7d32; border: 1px solid #c8e6c9;
  }
  .stat-pill.total { background: #e3f2fd; color: #1565c0; border-color: #bbdefb; font-weight: bold; }
  .badge {
    float: right; font-size: 10px; padding: 2px 6px;
    border-radius: 10px; margin-top: 1px;
  }
  .badge-sudah { background: #34a853; color: white; }
  .badge-belum { background: #ea4335; color: white; }

  .loading { text-align: center; padding: 30px; color: #999; }
  .empty   { text-align: center; padding: 20px; color: #999; font-style: italic; }
  .error   { background: #fce8e6; color: #c0392b; padding: 10px; border-radius: 6px; font-size: 12px; }
</style>
</head>
<body>

<div class="header">
  <h2>📊 Rekap Laporan Cabang</h2>
  <p>Pantau cabang yang sudah & belum lapor</p>
</div>

<div class="content">

  <div class="form-group">
    <label>📅 Pilih Tanggal</label>
    <select id="selTanggal">
      <option value="">-- Memuat tanggal... --</option>
    </select>
  </div>

  <button class="btn btn-primary" onclick="cariRekap()">🔍 Tampilkan Rekap</button>
  <button class="btn btn-export" onclick="eksporSheet()" id="btnEkspor" style="display:none">
    📄 Buat Sheet Rekap
  </button>

  <div id="hasil"></div>

</div>

<script>
  let hasilGlobal = null;

  // Load daftar tanggal saat halaman siap
  window.onload = function() {
    google.script.run
      .withSuccessHandler(function(tglList) {
        const sel = document.getElementById('selTanggal');
        sel.innerHTML = '<option value="">-- Pilih Tanggal --</option>';
        tglList.forEach(t => {
          const opt = document.createElement('option');
          opt.value = t;
          opt.textContent = formatTampil(t);
          sel.appendChild(opt);
        });
        // Auto pilih tanggal terbaru
        if (tglList.length > 0) {
          sel.value = tglList[0];
          cariRekap();
        }
      })
      .withFailureHandler(showError)
      .getDaftarTanggal();
  };

  function cariRekap() {
    const tgl = document.getElementById('selTanggal').value;
    if (!tgl) { alert('Pilih tanggal dulu ya!'); return; }

    document.getElementById('hasil').innerHTML =
      '<div class="loading">⏳ Sedang memuat data...</div>';
    document.getElementById('btnEkspor').style.display = 'none';

    google.script.run
      .withSuccessHandler(tampilkanHasil)
      .withFailureHandler(showError)
      .getRekap(tgl);
  }

  function tampilkanHasil(data) {
    hasilGlobal = data;
    const sudah = data.sudahLapor;
    const belum = data.belumLapor;

    let html = \`
      <div class="summary">
        <div class="stat-box stat-total">
          <div class="num">\${data.totalCabang}</div>
          <div class="lbl">Total Cabang</div>
        </div>
        <div class="stat-box stat-sudah">
          <div class="num">\${sudah.length}</div>
          <div class="lbl">Sudah Lapor</div>
        </div>
        <div class="stat-box stat-belum">
          <div class="num">\${belum.length}</div>
          <div class="lbl">Belum Lapor</div>
        </div>
      </div>

      <div class="tab-header">
        <button class="tab-btn active" onclick="switchTab('sudah', this)">
          ✅ Sudah Lapor (\${sudah.length})
        </button>
        <button class="tab-btn" onclick="switchTab('belum', this)">
          ❌ Belum Lapor (\${belum.length})
        </button>
      </div>

      <input class="search-box" type="text" placeholder="🔎 Cari nama cabang..." onkeyup="filterList(this.value)" />

      <div id="tab-sudah" class="tab-content active">
    \`;

    if (sudah.length === 0) {
      html += '<div class="empty">Tidak ada cabang yang lapor di tanggal ini.</div>';
    } else {
      sudah.forEach((c, i) => {
        const dup = c.jumlahLaporan > 1 ? \` (\${c.jumlahLaporan}x)\` : '';
        html += \`
          <div class="list-item item-sudah" data-nama="\${c.nama.toLowerCase()}" data-area="\${c.area.toLowerCase()}">
            <span class="badge badge-sudah">✓\${dup}</span>
            <div class="item-nama">\${c.nama}</div>
            <div class="item-info">\${c.area} &bull; Kode: \${c.kode}</div>
            <div class="item-stats">
              <span class="stat-pill">LVM: \${c.lvm}</span>
              <span class="stat-pill">EDC: \${c.edc}</span>
              <span class="stat-pill">EDC POT: \${c.edcPot}</span>
              <span class="stat-pill">Pasang: \${c.pasangLvm}</span>
              <span class="stat-pill">Retensi: \${c.retensiEdc}</span>
              <span class="stat-pill total">Total: \${c.total}</span>
            </div>
          </div>\`;
      });
    }

    html += '</div><div id="tab-belum" class="tab-content">';

    if (belum.length === 0) {
      html += '<div class="empty">🎉 Semua cabang sudah lapor!</div>';
    } else {
      belum.forEach(c => {
        html += \`
          <div class="list-item item-belum" data-nama="\${c.nama.toLowerCase()}" data-area="\${c.area.toLowerCase()}">
            <span class="badge badge-belum">✗</span>
            <div class="item-nama">\${c.nama}</div>
            <div class="item-info">\${c.area} &bull; Kode: \${c.kode}</div>
          </div>\`;
      });
    }

    html += '</div>';

    document.getElementById('hasil').innerHTML = html;
    document.getElementById('btnEkspor').style.display = 'block';
  }

  function switchTab(tab, btn) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tab-' + tab).classList.add('active');
    document.querySelector('.search-box').value = '';
    filterList('');
  }

  function filterList(keyword) {
    const kw = keyword.toLowerCase();
    document.querySelectorAll('.tab-content.active .list-item').forEach(el => {
      const cocok = el.dataset.nama.includes(kw) || el.dataset.area.includes(kw);
      el.style.display = cocok ? 'block' : 'none';
    });
  }

  function eksporSheet() {
    const tgl = document.getElementById('selTanggal').value;
    document.getElementById('btnEkspor').textContent = '⏳ Membuat sheet...';
    google.script.run
      .withSuccessHandler(function(nama) {
        document.getElementById('btnEkspor').textContent = '📄 Buat Sheet Rekap';
        alert('Sheet "' + nama + '" berhasil dibuat!');
      })
      .withFailureHandler(showError)
      .buatSheetRekap(tgl);
  }

  function showError(err) {
    document.getElementById('hasil').innerHTML =
      '<div class="error">⚠️ Error: ' + err.message + '</div>';
  }

  function formatTampil(tglStr) {
    const [y, m, d] = tglStr.split('-');
    const bulan = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
    return d + ' ' + bulan[parseInt(m)-1] + ' ' + y;
  }
</script>
</body>
</html>
`;
}
