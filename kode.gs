/**
 * =========================================================================
 * ACRU EMPLOYEE ANALYTICS SYSTEM - KABUPATEN TOBA
 * File: Code.gs
 * Deskripsi: Web App Core Router, Modular Evaluator, Safe Database Resolver,
 * dan Global Column Index Resolver (Bebas dari duplikasi fungsi Service).
 * =========================================================================
 */

function doGet() {
  try {
    ensureDatabaseReady();
  } catch (e) {
    Logger.log('Auto-init database skipped/info: ' + e.toString());
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Dashboard Analytics System - Kabupaten Toba')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Helper Evaluasi Modular (Mendukung Recursive Scriptlet <?!= include(...) ?>)
 * Dilengkapi pencarian kandidat nama cerdas (.html, case-insensitive) dan alert informatif
 */
function include(filename) {
  var clean = String(filename || '').trim().replace(/\.html$/i, '');
  var candidates = [
    filename,
    clean,
    clean + '.html',
    clean + '.html.html',
    clean.toLowerCase(),
    clean.toLowerCase() + '.html',
    filename.toUpperCase(),
    clean.toUpperCase()
  ];

  for (var i = 0; i < candidates.length; i++) {
    var c = candidates[i];
    try {
      return HtmlService.createTemplateFromFile(c).evaluate().getContent();
    } catch (e1) {
      try {
        return HtmlService.createHtmlOutputFromFile(c).getContent();
      } catch (e2) {}
    }
  }

  Logger.log("Gagal menyertakan modul: " + filename);
  return "<div style='margin: 3rem auto; max-width: 650px; padding: 2rem; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 16px; text-align: center; box-shadow: 0 10px 25px -5px rgba(0,0,0,0.08); font-family: -apple-system, BlinkMacSystemFont, sans-serif;'>" +
    "<div style='font-size: 2.5rem; margin-bottom: 0.75rem;'>📄</div>" +
    "<h3 style='margin: 0 0 0.5rem 0; font-size: 1.15rem; font-weight: 800; color: #0f172a;'>Berkas HTML Belum Ditambahkan di Apps Script</h3>" +
    "<p style='font-size: 0.85rem; color: #64748b; margin-bottom: 1.25rem;'>Modul <b>&lt;" + filename + "&gt;</b> belum ada di daftar file project Apps Script Anda sehingga konten belum tampil.</p>" +
    "<div style='background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 10px; padding: 1rem; text-align: left; font-size: 0.825rem; color: #334155; line-height: 1.7;'>" +
      "<b>Cara Menampilkannya (2 Langkah Mudah):</b><br>" +
      "1. Buka Editor Google Apps Script (tab sebelah kiri).<br>" +
      "2. Klik tanda <b>+ (Tambah file)</b> &rarr; Pilih <b>HTML</b> &rarr; ketik nama persis: <code>" + clean + "</code>.<br>" +
      "3. Tempelkan seluruh kode dari berkas <code>" + clean + ".html</code> lalu tekan <b>Ctrl + S</b> dan segarkan halaman ini." +
    "</div>" +
  "</div>";
}

function initDatabaseFromUI() {
  try {
    if (typeof initialDatabaseSetup === 'function') {
      initialDatabaseSetup();
      return { success: true, message: "Struktur database berhasil disinkronkan!" };
    }
    return { success: true, message: "Database siap digunakan." };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Verifikasi PIN Keamanan Mode Admin BKPSDM
 */
function verifyAdminPin(pin) {
  var cleanPin = String(pin || '').trim();
  var allowedPins = ['bkpsdm2026', 'toba2026', 'admin2026'];
  if (allowedPins.indexOf(cleanPin) !== -1) {
    return { success: true, message: "Akses Mode Admin BKPSDM berhasil diaktifkan." };
  }
  return { success: false, message: "PIN Keamanan yang Anda masukkan salah. Silakan coba lagi." };
}

/**
 * Normalisasi Kunci Pencocokan String & Akronim OPD Cerdas
 */
function cleanMatchKey(str) {
  if (!str) return '';
  var s = String(str).toUpperCase().trim();
  s = s.replace(/&/g, ' DAN ');
  s = s.replace(/\s+/g, ' ');

  // Normalisasi Akronim OPD
  s = s.replace(/\bSDM\b/g, 'SUMBER DAYA MANUSIA');
  s = s.replace(/\bBKPSDM\b/g, 'BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA');
  s = s.replace(/\bBAPPERIDA\b/g, 'BADAN PERENCANAAN PEMBANGUNAN RISET DAN INOVASI DAERAH');
  s = s.replace(/\bBAPPEDA\b/g, 'BADAN PERENCANAAN PEMBANGUNAN RISET DAN INOVASI DAERAH');
  s = s.replace(/\bBAPELITBANGDA\b/g, 'BADAN PERENCANAAN PEMBANGUNAN RISET DAN INOVASI DAERAH');
  s = s.replace(/\bDINKES\b/g, 'DINAS KESEHATAN');
  s = s.replace(/\bDISDIKPORA\b/g, 'DINAS PENDIDIKAN PEMUDA DAN OLAHRAGA');
  s = s.replace(/\bSETDA\b|\bSETDAKAB\b/g, 'SEKRETARIAT DAERAH');
  s = s.replace(/\bSETWAN\b/g, 'SEKRETARIAT DEWAN');
  s = s.replace(/\bKESRA\b/g, 'KESEJAHTERAAN RAKYAT');
  s = s.replace(/\bEKOBANG\b/g, 'PEREKONOMIAN DAN PEMBANGUNAN');
  s = s.replace(/\bSEKCAM\b|\bSEKRETARIS\s+KECAMATAN\b/g, 'SEKRETARIS CAMAT');
  s = s.replace(/\bSEKDA\b|\bSEKDAKAB\b|\bSEKRETARIS\s+DAERAH\s+KABUPATEN\b/g, 'SEKRETARIS DAERAH');
  s = s.replace(/\bINSPEKTUR\s+INSPEKTORAT\b/g, 'INSPEKTUR');
  s = s.replace(/\bUPTD\b/g, 'UPT');
  s = s.replace(/\bTATA\s+USAHA\b|\bTU\b/g, 'TATA USAHA');

  // Normalisasi Asisten Setda Toba
  if (s.indexOf('ASISTEN') !== -1) {
    s = s.replace(/\bASISTEN\s+(SEKDA|SEKRETARIS\s+DAERAH)\b/g, 'ASISTEN');
    s = s.replace(/\bASISTEN\s+BIDANG\b/g, 'ASISTEN');
    if (s.indexOf('PEMERINTAH') !== -1 || s.indexOf('KESRA') !== -1 || s.indexOf('KESEJAHTERAAN') !== -1 || /\bASISTEN\s*(I|1)\b/.test(s)) {
      return 'ASISTENPEMERINTAHANDANKESEJAHTERAANRAKYAT';
    }
    if (s.indexOf('PEREKONOMIAN') !== -1 || s.indexOf('PEMBANGUNAN') !== -1 || s.indexOf('EKOBANG') !== -1 || /\bASISTEN\s*(II|2)\b/.test(s)) {
      return 'ASISTENPEREKONOMIANDANPEMBANGUNAN';
    }
    if (s.indexOf('ADMINISTRASI') !== -1 || s.indexOf('UMUM') !== -1 || /\bASISTEN\s*(III|3)\b/.test(s)) {
      return 'ASISTENADMINISTRASIUMUM';
    }
  }

  // Normalisasi Eselon III dan IV:
  s = s.replace(/\bPERKIM\b|\bPERUMAHAN\s+DAN\s+PERMUKIMAN\b/g, 'PERUMAHAN DAN KAWASAN PERMUKIMAN');
  s = s.replace(/\b(KEPALA\s+BIDANG|KABID)\s+PERUMAHAN\s+DAN\s+(KAWASAN\s+)?PERMUKIMAN\b/g, 'KEPALA BIDANG PERUMAHAN');
  s = s.replace(/\bKASUBBAG\b|\bKASUBAG\b|\bKEPALA\s+SUBBAGIAN\b|\bKEPALA\s+SUB\s+BAGIAN\b/g, 'KEPALA SUB BAGIAN');
  s = s.replace(/\bKASUBBID\b|\bKASUBID\b|\bKEPALA\s+SUBBIDANG\b|\bKEPALA\s+SUB\s+BIDANG\b/g, 'KEPALA SUB BIDANG');
  s = s.replace(/\bKABAG\b|\bKEPALA\s+BAGIAN\b/g, 'KEPALA BAGIAN');
  s = s.replace(/\bKABID\b|\bKEPALA\s+BIDANG\b/g, 'KEPALA BIDANG');
  s = s.replace(/\bKASI\b|\bKEPALA\s+SEKSI\b/g, 'KEPALA SEKSI');
  s = s.replace(/\bKADIS\b|\bKEPALA\s+DINAS\b/g, 'KEPALA DINAS');
  s = s.replace(/\bKABAN\b|\bKEPALA\s+BADAN\b/g, 'KEPALA BADAN');

  // Prepend KEPALA jika jabatan hanya tertulis 'SUB BAGIAN ...' / 'SUB BIDANG ...' / 'SEKSI ...'
  if (s.indexOf('SUB BAGIAN') === 0 || s.indexOf('SUBBAGIAN') === 0) {
    s = 'KEPALA ' + s;
  } else if (s.indexOf('SUB BIDANG') === 0 || s.indexOf('SUBBIDANG') === 0) {
    s = 'KEPALA ' + s;
  } else if (s.indexOf('SEKSI ') === 0) {
    s = 'KEPALA ' + s;
  }

  s = s.replace(/\bPADA\b/g, '');
  s = s.replace(/\bKABUPATEN(\s+TOBA)?\b/g, '');
  s = s.replace(/\bKAB\s+TOBA\b/g, '');
  s = s.replace(/\bPEMERINTAH\s+KABUPATEN(\s+TOBA)?\b/g, '');
  s = s.replace(/\bPEMKAB(\s+TOBA)?\b/g, '');
  return s.replace(/[^A-Z0-9]/g, '');
}

/**
 * Memisahkan dan membersihkan nama OPD dari jabatan
 * Contoh: "Sekretaris - Dinas Kesehatan" -> "Sekretaris"
 * Contoh: "Kepala Sub Bagian Umum dan Kepegawaian - Dinas Kesehatan" -> "Kepala Sub Bagian Umum dan Kepegawaian"
 */
function stripOpdFromJabatan(jabStr, opdStr) {
  if (!jabStr) return '';
  var s = String(jabStr).trim();
  if (opdStr) {
    var opdClean = String(opdStr).trim();
    var patterns = [
      opdClean,
      opdClean.replace(/Kabupaten.*$/i, '').trim(),
      opdClean.replace(/Daerah.*$/i, '').trim()
    ];
    for (var i = 0; i < patterns.length; i++) {
      var p = patterns[i];
      if (p && p.length >= 4) {
        var esc = p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        s = s.replace(new RegExp('(\\s*[-–/]\\s*|\\s+pada\\s+|\\s+)(?:dinas|badan|satuan|kecamatan|kantor|bagian)?\\s*' + esc + '.*$', 'i'), '').trim();
        s = s.replace(new RegExp('\\s*[-–/]\\s*' + esc + '.*$', 'i'), '').trim();
      }
    }
  }
  s = s.replace(/\s*[-–/]\s*(Dinas|Badan|Kecamatan|RSUD|Inspektorat|Sekretariat).+$/i, '').trim();
  s = s.replace(/\s+pada\s+.+$/i, '').trim();
  s = s.replace(/\s+(Kabupaten|Kab\.)\s+Toba.*$/i, '').trim();
  s = s.replace(/\s*[-–/]\s*$/g, '').trim();
  return s;
}


/**
 * Helper Pencari Sheet Tahan Banting:
 * Mengatasi trailing whitespace, perbedaan case, dan karakter simbol pada tab sheet.
 */
function getSheetByNameRobust(ss, targetName) {
  if (!ss || !targetName) return null;

  // 1. Pencocokan langsung
  var sheet = ss.getSheetByName(targetName);
  if (sheet) return sheet;

  // 2. Pencocokan dengan pemangkasan spasi
  var trimmed = String(targetName).trim();
  sheet = ss.getSheetByName(trimmed);
  if (sheet) return sheet;

  // 3. Pencocokan normalisasi karakter alfabet dan numerik
  var cleanTarget = trimmed.toUpperCase().replace(/[^A-Z0-9]/g, '');
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    var curSheet = allSheets[i];
    var cleanCurName = curSheet.getName().toUpperCase().trim().replace(/[^A-Z0-9]/g, '');
    if (cleanCurName === cleanTarget || cleanCurName.indexOf(cleanTarget) !== -1 || cleanTarget.indexOf(cleanCurName) !== -1) {
      return curSheet;
    }
  }

  return null;
}

/**
 * Helper Pencari Indeks Kolom Cerdas:
 * Mencegah benturan pencarian kolom kuota numerik dengan kolom kode ID.
 */
function findColumnIndex(headers, aliases) {
  if (!headers || !Array.isArray(headers) || !aliases || !Array.isArray(aliases)) return -1;
  
  var cleanHeaders = headers.map(function(h) { 
    return String(h || '').toUpperCase().trim().replace(/[^A-Z0-9]/g, '');
  });

  // 1. Pencocokan persis (Exact Clean Match)
  for (var a = 0; a < aliases.length; a++) {
    var cleanAlias = String(aliases[a] || '').toUpperCase().trim().replace(/[^A-Z0-9]/g, '');
    if (!cleanAlias) continue;
    var idx = cleanHeaders.indexOf(cleanAlias);
    if (idx !== -1) return idx;
  }

  // 2. Pencocokan substring dengan proteksi benturan kolom ID
  for (var a = 0; a < aliases.length; a++) {
    var cleanAlias = String(aliases[a] || '').toUpperCase().trim().replace(/[^A-Z0-9]/g, '');
    if (cleanAlias.length < 3) continue;
    
    for (var i = 0; i < cleanHeaders.length; i++) {
      var headerItem = cleanHeaders[i];
      if (!headerItem) continue;

      // Proteksi krusial: jangan cocokkan kolom ID saat mencari kuota jumlah
      if (cleanAlias.indexOf('ID') === -1) {
        if (headerItem.startsWith('ID') || headerItem.endsWith('ID') || headerItem.indexOf('IDKEBUTUHAN') !== -1 || headerItem.indexOf('IDJABATAN') !== -1) {
          continue;
        }
      }

      if (headerItem.indexOf(cleanAlias) !== -1 || cleanAlias.indexOf(headerItem) !== -1) {
        if (headerItem.length >= 3) return i;
      }
    }
  }

  return -1;
}