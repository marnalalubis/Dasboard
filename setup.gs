/**
 * ====================================================================
 * ACRU EMPLOYEE ANALYTICS SYSTEM - KABUPATEN TOBA
 * FILE: setup.gs
 * DESKRIPSI: Inisialisasi struktur sheet, header kolom, master data
 * kebutuhan formasi ABK (JF & JP), serta safe migration non-destruktif.
 * ====================================================================
 */

/**
 * Entry Point Global Pemastian Kesiapan Database (Safe Migrate)
 */
function ensureDatabaseReady() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheetPns = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, 'Data_Pegawai_PNS') : ss.getSheetByName('Data_Pegawai_PNS');
    if (!sheetPns) {
      if (typeof initialDatabaseSetup === 'function') {
        initialDatabaseSetup();
      } else if (typeof setupDatabase === 'function') {
        setupDatabase();
      }
    }
  } catch (err) {
    Logger.log('ensureDatabaseReady info: ' + err.toString());
  }
}

function initialDatabaseSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  // 1. Setup Sheet Pegawai PNS
  ensureSheetWithHeaders(ss, 'Data_Pegawai_PNS', [
    'PNS ID', 'NIP BARU', 'NAMA', 'GELAR DEPAN', 'GELAR BELAKANG', 'TEMPAT LAHIR',
    'TANGGAL LAHIR', 'JENIS KELAMIN', 'AGAMA NAMA', 'GOL AKHIR NAMA', 'TMT GOLONGAN AKHIR',
    'TINGKAT PENDIDIKAN NAMA', 'PENDIDIKAN TERAKHIR NAMA', 'TAHUN LULUS', 'JENIS JABATAN NAMA',
    'JABATAN NAMA', 'TMT JABATAN', 'ESELON NAMA', 'TMT ESELON', 'SATUAN KERJA INDUK NAMA',
    'SATUAN KERJA KERJA NAMA', 'UNOR NAMA', 'KEDUDUKAN HUKUM NAMA'
  ]);

  // 2. Setup Sheet Pegawai PPPK
  ensureSheetWithHeaders(ss, 'Data_Pegawai_PPPK', [
    'PNS ID', 'NIP BARU', 'NAMA', 'GELAR DEPAN', 'GELAR BELAKANG', 'TEMPAT LAHIR',
    'TANGGAL LAHIR', 'JENIS KELAMIN', 'AGAMA NAMA', 'GOL AKHIR NAMA', 'TMT GOLONGAN AKHIR',
    'TINGKAT PENDIDIKAN NAMA', 'PENDIDIKAN TERAKHIR NAMA', 'TAHUN LULUS', 'JENIS JABATAN NAMA',
    'JABATAN NAMA', 'TMT JABATAN', 'ESELON NAMA', 'TMT ESELON', 'SATUAN KERJA INDUK NAMA',
    'SATUAN KERJA KERJA NAMA', 'UNOR NAMA', 'KEDUDUKAN HUKUM NAMA'
  ]);

  // 3. Setup Sheet Pegawai PPPK Paruh Waktu
  ensureSheetWithHeaders(ss, 'Data_Pegawai_PPPK_Paruh_Waktu', [
    'PNS ID', 'NIP BARU', 'NAMA', 'GELAR DEPAN', 'GELAR BELAKANG', 'TEMPAT LAHIR',
    'TANGGAL LAHIR', 'JENIS KELAMIN', 'AGAMA NAMA', 'GOL AKHIR NAMA', 'TMT GOLONGAN AKHIR',
    'TINGKAT PENDIDIKAN NAMA', 'PENDIDIKAN TERAKHIR NAMA', 'TAHUN LULUS', 'JENIS JABATAN NAMA',
    'JABATAN NAMA', 'TMT JABATAN', 'ESELON NAMA', 'TMT ESELON', 'SATUAN KERJA INDUK NAMA',
    'SATUAN KERJA KERJA NAMA', 'UNOR NAMA', 'KEDUDUKAN HUKUM NAMA'
  ]);

  // 4. Setup Sheet Master Struktur Jabatan (Struktural Eselon II - IV)
  var sheetStruktur = ensureSheetWithHeaders(ss, 'Data_Master_Struktur_Jabatan', [
    'ID_JABATAN', 'PERANGKAT_DAERAH', 'UNOR_NAMA', 'JABATAN_NAMA', 'ESELON', 'JUMLAH_FORMASI', 'GOL_MINIMAL'
  ]);
  seedMasterStrukturIfEmpty(sheetStruktur);

  // 5. Setup Sheet Master Kebutuhan Formasi ABK Jabatan Fungsional (JF)
  var sheetKebutuhanJF = ensureSheetWithHeaders(ss, 'Data_Master_Kebutuhan_JF', [
    'ID_KEBUTUHAN', 'PERANGKAT_DAERAH', 'JABATAN_NAMA', 'JENJANG', 'JUMLAH_KEBUTUHAN', 'SUMBER_ANALISIS', 'KETERANGAN'
  ]);
  seedMasterKebutuhanJfIfEmpty(sheetKebutuhanJF);

  // 6. Setup Sheet Master Kebutuhan Formasi ABK Jabatan Pelaksana (JP)
  var sheetKebutuhanJP = ensureSheetWithHeaders(ss, 'Data_Master_Kebutuhan_JP', [
    'ID_KEBUTUHAN', 'PERANGKAT_DAERAH', 'JABATAN_NAMA', 'JUMLAH_KEBUTUHAN', 'SUMBER_ANALISIS', 'KETERANGAN'
  ]);
  seedMasterKebutuhanJpIfEmpty(sheetKebutuhanJP);

  // 7. Setup Sheet Log Import
  ensureSheetWithHeaders(ss, 'Log_Import', [
    'LOG ID', 'WAKTU IMPORT', 'STATUS AKTIVITAS'
  ]);

  SpreadsheetApp.flush();
  Logger.log('Inisialisasi Database & Master Kebutuhan ABK Berhasil Selesai.');
}

/**
 * Helper: Memastikan sheet tersedia dan memiliki struktur header yang benar (Safe Migrate)
 */
function ensureSheetWithHeaders(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0f172a')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Seed data awal master struktur jabatan jika kosong
 */
function seedMasterStrukturIfEmpty(sheet) {
  if (!sheet || sheet.getLastRow() > 1) return;

  var sampleData = [
    ['STR-001', 'Sekretariat Daerah Kabupaten', 'Sekretariat Daerah', 'Sekretaris Daerah', 'II.a', 1, 'IV/d'],
    ['STR-002', 'Sekretariat Daerah Kabupaten', 'Asisten Pemerintahan dan Kesejahteraan Rakyat', 'Asisten Pemerintahan dan Kesejahteraan Rakyat', 'II.b', 1, 'IV/b'],
    ['STR-003', 'Sekretariat Daerah Kabupaten', 'Asisten Perekonomian dan Pembangunan', 'Asisten Perekonomian dan Pembangunan', 'II.b', 1, 'IV/b'],
    ['STR-004', 'Sekretariat Daerah Kabupaten', 'Asisten Administrasi Umum', 'Asisten Administrasi Umum', 'II.b', 1, 'IV/b'],
    ['STR-005', 'Dinas Kesehatan', 'Dinas Kesehatan', 'Kepala Dinas Kesehatan', 'II.b', 1, 'IV/b'],
    ['STR-006', 'Dinas Kesehatan', 'Sekretariat Dinas Kesehatan', 'Sekretaris Dinas Kesehatan', 'III.a', 1, 'IV/a'],
    ['STR-007', 'Badan Kepegawaian dan Pengembangan Sumber Daya Manusia', 'BKPSDM', 'Kepala BKPSDM', 'II.b', 1, 'IV/b'],
    ['STR-008', 'Badan Kepegawaian dan Pengembangan Sumber Daya Manusia', 'Sekretariat BKPSDM', 'Sekretaris BKPSDM', 'III.a', 1, 'IV/a'],
    ['STR-009', 'Badan Perencanaan Pembangunan, Riset dan Inovasi Daerah', 'BAPPERIDA', 'Kepala BAPPERIDA', 'II.b', 1, 'IV/b']
  ];

  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
}

/**
 * Seed patokan kebutuhan awal ABK Jabatan Fungsional (JF) jika sheet masih baru
 */
function seedMasterKebutuhanJfIfEmpty(sheet) {
  if (!sheet || sheet.getLastRow() > 1) return;

  var sampleData = [
    // BKPSDM
    ['KB-JF-001', 'Badan Kepegawaian dan Pengembangan Sumber Daya Manusia', 'Analis Sumber Daya Manusia Aparatur Ahli Madya', 'Ahli Madya', 4, 'ABK BKPSDM', 'Manajemen Karir & SKP'],
    ['KB-JF-002', 'Badan Kepegawaian dan Pengembangan Sumber Daya Manusia', 'Analis Sumber Daya Manusia Aparatur Ahli Muda', 'Ahli Muda', 6, 'ABK BKPSDM', 'Pengadaan & Mutasi Pegawai'],
    ['KB-JF-003', 'Badan Kepegawaian dan Pengembangan Sumber Daya Manusia', 'Analis Sumber Daya Manusia Aparatur Ahli Pertama', 'Ahli Pertama', 6, 'ABK BKPSDM', 'Pelayanan Administrasi ASN'],
    ['KB-JF-004', 'Badan Kepegawaian dan Pengembangan Sumber Daya Manusia', 'Pranata Komputer Terampil', 'Terampil', 2, 'ABK BKPSDM', 'Pengelolaan Database SIMPEG'],
    // Bakesbangpol
    ['KB-JF-005', 'Badan Kesatuan Bangsa dan Politik', 'Analis Kebijakan Ahli Muda', 'Ahli Muda', 4, 'ABK Kesbangpol', 'Kewaspadaan Nasional & Ormas'],
    ['KB-JF-006', 'Badan Kesatuan Bangsa dan Politik', 'Perencana Ahli Muda', 'Ahli Muda', 2, 'ABK Kesbangpol', 'Penyusunan Program Kerja'],
    ['KB-JF-007', 'Badan Kesatuan Bangsa dan Politik', 'Perencana Ahli Pertama', 'Ahli Pertama', 2, 'ABK Kesbangpol', 'Evaluasi Kinerja Anggaran'],
    // BPBD
    ['KB-JF-008', 'Badan Penanggulangan Bencana Daerah', 'Pranata Pencarian dan Pertolongan Pemula', 'Pemula', 20, 'ABK BPBD', 'Kesiapsiagaan & Reaksi Cepat Bencana'],
    // BPKAD
    ['KB-JF-009', 'Badan Pengelolaan Keuangan dan Aset Daerah', 'Perencana Ahli Pertama', 'Ahli Pertama', 2, 'ABK BPKAD', 'Perencanaan Anggaran Keuangan'],
    ['KB-JF-010', 'Badan Pengelolaan Keuangan dan Aset Daerah', 'Pranata Komputer Terampil', 'Terampil', 3, 'ABK BPKAD', 'Pengelolaan Sistem Informasi Keuangan Daerah'],
    // Dinas Kesehatan & RSUD
    ['KB-JF-011', 'Dinas Kesehatan', 'Dokter Ahli Madya', 'Ahli Madya', 15, 'ABK Dinkes 2024', 'Dokter Spesialis & Puskesmas'],
    ['KB-JF-012', 'Dinas Kesehatan', 'Dokter Gigi Ahli Madya', 'Ahli Madya', 8, 'ABK Dinkes 2024', 'UPT Puskesmas se-Kabupaten'],
    ['KB-JF-013', 'RSUD Porsea', 'Dokter Ahli Madya', 'Ahli Madya', 12, 'Standar Akreditasi RS', 'Pelayanan Medik Spesialistik'],
    // Disdikpora
    ['KB-JF-014', 'Dinas Pendidikan, Pemuda dan Olahraga', 'Guru Ahli Pertama', 'Ahli Pertama', 1200, 'DAPODIK 2024', 'Kebutuhan Guru SD/SMP Negeri'],
    // Inspektorat & Bapperida
    ['KB-JF-015', 'Inspektorat Daerah', 'Auditor Ahli Pertama', 'Ahli Pertama', 10, 'ABK Pengawasan', 'Pengawasan Internal Pemda'],
    ['KB-JF-016', 'Badan Perencanaan Pembangunan, Riset dan Inovasi Daerah', 'Perencana Ahli Muda', 'Ahli Muda', 6, 'ABK BAPPERIDA', 'Rencana Pembangunan Daerah']
  ];

  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
}

/**
 * Seed patokan kebutuhan awal ABK Jabatan Pelaksana (JP) jika sheet masih baru
 */
function seedMasterKebutuhanJpIfEmpty(sheet) {
  if (!sheet || sheet.getLastRow() > 1) return;

  var sampleData = [
    ['KB-JP-001', 'Sekretariat Daerah Kabupaten', 'Pengadministrasi Perkantoran', 15, 'ABK Setda', 'Pelayanan Administrasi Umum'],
    ['KB-JP-002', 'Sekretariat Daerah Kabupaten', 'Pengemudi', 10, 'ABK Setda', 'Operasional Pimpinan Daerah'],
    ['KB-JP-003', 'Dinas Kesehatan', 'Pengadministrasi Perkantoran', 12, 'ABK Dinkes', 'Dinas & UPT Puskesmas'],
    ['KB-JP-004', 'Badan Pengelolaan Keuangan dan Aset Daerah', 'Pengadministrasi Keuangan', 8, 'ABK BPKAD', 'Verifikasi & Perbendaharaan'],
    ['KB-JP-005', 'Badan Kepegawaian dan Pengembangan Sumber Daya Manusia', 'Pengadministrasi Perkantoran', 6, 'ABK BKPSDM', 'Pelayanan Mutasi & Dokumen'],
    ['KB-JP-006', 'Satuan Polisi Pamong Praja', 'Polisi Pamong Praja Pemula', 45, 'ABK Satpol PP', 'Penegakan Perda & Ketertiban'],
    ['KB-JP-007', 'Dinas Komunikasi dan Informatika', 'Pengelola Sistem Informasi', 5, 'ABK Diskominfo', 'Layanan Jaringan & Server']
  ];

  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
}