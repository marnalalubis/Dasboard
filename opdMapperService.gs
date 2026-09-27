/**
 * =========================================================================
 * ACRU EMPLOYEE ANALYTICS SYSTEM - KABUPATEN TOBA
 * File: opdMapperService.gs
 * Deskripsi: Master dictionary 44 OPD Pemkab Toba, normalisasi string cerdas,
 * algoritma hierarki pemetaan Satker/UNOR, dan utility pencari kolom.
 * =========================================================================
 */



/**
 * Normalisasi cerdas nama OPD
 * Mengonversi variasi nama lama (BAPPEDA, BAPELITBANGDA, Penelitian dan Pengembangan)
 * menjadi nama resmi: "Badan Perencanaan Pembangunan, Riset dan Inovasi Daerah"
 */
function normalizeOpdName(rawName) {
  if (!rawName) return '';
  var clean = String(rawName).trim();
  var upper = clean.toUpperCase();

  // 1. Deteksi variasi nama Sekretariat Daerah (Setda / Setdakab)
  if (
    upper === "SEKRETARIAT DAERAH" ||
    upper === "SEKRETARIAT DAERAH KABUPATEN" ||
    upper === "SEKRETARIAT DAERAH KABUPATEN TOBA" ||
    upper === "SEKRETARIAT DAERAH KAB. TOBA" ||
    upper === "SETDA" ||
    upper === "SETDAKAB"
  ) {
    return "Sekretariat Daerah Kabupaten";
  }

  // 2. Deteksi variasi nama lama BAPPEDA / BAPELITBANGDA / PENELITIAN DAN PENGEMBANGAN
  if (
    upper.indexOf("PERENCANAAN PEMBANGUNAN") !== -1 ||
    upper.indexOf("PENELITIAN DAN PENGEMBANGAN") !== -1 ||
    upper.indexOf("PENELITIAN") !== -1 ||
    upper.indexOf("RISET DAN INOVASI") !== -1 ||
    upper.indexOf("BAPPEDA") !== -1 ||
    upper.indexOf("BAPPERIDA") !== -1 ||
    upper.indexOf("BAPELITBANGDA") !== -1
  ) {
    // Pastikan tidak keliru mencocokkan Bagian Perencanaan dan Keuangan di Sekretariat Daerah
    if (upper.indexOf("SEKRETARIAT DAERAH") === -1 && upper.indexOf("SETDA") === -1) {
      return "Badan Perencanaan Pembangunan, Riset dan Inovasi Daerah";
    }
  }

  // 3. Deteksi Dinas Kesehatan
  if (upper === "DINKES" || upper === "DISKES" || upper === "DINAS KESEHATAN KABUPATEN TOBA" || upper === "DINAS KESEHATAN KAB. TOBA") {
    return "Dinas Kesehatan";
  }

  // 4. Deteksi Inspektorat
  if (upper === "INSPEKTORAT" || upper === "INSPEKTORAT KABUPATEN TOBA" || upper === "INSPEKTORAT KAB. TOBA") {
    return "Inspektorat Daerah";
  }

  // 5. Deteksi BKPSDM
  if (upper === "BKPSDM" || upper === "BKD" || upper === "BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA KABUPATEN TOBA") {
    return "Badan Kepegawaian dan Pengembangan Sumber Daya Manusia";
  }

  return clean;
}

var RAW_TARGET_ORDER_PD = [
  "Sekretariat Daerah Kabupaten",
  "Sekretariat Dewan Perwakilan Rakyat Daerah",
  "Inspektorat Daerah",
  "Dinas Pendidikan, Pemuda dan Olahraga",
  "Dinas Kesehatan",
  "RSUD Porsea", 
  "Dinas Pekerjaan Umum dan Tata Ruang",
  "Dinas Perumahan dan Kawasan Permukiman",
  "Dinas Sosial",
  "Dinas Ketahanan Pangan dan Perikanan",
  "Dinas Lingkungan Hidup",
  "Dinas Kependudukan dan Pencatatan Sipil",
  "Dinas Pemberdayaan Masyarakat, Desa, Perempuan dan Perlindungan Anak",
  "Dinas Pengendalian Penduduk dan Keluarga Berencana",
  "Dinas Perhubungan",
  "Dinas Komunikasi dan Informatika",
  "Dinas Koperasi, Usaha Kecil Menengah, Perdagangan dan Perindustrian",
  "Dinas Penanaman Modal, Pelayanan Terpadu Satu Pintu dan Ketenagakerjaan",
  "Dinas Perpustakaan dan Kearsipan",
  "Dinas Kebudayaan dan Pariwisata",
  "Dinas Pertanian",
  "Badan Penanggulangan Bencana Daerah",
  "Badan Perencanaan Pembangunan, Riset dan Inovasi Daerah",
  "Badan Pengelolaan Keuangan dan Aset Daerah",
  "Badan Pengelolaan Pendapatan Daerah",
  "Badan Kepegawaian dan Pengembangan Sumber Daya Manusia",
  "Badan Kesatuan Bangsa dan Politik",
  "Satuan Polisi Pamong Praja",
  "Kecamatan Balige",
  "Kecamatan Tampahan",
  "Kecamatan Laguboti",
  "Kecamatan Silaen",
  "Kecamatan Habinsaran",
  "Kecamatan Borbor",
  "Kecamatan Nassau",
  "Kecamatan Sigumpar",
  "Kecamatan Porsea",
  "Kecamatan Siantar Narumonda",
  "Kecamatan Uluan",
  "Kecamatan Pintu Pohan Meranti",
  "Kecamatan Lumban Julu",
  "Kecamatan Ajibata",
  "Kecamatan Parmaksian",
  "Kecamatan Bonatua Lunasi"
];

// Pastikan seluruh entri master array selalu bersih dan normal
var TARGET_ORDER_PD = RAW_TARGET_ORDER_PD.map(normalizeOpdName);

function mapUnorToTarget(rawInduk, rawSatker, rawUnor, rawJabatan) {
  var induk = String(rawInduk || '').toUpperCase().trim();
  var satker = String(rawSatker || '').toUpperCase().trim();
  var unor = String(rawUnor || '').toUpperCase().trim();
  var jab = String(rawJabatan || '').toUpperCase().trim();
  var combined = (induk + " " + satker + " " + unor + (jab ? " " + jab : "")).trim();

  if (!combined) return "Lainnya";

  // -------------------------------------------------------------
  // 1. SEKRETARIAT DEWAN (DPRD) - Legislatif
  // -------------------------------------------------------------
  if (
    combined.indexOf("SEKRETARIAT DEWAN") !== -1 ||
    combined.indexOf("SEKRETARIAT DPRD") !== -1 ||
    combined.indexOf("SETWAN") !== -1 ||
    satker.indexOf("DPRD") !== -1 ||
    unor.indexOf("DPRD") !== -1
  ) {
    return "Sekretariat Dewan Perwakilan Rakyat Daerah";
  }

  // -------------------------------------------------------------
  // 2. INSPEKTORAT DAERAH - Pengawasan
  // -------------------------------------------------------------
  if (combined.indexOf("INSPEKTORAT") !== -1) {
    return "Inspektorat Daerah";
  }

  // -------------------------------------------------------------
  // 3. RSUD PORSEA - Rumah Sakit Daerah
  // -------------------------------------------------------------
  if (
    combined.indexOf("RSUD") !== -1 ||
    combined.indexOf("RUMAH SAKIT UMUM DAERAH") !== -1 ||
    (combined.indexOf("PORSEA") !== -1 && combined.indexOf("RUMAH SAKIT") !== -1)
  ) {
    return "RSUD Porsea";
  }

  // -------------------------------------------------------------
  // 4. SEKRETARIAT DAERAH KABUPATEN (SETDA)
  // (Pimpinan Daerah, Staf Ahli Bupati, Asisten 1/2/3, & 9 Bagian Resmi Setda)
  // -------------------------------------------------------------
  var isSetdaContext = (
    combined.indexOf("STAF AHLI BUPATI") !== -1 ||
    combined.indexOf("STAF AHLI") !== -1 ||
    combined.indexOf("SEKRETARIS DAERAH") !== -1 ||
    combined.indexOf("SEKDA") !== -1 ||
    combined.indexOf("SETDAKAB") !== -1 ||
    (combined.indexOf("SETDA") !== -1 && combined.indexOf("SETWAN") === -1) ||
    satker.indexOf("SEKRETARIAT DAERAH") !== -1 ||
    unor.indexOf("SEKRETARIAT DAERAH") !== -1 ||
    combined.indexOf("ASISTEN PEMERINTAHAN") !== -1 ||
    combined.indexOf("ASISTEN PEREKONOMIAN") !== -1 ||
    combined.indexOf("ASISTEN ADMINISTRASI") !== -1 ||
    combined.indexOf("ASISTEN BIDANG") !== -1 ||
    combined.indexOf("ASISTEN I") !== -1 ||
    combined.indexOf("ASISTEN II") !== -1 ||
    combined.indexOf("ASISTEN III") !== -1 ||
    combined.indexOf("ASISTEN SEKRETARIS DAERAH") !== -1 ||
    combined.indexOf("BAGIAN TATA PEMERINTAHAN") !== -1 ||
    combined.indexOf("BAGIAN KESEJAHTERAAN RAKYAT") !== -1 ||
    combined.indexOf("BAGIAN HUKUM") !== -1 ||
    combined.indexOf("BAGIAN PEREKONOMIAN") !== -1 ||
    combined.indexOf("BAGIAN PENGADAAN BARANG") !== -1 ||
    combined.indexOf("BAGIAN ADMINISTRASI PEMBANGUNAN") !== -1 ||
    combined.indexOf("BAGIAN ORGANISASI") !== -1 ||
    combined.indexOf("BAGIAN PROTOKOL") !== -1 ||
    combined.indexOf("BAGIAN PERENCANAAN DAN KEUANGAN SEKRETARIAT DAERAH") !== -1 ||
    (combined.indexOf("BAGIAN UMUM") !== -1 && 
     combined.indexOf("SUBBAGIAN") === -1 && 
     combined.indexOf("SUB BAGIAN") === -1 && 
     combined.indexOf("SUB-BAGIAN") === -1 && 
     combined.indexOf("DINAS") === -1 && 
     combined.indexOf("BADAN") === -1)
  );

  if (isSetdaContext) {
    return "Sekretariat Daerah Kabupaten";
  }

  // -------------------------------------------------------------
  // 5. DINAS PENDIDIKAN, PEMUDA DAN OLAHRAGA
  // -------------------------------------------------------------
  var isDinasPendidikan = (
    combined.indexOf("DINAS PENDIDIKAN") !== -1 ||
    combined.indexOf("DISDIK") !== -1 ||
    combined.indexOf("PEMUDA DAN OLAHRAGA") !== -1 ||
    combined.indexOf("DISDIKPORA") !== -1 ||
    combined.indexOf("SEKOLAH") !== -1 ||
    combined.indexOf("SD NEGERI") !== -1 || combined.indexOf("SDN ") !== -1 ||
    combined.indexOf("SMP NEGERI") !== -1 || combined.indexOf("SMPN ") !== -1 ||
    combined.indexOf("TK NEGERI") !== -1 || combined.indexOf("TKN ") !== -1 ||
    combined.indexOf("PAUD") !== -1 ||
    (combined.indexOf("SKB ") !== -1 || combined.indexOf(" SKB") !== -1 || combined.indexOf("SANGGAR KEGIATAN BELAJAR") !== -1) ||
    combined.indexOf("GURU") !== -1 ||
    combined.indexOf("PENGAWAS SEKOLAH") !== -1 ||
    combined.indexOf("PAMONG BELAJAR") !== -1
  );
  if (isDinasPendidikan) {
    return "Dinas Pendidikan, Pemuda dan Olahraga";
  }

  // -------------------------------------------------------------
  // 5. DINAS KESEHATAN (Puskesmas, Pustu, Labkesda, Farmasi)
  // -------------------------------------------------------------
  var isDinasKesehatan = (
    combined.indexOf("DINAS KESEHATAN") !== -1 ||
    combined.indexOf("DINKES") !== -1 ||
    combined.indexOf("DISKES") !== -1 ||
    combined.indexOf("PUSKESMAS") !== -1 ||
    combined.indexOf("POSKESDES") !== -1 ||
    combined.indexOf("PUSTU") !== -1 ||
    combined.indexOf("LABKESDA") !== -1 ||
    (combined.indexOf("FARMASI") !== -1 && combined.indexOf("KESEHATAN") !== -1)
  );
  if (isDinasKesehatan) {
    return "Dinas Kesehatan";
  }

  // -------------------------------------------------------------
  // 6. DINAS LINGKUNGAN HIDUP
  // -------------------------------------------------------------
  if (
    combined.indexOf("LINGKUNGAN HIDUP") !== -1 ||
    combined.indexOf("DLH") !== -1 ||
    combined.indexOf("DISLH") !== -1
  ) {
    return "Dinas Lingkungan Hidup";
  }

  // -------------------------------------------------------------
  // 7. DINAS KEPENDUDUKAN DAN PENCATATAN SIPIL
  // -------------------------------------------------------------
  if (
    combined.indexOf("KEPENDUDUKAN") !== -1 ||
    combined.indexOf("PENCATATAN SIPIL") !== -1 ||
    combined.indexOf("DUKCAPIL") !== -1 ||
    combined.indexOf("DISDUKCAPIL") !== -1
  ) {
    return "Dinas Kependudukan dan Pencatatan Sipil";
  }

  // -------------------------------------------------------------
  // 8. DINAS SOSIAL
  // -------------------------------------------------------------
  if (
    combined.indexOf("DINAS SOSIAL") !== -1 ||
    combined.indexOf("DINSOS") !== -1 ||
    (satker.indexOf("SOSIAL") !== -1 && satker.indexOf("KECAMATAN") === -1 && satker.indexOf("SEKRETARIAT") === -1) ||
    (unor.indexOf("DINAS SOSIAL") !== -1)
  ) {
    return "Dinas Sosial";
  }

  // -------------------------------------------------------------
  // 9. DINAS PERUMAHAN DAN KAWASAN PERMUKIMAN
  // -------------------------------------------------------------
  if (
    combined.indexOf("PERUMAHAN") !== -1 ||
    combined.indexOf("PERMUKIMAN") !== -1 ||
    combined.indexOf("PERKIM") !== -1 ||
    combined.indexOf("DISPERKIM") !== -1
  ) {
    return "Dinas Perumahan dan Kawasan Permukiman";
  }

  // -------------------------------------------------------------
  // 10. DINAS PEKERJAAN UMUM DAN TATA RUANG
  // -------------------------------------------------------------
  if (
    combined.indexOf("PEKERJAAN UMUM") !== -1 ||
    combined.indexOf("PUPR") !== -1 ||
    combined.indexOf("PUTR") !== -1 ||
    combined.indexOf("BINA MARGA") !== -1 ||
    combined.indexOf("CIPTA KARYA") !== -1 ||
    (combined.indexOf("TATA RUANG") !== -1 && combined.indexOf("SEKRETARIAT DAERAH") === -1 && combined.indexOf("TATA PEMERINTAHAN") === -1)
  ) {
    return "Dinas Pekerjaan Umum dan Tata Ruang";
  }

  // -------------------------------------------------------------
  // 11. DINAS KETAHANAN PANGAN DAN PERIKANAN
  // Catatan: UPTD Balai Benih Ikan (BBI) Lumban Pea secara SOTK Pemkab Toba
  // berada di bawah Dinas Pertanian, sehingga Lumban Pea / BBI diarahkan ke Dinas Pertanian.
  // -------------------------------------------------------------
  var isBbiLumbanPea = (
    combined.indexOf("LUMBAN PEA") !== -1 ||
    combined.indexOf("BALAI BENIH IKAN") !== -1 ||
    combined.indexOf("BENIH IKAN") !== -1 ||
    combined.indexOf("BBI ") !== -1 ||
    combined.indexOf("BBI-") !== -1
  );
  if (isBbiLumbanPea) {
    return "Dinas Pertanian";
  }

  if (
    combined.indexOf("KETAHANAN PANGAN") !== -1 ||
    combined.indexOf("PERIKANAN") !== -1 ||
    combined.indexOf("BUDIDAYA IKAN") !== -1 ||
    (combined.indexOf("IKAN") !== -1 && combined.indexOf("PENDIDIKAN") === -1) ||
    combined.indexOf("DKPP") !== -1 ||
    combined.indexOf("DISPANG") !== -1 ||
    combined.indexOf("DUMPANG") !== -1
  ) {
    return "Dinas Ketahanan Pangan dan Perikanan";
  }

  // -------------------------------------------------------------
  // 12. DINAS PERHUBUNGAN
  // -------------------------------------------------------------
  if (
    combined.indexOf("PERHUBUNGAN") !== -1 ||
    combined.indexOf("DISHUB") !== -1
  ) {
    return "Dinas Perhubungan";
  }

  // -------------------------------------------------------------
  // 13. DINAS KOMUNIKASI DAN INFORMATIKA
  // (Pastikan tidak keliru dengan Bagian Protokol dan Komunikasi Pimpinan Setda)
  // -------------------------------------------------------------
  if (
    combined.indexOf("KOMINFO") !== -1 ||
    combined.indexOf("DISKOMINFO") !== -1 ||
    combined.indexOf("INFORMATIKA") !== -1 ||
    (combined.indexOf("KOMUNIKASI") !== -1 && combined.indexOf("PROTOKOL") === -1 && combined.indexOf("SEKRETARIAT DAERAH") === -1 && combined.indexOf("SETDA") === -1)
  ) {
    return "Dinas Komunikasi dan Informatika";
  }

  // -------------------------------------------------------------
  // 14. DINAS PEMBERDAYAAN MASYARAKAT, DESA, PEREMPUAN DAN PA
  // -------------------------------------------------------------
  if (
    combined.indexOf("PEMBERDAYAAN MASYARAKAT") !== -1 ||
    combined.indexOf("DPMD") !== -1 ||
    combined.indexOf("DP3A") !== -1 ||
    combined.indexOf("DPMPPA") !== -1 ||
    (combined.indexOf("PERLINDUNGAN ANAK") !== -1 && satker.indexOf("KECAMATAN") === -1) ||
    (combined.indexOf("PEMBERDAYAAN PEREMPUAN") !== -1 && satker.indexOf("KECAMATAN") === -1)
  ) {
    return "Dinas Pemberdayaan Masyarakat, Desa, Perempuan dan Perlindungan Anak";
  }

  // -------------------------------------------------------------
  // 15. DINAS PENGENDALIAN PENDUDUK DAN KELUARGA BERENCANA
  // -------------------------------------------------------------
  if (
    combined.indexOf("PENGENDALIAN PENDUDUK") !== -1 ||
    combined.indexOf("KELUARGA BERENCANA") !== -1 ||
    combined.indexOf("DPPKB") !== -1 ||
    combined.indexOf("DISDALDUK") !== -1 ||
    combined.indexOf("DALDUK") !== -1
  ) {
    return "Dinas Pengendalian Penduduk dan Keluarga Berencana";
  }

  // -------------------------------------------------------------
  // 16. DINAS KOPERASI, UKM, PERDAGANGAN DAN PERINDUSTRIAN
  // -------------------------------------------------------------
  if (
    combined.indexOf("KOPERASI") !== -1 ||
    combined.indexOf("PERDAGANGAN") !== -1 ||
    combined.indexOf("PERINDUSTRIAN") !== -1 ||
    combined.indexOf("DISPERINDAG") !== -1 ||
    combined.indexOf("DISKOPERINDAG") !== -1 ||
    (combined.indexOf("UKM") !== -1 && combined.indexOf("HUKUM") === -1) ||
    combined.indexOf("UMKM") !== -1
  ) {
    return "Dinas Koperasi, Usaha Kecil Menengah, Perdagangan dan Perindustrian";
  }

  // -------------------------------------------------------------
  // 17. DINAS PENANAMAN MODAL, PTSP DAN KETENAGAKERJAAN
  // -------------------------------------------------------------
  if (
    combined.indexOf("PENANAMAN MODAL") !== -1 ||
    combined.indexOf("PELAYANAN TERPADU") !== -1 ||
    combined.indexOf("PTSP") !== -1 ||
    combined.indexOf("DPMPTSP") !== -1 ||
    combined.indexOf("PERIZINAN") !== -1 ||
    combined.indexOf("KETENAGAKERJAAN") !== -1 ||
    combined.indexOf("TENAGA KERJA") !== -1
  ) {
    return "Dinas Penanaman Modal, Pelayanan Terpadu Satu Pintu dan Ketenagakerjaan";
  }

  // -------------------------------------------------------------
  // 18. DINAS PERPUSTAKAAN DAN KEARSIPAN
  // -------------------------------------------------------------
  if (
    combined.indexOf("PERPUSTAKAAN") !== -1 ||
    combined.indexOf("KEARSIPAN") !== -1 ||
    combined.indexOf("DISPUSIP") !== -1
  ) {
    return "Dinas Perpustakaan dan Kearsipan";
  }

  // -------------------------------------------------------------
  // 19. DINAS KEBUDAYAAN DAN PARIWISATA
  // -------------------------------------------------------------
  if (
    combined.indexOf("KEBUDAYAAN") !== -1 ||
    combined.indexOf("PARIWISATA") !== -1 ||
    combined.indexOf("DISBUDPAR") !== -1
  ) {
    return "Dinas Kebudayaan dan Pariwisata";
  }

  // -------------------------------------------------------------
  // 20. DINAS PERTANIAN
  // -------------------------------------------------------------
  var isPertanian = (
    combined.indexOf("PERTANIAN") !== -1 ||
    combined.indexOf("DISPERTAN") !== -1 ||
    combined.indexOf("PETERNAKAN") !== -1 ||
    combined.indexOf("PERKEBUNAN") !== -1 ||
    combined.indexOf("BALAI PENYULUHAN PERTANIAN") !== -1 ||
    combined.indexOf("BPP ") !== -1 ||
    combined.indexOf("UPTD PERTANIAN") !== -1 ||
    combined.indexOf("RUMAH POTONG HEWAN") !== -1 ||
    combined.indexOf("RPH") !== -1 ||
    combined.indexOf("PUSKESWAN") !== -1 ||
    combined.indexOf("KESEHATAN HEWAN") !== -1 ||
    combined.indexOf("VETERINER") !== -1 ||
    combined.indexOf("BALAI BENIH IKAN") !== -1 ||
    combined.indexOf("BENIH IKAN") !== -1 ||
    combined.indexOf("LUMBAN PEA") !== -1 ||
    combined.indexOf("BBI ") !== -1 ||
    combined.indexOf("BBI-") !== -1
  );
  if (isPertanian) {
    return "Dinas Pertanian";
  }

  // -------------------------------------------------------------
  // 21. BADAN PENANGGULANGAN BENCANA DAERAH (BPBD)
  // -------------------------------------------------------------
  if (
    combined.indexOf("PENANGGULANGAN BENCANA") !== -1 ||
    combined.indexOf("BPBD") !== -1
  ) {
    return "Badan Penanggulangan Bencana Daerah";
  }

  // -------------------------------------------------------------
  // 22. BADAN PERENCANAAN PEMBANGUNAN, RISET DAN INOVASI DAERAH (BAPPERIDA)
  // -------------------------------------------------------------
  if (
    combined.indexOf("BAPPERIDA") !== -1 ||
    combined.indexOf("BAPPEDA") !== -1 ||
    combined.indexOf("BAPELITBANGDA") !== -1 ||
    combined.indexOf("RISET DAN INOVASI") !== -1 ||
    (combined.indexOf("PERENCANAAN PEMBANGUNAN") !== -1 && combined.indexOf("SEKRETARIAT DAERAH") === -1 && combined.indexOf("SETDA") === -1) ||
    (combined.indexOf("PENELITIAN DAN PENGEMBANGAN") !== -1 && combined.indexOf("SEKRETARIAT DAERAH") === -1 && combined.indexOf("SETDA") === -1)
  ) {
    return "Badan Perencanaan Pembangunan, Riset dan Inovasi Daerah";
  }

  // -------------------------------------------------------------
  // 23. BADAN PENGELOLAAN KEUANGAN DAN ASET DAERAH (BPKAD)
  // -------------------------------------------------------------
  if (
    combined.indexOf("KEUANGAN DAN ASET") !== -1 ||
    combined.indexOf("BPKAD") !== -1
  ) {
    return "Badan Pengelolaan Keuangan dan Aset Daerah";
  }

  // -------------------------------------------------------------
  // 24. BADAN PENGELOLAAN PENDAPATAN DAERAH (BAPENDA)
  // -------------------------------------------------------------
  if (
    combined.indexOf("PENDAPATAN DAERAH") !== -1 ||
    combined.indexOf("BAPENDA") !== -1 ||
    combined.indexOf("DISPENDA") !== -1
  ) {
    return "Badan Pengelolaan Pendapatan Daerah";
  }

  // -------------------------------------------------------------
  // 25. BADAN KEPEGAWAIAN DAN PENGEMBANGAN SDM (BKPSDM)
  // -------------------------------------------------------------
  if (
    combined.indexOf("KEPEGAWAIAN DAN PENGEMBANGAN") !== -1 ||
    combined.indexOf("BKPSDM") !== -1 ||
    combined.indexOf("BKD") !== -1
  ) {
    return "Badan Kepegawaian dan Pengembangan Sumber Daya Manusia";
  }

  // -------------------------------------------------------------
  // 26. BADAN KESATUAN BANGSA DAN POLITIK (KESBANGPOL)
  // -------------------------------------------------------------
  if (
    combined.indexOf("KESATUAN BANGSA") !== -1 ||
    combined.indexOf("KESBANGPOL") !== -1
  ) {
    return "Badan Kesatuan Bangsa dan Politik";
  }

  // -------------------------------------------------------------
  // 27. SATUAN POLISI PAMONG PRAJA (SATPOL PP)
  // -------------------------------------------------------------
  if (
    combined.indexOf("POLISI PAMONG PRAJA") !== -1 ||
    combined.indexOf("SATPOL PP") !== -1 ||
    combined.indexOf("SATPOLPP") !== -1 ||
    combined.indexOf("POL PP") !== -1
  ) {
    return "Satuan Polisi Pamong Praja";
  }

  // -------------------------------------------------------------
  // 28. 16 KECAMATAN SE-KABUPATEN TOBA
  // -------------------------------------------------------------
  var isKecamatanContext = (
    satker.indexOf("KECAMATAN") !== -1 || satker.indexOf("CAMAT") !== -1 || 
    induk.indexOf("KECAMATAN") !== -1 || induk.indexOf("CAMAT") !== -1 || 
    unor.indexOf("KECAMATAN") !== -1 || unor.indexOf("CAMAT") !== -1 || 
    unor.indexOf("KANTOR CAMAT") !== -1 ||
    combined.indexOf("KELURAHAN") !== -1 || combined.indexOf("LURAH") !== -1
  );

  if (isKecamatanContext) {
    if (combined.indexOf("BALIGE") !== -1) return "Kecamatan Balige";
    if (combined.indexOf("TAMPAHAN") !== -1) return "Kecamatan Tampahan";
    if (combined.indexOf("LAGUBOTI") !== -1) return "Kecamatan Laguboti";
    if (combined.indexOf("SILAEN") !== -1) return "Kecamatan Silaen";
    if (combined.indexOf("HABINSARAN") !== -1) return "Kecamatan Habinsaran";
    if (combined.indexOf("BORBOR") !== -1) return "Kecamatan Borbor";
    if (combined.indexOf("NASSAU") !== -1) return "Kecamatan Nassau";
    if (combined.indexOf("SIGUMPAR") !== -1) return "Kecamatan Sigumpar";
    if (combined.indexOf("PORSEA") !== -1 && combined.indexOf("RSUD") === -1) return "Kecamatan Porsea";
    if (combined.indexOf("NARUMONDA") !== -1 || combined.indexOf("SIANTAR NARUMONDA") !== -1) return "Kecamatan Siantar Narumonda";
    if (combined.indexOf("ULUAN") !== -1) return "Kecamatan Uluan";
    if (combined.indexOf("PINTU POHAN") !== -1 || combined.indexOf("PINTUPOHAN") !== -1) return "Kecamatan Pintu Pohan Meranti";
    if (combined.indexOf("LUMBAN JULU") !== -1 || combined.indexOf("LUMBANJULU") !== -1) return "Kecamatan Lumban Julu";
    if (combined.indexOf("AJIBATA") !== -1) return "Kecamatan Ajibata";
    if (combined.indexOf("PARMAKSIAN") !== -1) return "Kecamatan Parmaksian";
    if (combined.indexOf("BONATUA LUNASI") !== -1 || combined.indexOf("BONATUA") !== -1) return "Kecamatan Bonatua Lunasi";
  }

  // -------------------------------------------------------------
  // 29. FALLBACK MATCHING DENGAN TARGET_ORDER_PD
  // -------------------------------------------------------------
  for (var k = 0; k < TARGET_ORDER_PD.length; k++) {
    var pd = TARGET_ORDER_PD[k];
    if (pd === "Sekretariat Daerah Kabupaten") {
      if (combined.indexOf("SEKRETARIAT DAERAH") !== -1 || combined.indexOf("SETDA") !== -1) {
        return "Sekretariat Daerah Kabupaten";
      }
      continue;
    }
    var cleanPd = pd.toUpperCase().replace("KABUPATEN", "").replace("DAERAH", "").trim();
    if (cleanPd.length >= 5 && combined.indexOf(cleanPd) !== -1) {
      return normalizeOpdName(pd);
    }
  }

  return "Lainnya";
}