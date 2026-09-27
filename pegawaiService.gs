/**
 * =========================================================================
 * ACRU EMPLOYEE ANALYTICS SYSTEM - KABUPATEN TOBA
 * File: pegawaiService.gs
 * Deskripsi: Service Pegawai ASN dengan filter multidimensi (Status ASN,
 * Golongan, Pendidikan, Satker Multi-Select, Eselon, & 3 Jenis Jabatan:
 * Struktural, Fungsional, Pelaksana), normalisasi BAPPERIDA, agregasi analitik,
 * dan paginasi backend berkinerja tinggi.
 * =========================================================================
 */

/**
 * Daftar resmi 44 Perangkat Daerah Kabupaten Toba
 */
function getMasterOPDList() {
  var list = [
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

  return list.map(normalizeOpdName);
}



/**
 * Smart Categorizer Jenis Jabatan:
 * Mengelompokkan pegawai menjadi: 'STRUKTURAL', 'FUNGSIONAL', atau 'PELAKSANA'
 */
function categorizePegawaiJabatanLocal(rawJenis, rawJabatan, rawEselon) {
  var j = String(rawJenis || '').toUpperCase().trim();
  var jab = String(rawJabatan || '').toUpperCase().trim();
  var esl = String(rawEselon || '').toUpperCase().trim();

  // 1. Deteksi Jabatan Struktural (Manajerial)
  var hasEselon = esl && esl !== '-' && esl !== 'NON ESELON' && esl !== 'NON' && esl !== '0';
  var isStrukturalText = j.indexOf('STRUKTURAL') !== -1 || j.indexOf('PIMPINAN') !== -1 || j.indexOf('ADMINISTRATOR') !== -1 || j.indexOf('PENGAWAS') !== -1;
  
  if (hasEselon || isStrukturalText) {
    return {
      type: 'STRUKTURAL',
      label: 'Jabatan Struktural',
      jfCategory: '-',
      jfLevel: '-'
    };
  }

  // 2. Deteksi Jabatan Fungsional (JF)
  var isFungsionalText = j.indexOf('FUNGSIONAL') !== -1 || j.indexOf('JFT') !== -1 || j.indexOf('TERTENTU') !== -1;
  var isJfLevelKeahlian = jab.indexOf('AHLI UTAMA') !== -1 || jab.indexOf('AHLI MADYA') !== -1 || jab.indexOf('AHLI MUDA') !== -1 || jab.indexOf('AHLI PERTAMA') !== -1;
  var isJfLevelKeterampilan = jab.indexOf('PENYELIA') !== -1 || jab.indexOf('MAHIR') !== -1 || jab.indexOf('PELAKSANA LANJUTAN') !== -1 || jab.indexOf('TERAMPIL') !== -1 || (jab.indexOf('PELAKSANA') !== -1 && isFungsionalText) || jab.indexOf('PEMULA') !== -1;
  var isJfProfesi = jab.indexOf('GURU') !== -1 || jab.indexOf('DOKTER') !== -1 || jab.indexOf('BIDAN') !== -1 || jab.indexOf('PERAWAT') !== -1 || jab.indexOf('AUDITOR') !== -1 || jab.indexOf('PRANATA') !== -1 || jab.indexOf('PENGAWAS SEKOLAH') !== -1 || jab.indexOf('PAMONG') !== -1 || jab.indexOf('PENYULUH') !== -1 || jab.indexOf('EPIDEMIOLOG') !== -1 || jab.indexOf('SANITARIAN') !== -1 || jab.indexOf('NUTRISIONIS') !== -1 || jab.indexOf('APOTEKER') !== -1 || jab.indexOf('ASISTEN APOTEKER') !== -1;

  if (isFungsionalText || isJfLevelKeahlian || isJfLevelKeterampilan || isJfProfesi) {
    var jfCat = 'KEAHLIAN';
    var jfLvl = 'Ahli Pertama';

    if (jab.indexOf('AHLI UTAMA') !== -1 || (jab.indexOf('UTAMA') !== -1 && isFungsionalText)) {
      jfCat = 'KEAHLIAN'; jfLvl = 'Ahli Utama';
    } else if (jab.indexOf('AHLI MADYA') !== -1 || jab.indexOf('MADYA') !== -1) {
      jfCat = 'KEAHLIAN'; jfLvl = 'Ahli Madya';
    } else if (jab.indexOf('AHLI MUDA') !== -1 || jab.indexOf('MUDA') !== -1) {
      jfCat = 'KEAHLIAN'; jfLvl = 'Ahli Muda';
    } else if (jab.indexOf('AHLI PERTAMA') !== -1 || jab.indexOf('PERTAMA') !== -1) {
      jfCat = 'KEAHLIAN'; jfLvl = 'Ahli Pertama';
    } else if (jab.indexOf('PENYELIA') !== -1) {
      jfCat = 'KETERAMPILAN'; jfLvl = 'Penyelia';
    } else if (jab.indexOf('MAHIR') !== -1 || jab.indexOf('PELAKSANA LANJUTAN') !== -1) {
      jfCat = 'KETERAMPILAN'; jfLvl = 'Mahir';
    } else if (jab.indexOf('TERAMPIL') !== -1) {
      jfCat = 'KETERAMPILAN'; jfLvl = 'Terampil';
    } else if (jab.indexOf('PEMULA') !== -1) {
      jfCat = 'KETERAMPILAN'; jfLvl = 'Pemula';
    }

    return {
      type: 'FUNGSIONAL',
      label: 'Jabatan Fungsional',
      jfCategory: jfCat,
      jfLevel: jfLvl
    };
  }

  // 3. Default: Jabatan Pelaksana (Staf Administrasi / Teknis)
  return {
    type: 'PELAKSANA',
    label: 'Jabatan Pelaksana',
    jfCategory: '-',
    jfLevel: '-'
  };
}

/**
 * Mengambil opsi filter dinamis untuk dropdown (Golongan, Pendidikan, Satker, Jenis Jabatan, Eselon)
 */
function getFilterOptions() {
  try {
    try {
      if (typeof ensureDatabaseReady === 'function') ensureDatabaseReady();
    } catch (e) {
      Logger.log("Init skipped in getFilterOptions: " + e);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Spreadsheet aktif tidak ditemukan.");
    
    const targetSheets = ['Data_Pegawai_PNS', 'Data_Pegawai_PPPK', 'Data_Pegawai_PPPK_Paruh_Waktu'];
    const activeSheets = [];
    
    for (let sIdx = 0; sIdx < targetSheets.length; sIdx++) {
      const name = targetSheets[sIdx];
      const sh = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, name) : ss.getSheetByName(name);
      if (sh && sh.getLastRow() > 1) activeSheets.push(sh);
    }

    const setGol = new Set();
    const setPend = new Set();
    const setSatker = new Set();
    const setEselon = new Set();

    // 1. Injeksi 44 OPD resmi Kabupaten Toba sebagai pondasi utama
    const masterPD = getMasterOPDList();
    masterPD.forEach(function(pd) { 
      setSatker.add(normalizeOpdName(pd)); 
    });

    // 2. Pindai sheet data pegawai
    for (let sIdx = 0; sIdx < activeSheets.length; sIdx++) {
      const sheet = activeSheets[sIdx];
      const range = sheet.getDataRange();
      const data = range.getDisplayValues();
      if (data.length <= 1) continue;

      const headers = data[0];
      const idxGol = findColumnIndex(headers, ["GOL AKHIR NAMA", "GOLONGAN", "GOL", "GOL RUANG", "PANGKAT GOLONGAN"]);
      const idxPend = findColumnIndex(headers, ["TINGKAT PENDIDIKAN NAMA", "PENDIDIKAN", "TINGKAT PENDIDIKAN", "PENDIDIKAN TERAKHIR"]);
      const idxSatker = findColumnIndex(headers, ["SATUAN KERJA KERJA NAMA", "SATKER", "SATUAN KERJA", "SATUAN KERJA NAMA"]);
      const idxSatkerInduk = findColumnIndex(headers, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK", "INDUK"]);
      const idxUnor = findColumnIndex(headers, ["UNOR NAMA", "UNOR", "UNIT ORGANISASI", "UNIT KERJA"]);
      const idxEselon = findColumnIndex(headers, ["ESELON NAMA", "ESELON", "ESELON_NAMA", "TINGKAT ESELON", "NAMA ESELON"]);
      const idxJabatan = findColumnIndex(headers, ["JABATAN NAMA", "JABATAN", "NAMA JABATAN"]);

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.every(function(cell) { return String(cell).trim() === ''; })) continue;
        
        if (idxGol !== -1 && row[idxGol]) setGol.add(String(row[idxGol]).trim());
        if (idxPend !== -1 && row[idxPend]) setPend.add(String(row[idxPend]).trim());
        if (idxEselon !== -1 && row[idxEselon]) setEselon.add(String(row[idxEselon]).trim());

        const sInduk = idxSatkerInduk !== -1 ? row[idxSatkerInduk] : '';
        const sKerja = idxSatker !== -1 ? row[idxSatker] : '';
        const sUnor = idxUnor !== -1 ? row[idxUnor] : '';
        const sJab = idxJabatan !== -1 ? row[idxJabatan] : '';

        if (typeof mapUnorToTarget === 'function') {
          const mappedPD = mapUnorToTarget(sInduk, sKerja, sUnor, sJab);
          if (mappedPD && mappedPD !== 'Lainnya') {
            setSatker.add(normalizeOpdName(mappedPD));
          }
        } else if (sKerja) {
          setSatker.add(normalizeOpdName(sKerja));
        }
      }
    }

    // 3. Sanitasi Ketat Satker
    const cleanSatkerList = Array.from(setSatker).map(normalizeOpdName).filter(function(name) {
      if (!name) return false;
      const u = name.toUpperCase();
      if (
        u.indexOf("PENELITIAN DAN PENGEMBANGAN") !== -1 ||
        u.indexOf("PENELITIAN") !== -1 ||
        u === "BAPPEDA" ||
        u === "BAPELITBANGDA"
      ) {
        return false;
      }
      return true;
    });

    let sortedSatker = [];
    masterPD.forEach(function(item) {
      if (sortedSatker.indexOf(item) === -1 && item) {
        sortedSatker.push(item);
      }
    });

    cleanSatkerList.forEach(function(item) {
      if (sortedSatker.indexOf(item) === -1 && item) {
        sortedSatker.push(item);
      }
    });

    // 4. Opsi Standar 3 Jenis Jabatan
    const standardJenisJabatan = [
      { id: "STRUKTURAL", label: "Jabatan Struktural" },
      { id: "FUNGSIONAL", label: "Jabatan Fungsional" },
      { id: "PELAKSANA", label: "Jabatan Pelaksana" }
    ];

    return {
      gol: Array.from(setGol).filter(Boolean).sort(),
      pend: Array.from(setPend).filter(Boolean).sort(),
      satker: sortedSatker.filter(Boolean),
      jenisJabatan: standardJenisJabatan,
      eselon: Array.from(setEselon).filter(Boolean).sort()
    };
  } catch (error) {
    Logger.log('Error getFilterOptions: ' + error.toString());
    return {
      gol: [],
      pend: [],
      satker: getMasterOPDList(),
      jenisJabatan: [
        { id: "STRUKTURAL", label: "Jabatan Struktural" },
        { id: "FUNGSIONAL", label: "Jabatan Fungsional" },
        { id: "PELAKSANA", label: "Jabatan Pelaksana" }
      ],
      eselon: [],
      error: error.toString()
    };
  }
}

/**
 * Mengambil data pegawai dengan paginasi, filter multidimensi (termasuk Jenis Jabatan),
 * dan data analitik grafik (Gender, Pendidikan, & Komposisi Jenis Jabatan)
 */
function getPegawaiDashboardData(filters, limit, offset) {
  try {
    try {
      if (typeof ensureDatabaseReady === 'function') ensureDatabaseReady();
    } catch (e) {
      Logger.log("Init skipped: " + e);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Spreadsheet aktif tidak ditemukan.");

    filters = filters || {};
    limit = parseInt(limit, 10) || 15;
    offset = parseInt(offset, 10) || 0;

    const sheetsToProcess = [];
    if (!filters.jenisAsn || filters.jenisAsn === 'PNS') {
      const s = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, 'Data_Pegawai_PNS') : ss.getSheetByName('Data_Pegawai_PNS');
      if (s && s.getLastRow() > 1) sheetsToProcess.push({ sheetObj: s, typeTag: 'PNS' });
    }
    if (!filters.jenisAsn || filters.jenisAsn === 'PPPK') {
      const s = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, 'Data_Pegawai_PPPK') : ss.getSheetByName('Data_Pegawai_PPPK');
      if (s && s.getLastRow() > 1) sheetsToProcess.push({ sheetObj: s, typeTag: 'PPPK' });
    }
    if (!filters.jenisAsn || filters.jenisAsn === 'PPPK_PARUH_WAKTU') {
      const s = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, 'Data_Pegawai_PPPK_Paruh_Waktu') : ss.getSheetByName('Data_Pegawai_PPPK_Paruh_Waktu');
      if (s && s.getLastRow() > 1) sheetsToProcess.push({ sheetObj: s, typeTag: 'PPPK_PARUH_WAKTU' });
    }

    if (sheetsToProcess.length === 0) {
      const all = ss.getSheets();
      for (let sIdx = 0; sIdx < all.length; sIdx++) {
        const s = all[sIdx];
        if (s.getLastRow() > 1 && !s.getName().startsWith('Log_') && !s.getName().includes('Struktur')) {
          sheetsToProcess.push({ sheetObj: s, typeTag: 'PNS' });
        }
      }
    }

    let allFilteredRows = [];
    const genderDist = { 'LAKI-LAKI': 0, 'PEREMPUAN': 0 };
    const eduDist = {};
    const religionDist = {};
    const empTypeDist = { 'PNS': 0, 'PPPK': 0, 'PPPK_PARUH_WAKTU': 0 };
    const jabatanDist = { 'STRUKTURAL': 0, 'FUNGSIONAL': 0, 'PELAKSANA': 0 };

    let selectedSatkers = [];
    if (Array.isArray(filters.satkerList) && filters.satkerList.length > 0) {
      selectedSatkers = filters.satkerList.map(function(s) { 
        return normalizeOpdName(String(s).trim()); 
      }).filter(Boolean);
    } else if (filters.satker && typeof filters.satker === 'string' && filters.satker.trim() !== '') {
      selectedSatkers = [normalizeOpdName(filters.satker.trim())];
    }

    for (let shIdx = 0; shIdx < sheetsToProcess.length; shIdx++) {
      const sInfo = sheetsToProcess[shIdx];
      const rawData = sInfo.sheetObj.getDataRange().getDisplayValues();
      if (rawData.length <= 1) continue;

      const headers = rawData[0];
      const idxPnsId = findColumnIndex(headers, ["PNS ID", "ID", "NO", "KODE"]);
      const idxNip = findColumnIndex(headers, ["NIP BARU", "NIP", "NIP_BARU", "NOMOR INDUK"]);
      const idxNama = findColumnIndex(headers, ["NAMA", "NAMA LENGKAP", "NAMA PEGAWAI"]);
      const idxGelarD = findColumnIndex(headers, ["GELAR DEPAN", "GLR_DEPAN"]);
      const idxGelarB = findColumnIndex(headers, ["GELAR BELAKANG", "GLR_BELAKANG"]);
      const idxGol = findColumnIndex(headers, ["GOL AKHIR NAMA", "GOLONGAN", "GOL", "PANGKAT/GOLONGAN", "GOL RUANG"]);
      const idxPend = findColumnIndex(headers, ["TINGKAT PENDIDIKAN NAMA", "PENDIDIKAN", "TINGKAT PENDIDIKAN", "PENDIDIKAN TERAKHIR"]);
      const idxSatker = findColumnIndex(headers, ["SATUAN KERJA KERJA NAMA", "SATKER", "SATUAN KERJA", "SATUAN KERJA NAMA"]);
      const idxSatkerInduk = findColumnIndex(headers, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK", "INDUK"]);
      const idxKelamin = findColumnIndex(headers, ["JENIS KELAMIN", "JENIS_KELAMIN", "JENIS KELAMIN NAMA", "NAMA KELAMIN", "KELAMIN", "JK", "GENDER", "SEX"]);
      const idxAgama = findColumnIndex(headers, ["AGAMA NAMA", "AGAMA"]);
      const idxJenisJabatan = findColumnIndex(headers, ["JENIS JABATAN NAMA", "JENIS JABATAN", "JENIS_JABATAN"]);
      const idxUnor = findColumnIndex(headers, ["UNOR NAMA", "UNOR", "UNIT ORGANISASI", "UNIT KERJA"]);
      const idxJabatanNama = findColumnIndex(headers, ["JABATAN NAMA", "JABATAN", "NAMA JABATAN"]);
      const idxEselon = findColumnIndex(headers, ["ESELON NAMA", "ESELON", "ESELON_NAMA", "TINGKAT ESELON", "NAMA ESELON"]);

      for (let r = 1; r < rawData.length; r++) {
        const row = rawData[r];
        if (row.every(function(cell) { return String(cell).trim() === ''; })) continue;

        const cellNip = idxNip !== -1 ? String(row[idxNip] || '').trim() : '';
        const cellNama = idxNama !== -1 ? String(row[idxNama] || '').trim() : '';
        const cellGol = idxGol !== -1 ? String(row[idxGol] || '').trim() : '';
        const cellPend = idxPend !== -1 ? String(row[idxPend] || '').trim() : '';
        const rawCellSatker = idxSatker !== -1 ? String(row[idxSatker] || '').trim() : '';
        const rawCellSatkerInduk = idxSatkerInduk !== -1 ? String(row[idxSatkerInduk] || '').trim() : '';
        const rawCellJenisJabatan = idxJenisJabatan !== -1 ? String(row[idxJenisJabatan] || '').trim() : '';
        const rawCellUnor = idxUnor !== -1 ? String(row[idxUnor] || '').trim() : '';
        const cellJabatanNama = idxJabatanNama !== -1 ? String(row[idxJabatanNama] || '').trim() : '';
        const cellEselon = idxEselon !== -1 ? String(row[idxEselon] || '').trim() : '';

        // Tentukan nama resmi perangkat daerah
        const mappedPD = (typeof mapUnorToTarget === 'function')
          ? normalizeOpdName(mapUnorToTarget(rawCellSatkerInduk, rawCellSatker, rawCellUnor, cellJabatanNama))
          : normalizeOpdName(rawCellSatker);

        // Identifikasi Cerdas Jenis Jabatan
        const jabMeta = categorizePegawaiJabatanLocal(rawCellJenisJabatan, cellJabatanNama, cellEselon);
        const catJabatan = jabMeta.type; // 'STRUKTURAL' | 'FUNGSIONAL' | 'PELAKSANA'

        // 1. Filter Pencarian Cepat
        if (filters.search) {
          const sQuery = filters.search.toLowerCase();
          const matchesSearch = cellNama.toLowerCase().indexOf(sQuery) !== -1 ||
                                cellNip.toLowerCase().indexOf(sQuery) !== -1 ||
                                cellJabatanNama.toLowerCase().indexOf(sQuery) !== -1;
          if (!matchesSearch) continue;
        }

        // 2. Filter Golongan, Pendidikan, & Eselon
        if (filters.gol && idxGol !== -1 && cellGol !== filters.gol.trim()) continue;
        if (filters.pend && idxPend !== -1 && cellPend !== filters.pend.trim()) continue;

        // 3. Filter Jenis Jabatan (Struktural, Fungsional, Pelaksana)
        if (filters.jenisJabatan) {
          const fJab = String(filters.jenisJabatan).toUpperCase().trim();
          if (fJab === 'STRUKTURAL' || fJab.indexOf('STRUKTUR') !== -1) {
            if (catJabatan !== 'STRUKTURAL') continue;
          } else if (fJab === 'FUNGSIONAL' || fJab.indexOf('FUNGSI') !== -1) {
            if (catJabatan !== 'FUNGSIONAL') continue;
          } else if (fJab === 'PELAKSANA' || fJab.indexOf('PELAKS') !== -1) {
            if (catJabatan !== 'PELAKSANA') continue;
          } else if (fJab !== '' && rawCellJenisJabatan.toUpperCase() !== fJab) {
            continue;
          }
        }

        // 4. Filter Satuan Kerja Multi-Select (Ketat dan Presisi sesuai mappedPD)
        if (selectedSatkers.length > 0) {
          const targetPdNorm = String(mappedPD || '').toLowerCase().trim();

          const matchesAnySatker = selectedSatkers.some(function(filterPD) {
            const cleanFilter = String(filterPD || '').toLowerCase().trim();
            const normFilter = normalizeOpdName(cleanFilter).toLowerCase().trim();

            // 1. Direct exact match
            if (targetPdNorm === cleanFilter || targetPdNorm === normFilter) return true;

            // 2. Canonical alias: Sekretariat Daerah
            const isFilterSetda = normFilter.indexOf("sekretariat daerah") !== -1 || normFilter === "setda" || normFilter === "setdakab";
            const isTargetSetda = targetPdNorm.indexOf("sekretariat daerah") !== -1 || targetPdNorm === "setda" || targetPdNorm === "setdakab";
            if (isFilterSetda && isTargetSetda) return true;

            // 3. Canonical alias: BAPPERIDA / BAPPEDA / Riset dan Inovasi
            const isFilterBapperida = (
              normFilter.indexOf("riset dan inovasi") !== -1 ||
              normFilter.indexOf("bapperida") !== -1 ||
              normFilter.indexOf("penelitian dan pengembangan") !== -1 ||
              normFilter.indexOf("bappeda") !== -1
            );
            const isTargetBapperida = (
              targetPdNorm.indexOf("riset dan inovasi") !== -1 ||
              targetPdNorm.indexOf("bapperida") !== -1 ||
              targetPdNorm.indexOf("penelitian dan pengembangan") !== -1 ||
              targetPdNorm.indexOf("bappeda") !== -1
            );
            if (isFilterBapperida && isTargetBapperida) return true;

            // 4. Fallback HANYA jika pegawai belum terpetakan ('Lainnya' atau kosong)
            if (!targetPdNorm || targetPdNorm === 'lainnya') {
              const rawCombined = (rawCellSatkerInduk + ' ' + rawCellSatker + ' ' + rawCellUnor + ' ' + cellJabatanNama).toLowerCase();
              if (normFilter.length >= 6 && rawCombined.indexOf(normFilter) !== -1) {
                return true;
              }
            }

            return false;
          });

          if (!matchesAnySatker) continue;
        }

        // 5. Filter Eselon
        if (filters.eselon) {
          const upperCellEselon = cellEselon.toUpperCase().trim();
          const filterEselon = filters.eselon.toUpperCase().trim();
          let matchesEselon = false;

          if (upperCellEselon === filterEselon) {
            matchesEselon = true;
          } else if (filterEselon === 'II' || filterEselon === 'ESELON II') {
            matchesEselon = (upperCellEselon.indexOf('II') !== -1 || upperCellEselon.indexOf('2') !== -1) && 
                            upperCellEselon.indexOf('III') === -1 && upperCellEselon.indexOf('VII') === -1;
          } else if (filterEselon === 'III' || filterEselon === 'ESELON III') {
            matchesEselon = upperCellEselon.indexOf('III') !== -1 || upperCellEselon.indexOf('3') !== -1;
          } else if (filterEselon === 'IV' || filterEselon === 'ESELON IV') {
            matchesEselon = upperCellEselon.indexOf('IV') !== -1 || upperCellEselon.indexOf('4') !== -1;
          } else if (filterEselon === 'NON ESELON') {
            matchesEselon = upperCellEselon.indexOf('NON') !== -1 || upperCellEselon === '-' || upperCellEselon === '';
          } else {
            matchesEselon = upperCellEselon.indexOf(filterEselon) !== -1 || filterEselon.indexOf(upperCellEselon) !== -1;
          }

          if (!matchesEselon) continue;
        }

        // Akumulasi Metrik Distribusi
        let rawGender = idxKelamin !== -1 ? String(row[idxKelamin] || '').toUpperCase().trim() : '';
        let gender = 'LAKI-LAKI';

        if (rawGender.indexOf('PEREMPUAN') !== -1 || rawGender.indexOf('WANITA') !== -1 || rawGender === 'P' || rawGender === 'F' || rawGender.indexOf('FEMALE') !== -1) {
          gender = 'PEREMPUAN';
        } else if (rawGender.indexOf('LAKI') !== -1 || rawGender.indexOf('PRIA') !== -1 || rawGender === 'L' || rawGender === 'M' || rawGender.indexOf('MALE') !== -1) {
          gender = 'LAKI-LAKI';
        } else if (rawGender) {
          gender = rawGender.indexOf('P') === 0 ? 'PEREMPUAN' : 'LAKI-LAKI';
        }

        const edu = idxPend !== -1 ? (String(row[idxPend] || 'LAINNYA').trim().toUpperCase() || 'LAINNYA') : 'LAINNYA';
        const religion = idxAgama !== -1 ? (String(row[idxAgama] || 'KRISTEN').trim().toUpperCase() || 'KRISTEN') : 'KRISTEN';
        const empType = sInfo.typeTag;

        genderDist[gender] = (genderDist[gender] || 0) + 1;
        eduDist[edu] = (eduDist[edu] || 0) + 1;
        religionDist[religion] = (religionDist[religion] || 0) + 1;
        empTypeDist[empType] = (empTypeDist[empType] || 0) + 1;
        jabatanDist[catJabatan] = (jabatanDist[catJabatan] || 0) + 1;

        const gDepan = idxGelarD !== -1 && row[idxGelarD] ? String(row[idxGelarD]).trim() + ' ' : '';
        const gBelakang = idxGelarB !== -1 && row[idxGelarB] ? ', ' + String(row[idxGelarB]).trim() : '';
        const namaLengkap = gDepan + cellNama + gBelakang;

        let displaySatker = (mappedPD && mappedPD !== 'Lainnya') ? mappedPD : (rawCellSatker || 'Pemerintah Kab. Toba');
        let displayUnor = rawCellUnor || rawCellSatker || '-';

        if (
          displaySatker.toUpperCase().indexOf("PERENCANAAN PEMBANGUNAN") !== -1 ||
          displaySatker.toUpperCase().indexOf("PENELITIAN") !== -1 ||
          displaySatker.toUpperCase().indexOf("BAPPEDA") !== -1 ||
          displaySatker.toUpperCase().indexOf("BAPELITBANGDA") !== -1 ||
          displaySatker.toUpperCase().indexOf("RISET DAN INOVASI") !== -1
        ) {
          if (displaySatker.toUpperCase().indexOf("SEKRETARIAT") === -1) {
            displaySatker = "Badan Perencanaan Pembangunan, Riset dan Inovasi Daerah";
          }
        }

        if (
          displayUnor.toUpperCase().indexOf("PERENCANAAN PEMBANGUNAN") !== -1 ||
          displayUnor.toUpperCase().indexOf("PENELITIAN") !== -1 ||
          displayUnor.toUpperCase().indexOf("BAPPEDA") !== -1 ||
          displayUnor.toUpperCase().indexOf("BAPELITBANGDA") !== -1 ||
          displayUnor.toUpperCase().indexOf("RISET DAN INOVASI") !== -1
        ) {
          if (displayUnor.toUpperCase().indexOf("SEKRETARIAT") === -1) {
            displayUnor = normalizeOpdName(displayUnor);
          }
        }

        allFilteredRows.push({
          id: idxPnsId !== -1 && row[idxPnsId] ? String(row[idxPnsId]).trim() : ('PGW-' + (allFilteredRows.length + 1)),
          nip: cellNip || '-',
          nama: namaLengkap || 'Tanpa Nama',
          golongan: cellGol || '-',
          unorNama: displayUnor,
          jabatanNama: cellJabatanNama || '-',
          jenisJabatan: rawCellJenisJabatan || jabMeta.label,
          kategoriJabatan: catJabatan, // 'STRUKTURAL' | 'FUNGSIONAL' | 'PELAKSANA'
          jfCategory: jabMeta.jfCategory,
          jfLevel: jabMeta.jfLevel,
          eselon: cellEselon || '-',
          pendidikan: cellPend || '-',
          satker: displaySatker,
          jenisPeg: empType,
          kelamin: gender
        });
      }
    }

    const paginatedRows = allFilteredRows.slice(offset, offset + limit);

    return {
      success: true,
      data: paginatedRows,
      total: allFilteredRows.length,
      chartData: {
        gender: genderDist,
        education: eduDist,
        religion: religionDist,
        empType: empTypeDist,
        jabatan: jabatanDist
      }
    };
  } catch (error) {
    Logger.log("Error getPegawaiDashboardData: " + error.toString());
    return { success: false, data: [], total: 0, chartData: {}, error: error.toString() };
  }
}

function getPegawaiDropdownOptions() {
  return getFilterOptions();
}

function getPegawaiList(filters, limit, offset) {
  return getPegawaiDashboardData(filters, limit, offset);
}

function getPegawaiSummaryStats() {
  const res = getPegawaiDashboardData({}, 1, 0);
  if (res && res.success && res.chartData) {
    const c = res.chartData;
    return {
      success: true,
      data: {
        totalPegawai: res.total || 0,
        totalPns: (c.empType && c.empType['PNS']) || 0,
        totalPppk: (c.empType && c.empType['PPPK']) || 0,
        totalPppkPw: (c.empType && c.empType['PPPK_PARUH_WAKTU']) || 0
      }
    };
  }
  return {
    success: true,
    data: { totalPegawai: 0, totalPns: 0, totalPppk: 0, totalPppkPw: 0 }
  };
}