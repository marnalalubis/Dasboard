/**
 * =========================================================================
 * ACRU EMPLOYEE ANALYTICS SYSTEM - KABUPATEN TOBA
 * File: strukturService.gs
 * Deskripsi: Backend Service Manajemen Formasi Jabatan Fungsional (JF),
 * Gap Analysis ABK, Bezetting Pegawai, dan Formasi Jabatan Pelaksana (JP).
 * =========================================================================
 */

/**
 * Mesin Pencari Tab Master Formasi Ultra-Resilient
 */
function resolveSheetJFSafe_(ss) {
  if (!ss) return null;
  var sh = ss.getSheetByName('Data_Master_Kebutuhan_JF');
  if (sh) return sh;

  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var clean = sheets[i].getName().toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (clean === 'DATAMASTERKEBUTUHANJF' || clean.indexOf('KEBUTUHANJF') !== -1) {
      return sheets[i];
    }
  }
  return null;
}


/**
 * Deteksi Kategori JF (Keahlian vs Keterampilan) & Jenjang Jabatan
 */
function detectJfCategoryAndLevel(rawJabatan, rawJenjang) {
  var jab = String(rawJabatan || '').toUpperCase().trim();
  var jnj = String(rawJenjang || '').toUpperCase().trim();

  var category = 'KEAHLIAN';
  var level = 'Ahli Pertama';

  if (jnj.indexOf('UTAMA') !== -1 || jab.indexOf('AHLI UTAMA') !== -1) {
    category = 'KEAHLIAN'; level = 'Ahli Utama';
  } else if (jnj.indexOf('MADYA') !== -1 || jab.indexOf('AHLI MADYA') !== -1 || jab.indexOf('MADYA') !== -1) {
    category = 'KEAHLIAN'; level = 'Ahli Madya';
  } else if (jnj.indexOf('MUDA') !== -1 || jab.indexOf('AHLI MUDA') !== -1 || jab.indexOf('MUDA') !== -1) {
    category = 'KEAHLIAN'; level = 'Ahli Muda';
  } else if (jnj.indexOf('PERTAMA') !== -1 || jab.indexOf('AHLI PERTAMA') !== -1 || jab.indexOf('PERTAMA') !== -1) {
    category = 'KEAHLIAN'; level = 'Ahli Pertama';
  } else if (jnj.indexOf('PENYELIA') !== -1 || jab.indexOf('PENYELIA') !== -1) {
    category = 'KETERAMPILAN'; level = 'Penyelia';
  } else if (jnj.indexOf('MAHIR') !== -1 || jab.indexOf('MAHIR') !== -1 || jab.indexOf('PELAKSANA LANJUTAN') !== -1) {
    category = 'KETERAMPILAN'; level = 'Mahir';
  } else if (jnj.indexOf('TERAMPIL') !== -1 || jab.indexOf('TERAMPIL') !== -1 || jab.indexOf('PELAKSANA') !== -1) {
    category = 'KETERAMPILAN'; level = 'Terampil';
  } else if (jnj.indexOf('PEMULA') !== -1 || jab.indexOf('PEMULA') !== -1) {
    category = 'KETERAMPILAN'; level = 'Pemula';
  } else {
    if (jab.indexOf('BIDAN TERAMPIL') !== -1 || jab.indexOf('PERAWAT TERAMPIL') !== -1) {
      category = 'KETERAMPILAN'; level = 'Terampil';
    } else {
      category = 'KEAHLIAN'; level = 'Ahli Pertama';
    }
  }

  return { category: category, level: level };
}

/**
 * =========================================================================
 * 1. BEZETTING & KEBUTUHAN FORMASI ABK JABATAN FUNGSIONAL (JF)
 * =========================================================================
 */
function getBezettingJFData(filters) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Spreadsheet aktif tidak ditemukan.");

    filters = filters || {};
    var filterSearch = String(filters.search || '').toLowerCase().trim();
    var filterStatus = String(filters.statusKebutuhan || 'ALL').toUpperCase().trim();
    var filterKategori = String(filters.kategori || 'ALL').toUpperCase().trim();
    var filterJenjang = String(filters.jenjang || 'ALL').trim();
    var filterOpd = String(filters.opd || 'ALL').trim();

    // -------------------------------------------------------------
    // TAHAP 1: Baca Data Master Kebutuhan Formasi ABK JF
    // -------------------------------------------------------------
    var sheetJF = resolveSheetJFSafe_(ss);
    var masterJfList = [];
    var masterJfMap = {};
    var grandMasterKebutuhan = 0;

    if (sheetJF && sheetJF.getLastRow() > 1) {
      var dataMaster = sheetJF.getDataRange().getDisplayValues();
      var hMaster = dataMaster[0];

      var idxMId = -1, idxMOpd = -1, idxMJab = -1, idxMJenjang = -1, idxMKeb = -1, idxMSumber = -1, idxMKet = -1;

      for (var c = 0; c < hMaster.length; c++) {
        var hClean = String(hMaster[c] || '').toUpperCase().trim().replace(/[^A-Z0-9]/g, '');
        if (!hClean) continue;

        if (idxMId === -1 && (hClean === 'IDKEBUTUHAN' || hClean === 'ID' || hClean === 'KODE')) {
          idxMId = c;
        } else if (idxMOpd === -1 && (hClean.indexOf('PERANGKATDAERAH') !== -1 || hClean.indexOf('OPD') !== -1 || hClean.indexOf('SATKER') !== -1)) {
          idxMOpd = c;
        } else if (idxMJab === -1 && (hClean === 'JABATANNAMA' || hClean === 'NAMAJABATAN' || hClean === 'JABATAN' || hClean === 'NOMENKLATUR')) {
          idxMJab = c;
        } else if (idxMJenjang === -1 && (hClean.indexOf('JENJANG') !== -1 || hClean === 'TINGKAT')) {
          idxMJenjang = c;
        } else if (idxMKeb === -1 && hClean.indexOf('ID') === -1 && (hClean.indexOf('KEBUTUHAN') !== -1 || hClean.indexOf('FORMASI') !== -1 || hClean.indexOf('KUOTA') !== -1 || hClean.indexOf('JUMLAH') !== -1 || hClean.indexOf('KEBL') !== -1)) {
          idxMKeb = c;
        } else if (idxMSumber === -1 && (hClean.indexOf('SUMBER') !== -1 || hClean.indexOf('ANALISIS') !== -1)) {
          idxMSumber = c;
        } else if (idxMKet === -1 && (hClean.indexOf('KETERANGAN') !== -1 || hClean.indexOf('CATATAN') !== -1)) {
          idxMKet = c;
        }
      }

      // Fallback indeks kolom standar bila header tidak standar
      if (idxMOpd === -1 && hMaster.length > 1) idxMOpd = 1;
      if (idxMJab === -1 && hMaster.length > 2) idxMJab = 2;
      if (idxMJenjang === -1 && hMaster.length > 3) idxMJenjang = 3;
      if (idxMKeb === -1 && hMaster.length > 4) idxMKeb = 4;

      for (var m = 1; m < dataMaster.length; m++) {
        var mRow = dataMaster[m];
        if (mRow.every(function(cell) { return String(cell).trim() === ''; })) continue;

        var rawOpd = idxMOpd !== -1 ? String(mRow[idxMOpd] || '').trim() : '';
        var rawJab = idxMJab !== -1 ? String(mRow[idxMJab] || '').trim() : '';
        var rawJnj = idxMJenjang !== -1 ? String(mRow[idxMJenjang] || '').trim() : '';
        var rawKebVal = idxMKeb !== -1 ? mRow[idxMKeb] : 0;
        var rawSumber = idxMSumber !== -1 ? String(mRow[idxMSumber] || '').trim() : '';
        var rawKet = idxMKet !== -1 ? String(mRow[idxMKet] || '').trim() : '';

        if (!rawOpd && !rawJab) continue;

        var normOpd = typeof normalizeOpdName === 'function' ? normalizeOpdName(rawOpd) : rawOpd;
        var parsedKeb = parseInt(String(rawKebVal || '0').replace(/[^0-9]/g, ''), 10) || 0;
        var catLevel = detectJfCategoryAndLevel(rawJab, rawJnj);

        var masterItem = {
          id: idxMId !== -1 ? String(mRow[idxMId] || '').trim() : ('KB-JF-' + m),
          opd: normOpd,
          jabatan: rawJab,
          jenjang: catLevel.level,
          kategori: catLevel.category,
          kebutuhan: parsedKeb,
          sumber: rawSumber,
          keterangan: rawKet
        };

        masterJfList.push(masterItem);
        grandMasterKebutuhan += parsedKeb;

        // Kunci lookup ganda untuk rekonsiliasi presisi
        var keyOpdJab = cleanMatchKey(normOpd) + '___' + cleanMatchKey(rawJab);
        var keyOpdJabLvl = keyOpdJab + '___' + cleanMatchKey(catLevel.level);
        masterJfMap[keyOpdJab] = masterItem;
        masterJfMap[keyOpdJabLvl] = masterItem;
      }
    }

    // -------------------------------------------------------------
    // TAHAP 2: Agregasi Bezetting Pegawai JF (PNS, PPPK, & PPPK PW)
    // -------------------------------------------------------------
    var targetSheets = [
      { name: 'Data_Pegawai_PNS', statusTag: 'PNS' },
      { name: 'Data_Pegawai_PPPK', statusTag: 'PPPK' },
      { name: 'Data_Pegawai_PPPK_Paruh_Waktu', statusTag: 'PW' }
    ];

    var bezettingMap = {};

    for (var s = 0; s < targetSheets.length; s++) {
      var shMeta = targetSheets[s];
      var sheetObj = typeof getSheetByNameRobust === 'function'
        ? getSheetByNameRobust(ss, shMeta.name)
        : ss.getSheetByName(shMeta.name);

      if (!sheetObj || sheetObj.getLastRow() <= 1) continue;

      var rawData = sheetObj.getDataRange().getDisplayValues();
      var headers = rawData[0];

      var idxNip = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["NIP BARU", "NIP", "NIP_BARU"]) : -1;
      var idxNama = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["NAMA", "NAMA LENGKAP"]) : -1;
      var idxGelarD = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["GELAR DEPAN"]) : -1;
      var idxGelarB = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["GELAR BELAKANG"]) : -1;
      var idxGol = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["GOL AKHIR NAMA", "GOLONGAN", "GOL"]) : -1;
      var idxPend = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["TINGKAT PENDIDIKAN NAMA", "PENDIDIKAN"]) : -1;
      var idxSatker = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["SATUAN KERJA KERJA NAMA", "SATKER"]) : -1;
      var idxSatkerInduk = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK"]) : -1;
      var idxUnor = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["UNOR NAMA", "UNOR"]) : -1;
      var idxJenisJab = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JENIS JABATAN NAMA", "JENIS JABATAN"]) : -1;
      var idxJabNama = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JABATAN NAMA", "JABATAN"]) : -1;
      var idxEselon = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["ESELON NAMA", "ESELON"]) : -1;

      for (var r = 1; r < rawData.length; r++) {
        var row = rawData[r];
        if (row.every(function(cell) { return String(cell).trim() === ''; })) continue;

        var rawCellJenisJab = idxJenisJab !== -1 ? String(row[idxJenisJab] || '').trim() : '';
        var rawCellJabNama = idxJabNama !== -1 ? String(row[idxJabNama] || '').trim() : '';
        var rawCellEselon = idxEselon !== -1 ? String(row[idxEselon] || '').trim() : '';

        var jabMeta = typeof categorizePegawaiJabatanLocal === 'function'
          ? categorizePegawaiJabatanLocal(rawCellJenisJab, rawCellJabNama, rawCellEselon)
          : detectJfCategoryAndLevel(rawCellJabNama, '');

        if (jabMeta.type && jabMeta.type !== 'FUNGSIONAL') {
          continue; // Lewati struktural dan pelaksana
        }

        var sInduk = idxSatkerInduk !== -1 ? String(row[idxSatkerInduk] || '').trim() : '';
        var sKerja = idxSatker !== -1 ? String(row[idxSatker] || '').trim() : '';
        var sUnor = idxUnor !== -1 ? String(row[idxUnor] || '').trim() : '';

        var targetOpd = '';
        if (typeof mapUnorToTarget === 'function') {
          targetOpd = mapUnorToTarget(sInduk, sKerja, sUnor);
        } else {
          targetOpd = sKerja || sInduk || sUnor || 'Perangkat Daerah';
        }
        targetOpd = typeof normalizeOpdName === 'function' ? normalizeOpdName(targetOpd) : targetOpd;

        var nip = idxNip !== -1 ? String(row[idxNip] || '-').trim() : '-';
        var namaMurni = idxNama !== -1 ? String(row[idxNama] || 'Pegawai').trim() : 'Pegawai';
        var gd = idxGelarD !== -1 && row[idxGelarD] ? String(row[idxGelarD]).trim() + ' ' : '';
        var gb = idxGelarB !== -1 && row[idxGelarB] ? ', ' + String(row[idxGelarB]).trim() : '';
        var fullNama = gd + namaMurni + gb;
        var gol = idxGol !== -1 ? String(row[idxGol] || '-').trim() : '-';
        var pend = idxPend !== -1 ? String(row[idxPend] || '-').trim() : '-';

        var key = cleanMatchKey(targetOpd) + '___' + cleanMatchKey(rawCellJabNama);

        if (!bezettingMap[key]) {
          var catLvl = detectJfCategoryAndLevel(rawCellJabNama, jabMeta.jfLevel || '');
          bezettingMap[key] = {
            opd: targetOpd,
            jabatan: rawCellJabNama,
            kategori: jabMeta.jfCategory || catLvl.category,
            jenjang: jabMeta.jfLevel || catLvl.level,
            pns: 0,
            pppk: 0,
            p3k_pw: 0,
            pegawaiList: []
          };
        }

        if (shMeta.statusTag === 'PNS') bezettingMap[key].pns++;
        else if (shMeta.statusTag === 'PPPK') bezettingMap[key].pppk++;
        else bezettingMap[key].p3k_pw++;

        bezettingMap[key].pegawaiList.push({
          nip: nip,
          nama: fullNama,
          golongan: gol,
          pendidikan: pend,
          statusAsn: shMeta.statusTag
        });
      }
    }

    // -------------------------------------------------------------
    // TAHAP 3: Full Outer Merge (Kebutuhan Master ABK + Bezetting Riil)
    // -------------------------------------------------------------
    var combinedMap = {};

    // 1. Masukkan semua formasi dari Master Kebutuhan ABK
    for (var i = 0; i < masterJfList.length; i++) {
      var itemM = masterJfList[i];
      var mKey = cleanMatchKey(itemM.opd) + '___' + cleanMatchKey(itemM.jabatan);

      var bez = bezettingMap[mKey];
      var pnsCount = bez ? bez.pns : 0;
      var pppkCount = bez ? bez.pppk : 0;
      var pwCount = bez ? bez.p3k_pw : 0;
      var totBez = pnsCount + pppkCount + pwCount;

      combinedMap[mKey] = {
        perangkatDaerah: itemM.opd,
        jabatanNama: itemM.jabatan,
        jfCategory: itemM.kategori,
        jfLevel: itemM.jenjang,
        pns: pnsCount,
        pppk: pppkCount,
        p3k_pw: pwCount,
        totalBezetting: totBez,
        kebutuhan: itemM.kebutuhan,
        sumber: itemM.sumber,
        keterangan: itemM.keterangan,
        pegawaiList: bez ? bez.pegawaiList : []
      };
    }

    // 2. Gabungkan jabatan riil yang belum ada di spreadsheet master ABK
    for (var bKey in bezettingMap) {
      if (!combinedMap[bKey]) {
        var itemB = bezettingMap[bKey];
        var totBezB = itemB.pns + itemB.pppk + itemB.p3k_pw;
        var matchInMaster = masterJfMap[bKey] || null;

        combinedMap[bKey] = {
          perangkatDaerah: itemB.opd,
          jabatanNama: itemB.jabatan,
          jfCategory: itemB.kategori,
          jfLevel: itemB.jenjang,
          pns: itemB.pns,
          pppk: itemB.pppk,
          p3k_pw: itemB.p3k_pw,
          totalBezetting: totBezB,
          kebutuhan: matchInMaster ? matchInMaster.kebutuhan : 0,
          sumber: matchInMaster ? matchInMaster.sumber : '',
          keterangan: matchInMaster ? matchInMaster.keterangan : '',
          pegawaiList: itemB.pegawaiList
        };
      }
    }

    // -------------------------------------------------------------
    // TAHAP 4: Hitung Gap, Status Formasi, & Filter Antarmuka
    // -------------------------------------------------------------
    var resultRows = [];
    var totalJfPegawai = 0;
    var totalKebutuhanFormasi = 0;
    var totalDefisitPos = 0;
    var totalSurplusPos = 0;
    var totalGapPegawai = 0;

    for (var key in combinedMap) {
      var rObj = combinedMap[key];

      var gap = 0;
      var statusKebutuhan = 'BELUM_DIATUR';

      if (rObj.kebutuhan > 0) {
        gap = rObj.totalBezetting - rObj.kebutuhan;
        if (gap < 0) {
          statusKebutuhan = 'DEFISIT';
          totalDefisitPos++;
          totalGapPegawai += Math.abs(gap);
        } else if (gap > 0) {
          statusKebutuhan = 'SURPLUS';
          totalSurplusPos++;
        } else {
          statusKebutuhan = 'IDEAL';
        }
      }

      rObj.gap = gap;
      rObj.statusKebutuhan = statusKebutuhan;
      rObj.statusLabel = statusKebutuhan === 'DEFISIT' ? ('Kurang ' + Math.abs(gap)) : (statusKebutuhan === 'SURPLUS' ? ('Lebih +' + gap) : (statusKebutuhan === 'IDEAL' ? 'Ideal' : 'Belum Diatur'));

      // Filter OPD
      if (filterOpd && filterOpd !== 'ALL') {
        var cleanFilterOpd = cleanMatchKey(filterOpd);
        var cleanRowOpd = cleanMatchKey(rObj.perangkatDaerah);
        if (cleanRowOpd !== cleanFilterOpd && cleanRowOpd.indexOf(cleanFilterOpd) === -1 && cleanFilterOpd.indexOf(cleanRowOpd) === -1) {
          continue;
        }
      }

      // Filter Status Formasi
      if (filterStatus && filterStatus !== 'ALL') {
        if (rObj.statusKebutuhan !== filterStatus) continue;
      }

      // Filter Kategori
      if (filterKategori && filterKategori !== 'ALL') {
        if (String(rObj.jfCategory).toUpperCase() !== filterKategori) continue;
      }

      // Filter Jenjang
      if (filterJenjang && filterJenjang !== 'ALL') {
        if (String(rObj.jfLevel).toUpperCase() !== filterJenjang.toUpperCase()) continue;
      }

      // Filter Pencarian Teks
      if (filterSearch) {
        var matchSearch = rObj.jabatanNama.toLowerCase().indexOf(filterSearch) !== -1 ||
                          rObj.perangkatDaerah.toLowerCase().indexOf(filterSearch) !== -1;

        if (!matchSearch && rObj.pegawaiList.length > 0) {
          matchSearch = rObj.pegawaiList.some(function(p) {
            return p.nama.toLowerCase().indexOf(filterSearch) !== -1 ||
                   p.nip.toLowerCase().indexOf(filterSearch) !== -1;
          });
        }
        if (!matchSearch) continue;
      }

      totalJfPegawai += rObj.totalBezetting;
      totalKebutuhanFormasi += rObj.kebutuhan;
      resultRows.push(rObj);
    }

    // Urutkan alfabetis: Perangkat Daerah lalu Nama Jabatan
    resultRows.sort(function(a, b) {
      var cmpOpd = a.perangkatDaerah.localeCompare(b.perangkatDaerah);
      if (cmpOpd !== 0) return cmpOpd;
      return a.jabatanNama.localeCompare(b.jabatanNama);
    });

    // Gunakan total master aktual jika tanpa filter instansi
    var finalKpiKebutuhan = (filterOpd && filterOpd !== 'ALL') ? totalKebutuhanFormasi : (grandMasterKebutuhan > 0 ? grandMasterKebutuhan : totalKebutuhanFormasi);

    return {
      success: true,
      data: resultRows,
      stats: {
        totalJF: totalJfPegawai,
        totalKebutuhan: finalKpiKebutuhan,
        totalDefisitPos: totalDefisitPos,
        totalSurplusPos: totalSurplusPos,
        totalGapPegawai: totalGapPegawai
      }
    };
  } catch (error) {
    Logger.log("Error getBezettingJFData: " + error.toString());
    return { success: false, data: [], stats: { totalJF: 0, totalKebutuhan: 0, totalDefisitPos: 0, totalSurplusPos: 0, totalGapPegawai: 0 }, error: error.toString() };
  }
}

/**
 * Simpan / Sesuaikan Kebutuhan ABK JF dari Modal Pop-up
 */
function saveFormasiKebutuhanJF(payload) {
  try {
    if (!payload || !payload.opd || !payload.jabatan) {
      throw new Error("Data Perangkat Daerah dan Nomenklatur Jabatan wajib diisi.");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = resolveSheetJFSafe_(ss);
    if (!sheet) throw new Error("Sheet Data_Master_Kebutuhan_JF tidak ditemukan.");

    var normOpd = typeof normalizeOpdName === 'function' ? normalizeOpdName(payload.opd) : payload.opd;
    var normJab = String(payload.jabatan).trim();
    var jnj = String(payload.jenjang || 'Ahli Pertama').trim();
    var keb = parseInt(payload.kebutuhan, 10) || 1;
    var sumber = String(payload.sumber || 'Anjab & ABK Pemkab Toba').trim();
    var ket = String(payload.keterangan || '').trim();

    var targetKey = cleanMatchKey(normOpd) + '___' + cleanMatchKey(normJab);
    var targetKeyWithLvl = targetKey + '___' + cleanMatchKey(jnj);
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];

    var idxOpd = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["PERANGKAT_DAERAH", "OPD"]) : 1;
    var idxJab = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JABATAN_NAMA", "JABATAN"]) : 2;
    var idxJnj = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JENJANG"]) : 3;
    var idxKeb = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JUMLAH_KEBUTUHAN", "JUMLAH_FORMASI", "KEBUTUHAN"]) : 4;
    var idxSumber = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["SUMBER_ANALISIS", "SUMBER"]) : 5;
    var idxKet = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["KETERANGAN"]) : 6;

    var foundRow = -1;
    var fallbackRow = -1;
    for (var i = 1; i < data.length; i++) {
      var curOpd = idxOpd !== -1 ? data[i][idxOpd] : '';
      var curJab = idxJab !== -1 ? data[i][idxJab] : '';
      var curJnj = idxJnj !== -1 ? data[i][idxJnj] : '';
      var curKey = cleanMatchKey(curOpd) + '___' + cleanMatchKey(curJab);
      var curKeyWithLvl = curKey + '___' + cleanMatchKey(curJnj);

      if (curKeyWithLvl === targetKeyWithLvl) {
        foundRow = i + 1;
        break;
      }
      if (curKey === targetKey && fallbackRow === -1) {
        fallbackRow = i + 1;
      }
    }
    if (foundRow === -1 && fallbackRow !== -1) {
      foundRow = fallbackRow;
    }

    if (foundRow !== -1) {
      if (idxJnj !== -1) sheet.getRange(foundRow, idxJnj + 1).setValue(jnj);
      if (idxKeb !== -1) sheet.getRange(foundRow, idxKeb + 1).setValue(keb);
      if (idxSumber !== -1 && sumber) sheet.getRange(foundRow, idxSumber + 1).setValue(sumber);
      if (idxKet !== -1 && ket) sheet.getRange(foundRow, idxKet + 1).setValue(ket);
    } else {
      var newId = 'KB-JF-' + String(sheet.getLastRow()).padStart(3, '0');
      var newRow = [newId, normOpd, normJab, jnj, keb, sumber, ket];
      sheet.appendRow(newRow);
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Formasi ABK JF berhasil disimpan dan disinkronkan!" };
  } catch (err) {
    Logger.log("Error saveFormasiKebutuhanJF: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * Simpan / Sesuaikan Kebutuhan ABK JP (Jabatan Pelaksana) dari Modal Pop-up
 */
function saveFormasiKebutuhanJP(payload) {
  try {
    if (!payload || !payload.opd || !payload.jabatan) {
      throw new Error("Data Perangkat Daerah dan Nomenklatur Jabatan wajib diisi.");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = typeof getSheetByNameRobust === 'function'
      ? getSheetByNameRobust(ss, 'Data_Master_Kebutuhan_JP')
      : ss.getSheetByName('Data_Master_Kebutuhan_JP');
    if (!sheet) throw new Error("Sheet Data_Master_Kebutuhan_JP tidak ditemukan.");

    var normOpd = typeof normalizeOpdName === 'function' ? normalizeOpdName(payload.opd) : payload.opd;
    var normJab = String(payload.jabatan).trim();
    var keb = parseInt(payload.kebutuhan, 10) || 0;
    var sumber = String(payload.sumber || 'Anjab & ABK Pemkab Toba').trim();
    var ket = String(payload.keterangan || '').trim();

    var targetKey = cleanMatchKey(normOpd) + '___' + cleanMatchKey(normJab);
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];

    var idxOpd = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["PERANGKAT_DAERAH", "OPD"]) : 1;
    var idxJab = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JABATAN_NAMA", "JABATAN"]) : 2;
    var idxKeb = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JUMLAH_KEBUTUHAN", "JUMLAH_FORMASI", "KEBUTUHAN"]) : 3;
    var idxSumber = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["SUMBER_ANALISIS", "SUMBER"]) : 4;
    var idxKet = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["KETERANGAN", "KET"]) : 5;

    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      var curOpd = idxOpd !== -1 ? data[i][idxOpd] : '';
      var curJab = idxJab !== -1 ? data[i][idxJab] : '';
      var curKey = cleanMatchKey(curOpd) + '___' + cleanMatchKey(curJab);

      if (curKey === targetKey) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow !== -1) {
      if (idxKeb !== -1) sheet.getRange(foundRow, idxKeb + 1).setValue(keb);
      if (idxSumber !== -1 && sumber) sheet.getRange(foundRow, idxSumber + 1).setValue(sumber);
      if (idxKet !== -1 && ket) sheet.getRange(foundRow, idxKet + 1).setValue(ket);
    } else {
      var newId = 'KB-JP-' + String(sheet.getLastRow()).padStart(3, '0');
      var newRow = [newId, normOpd, normJab, keb, sumber, ket];
      sheet.appendRow(newRow);
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Formasi ABK Jabatan Pelaksana berhasil disimpan dan disinkronkan!" };
  } catch (err) {
    Logger.log("Error saveFormasiKebutuhanJP: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

function generateExportBezettingJFExcel(filters) {
  try {
    var res = getBezettingJFData(filters);
    if (!res || !res.success) throw new Error("Gagal mengambil data bezetting JF.");

    var rows = res.data || [];
    var stats = res.stats || {};
    var tanggalCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var dateFileStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');

    var rowsXml = '';
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var gapStr = r.gap > 0 ? ('+' + r.gap) : (r.kebutuhan > 0 ? r.gap : '-');
      var statusColor = r.statusKebutuhan === 'DEFISIT' ? '#ef4444' : (r.statusKebutuhan === 'SURPLUS' ? '#f59e0b' : (r.statusKebutuhan === 'IDEAL' ? '#10b981' : '#64748b'));

      rowsXml += '<tr>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;">' + (i + 1) + '</td>' +
        '<td style="border:0.5pt solid #cccccc;font-weight:bold;">' + r.perangkatDaerah + '</td>' +
        '<td style="border:0.5pt solid #cccccc;">' + r.jabatanNama + '</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;">' + r.jfCategory + '</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;">' + r.jfLevel + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + r.pns + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + r.pppk + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + r.p3k_pw + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background:#e0f2fe;" mso-number-format="#,##0">' + r.totalBezetting + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background:#fef3c7;" mso-number-format="#,##0">' + (r.kebutuhan > 0 ? r.kebutuhan : '-') + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;">' + gapStr + '</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;color:' + statusColor + ';font-weight:bold;">' + r.statusLabel + '</td>' +
      '</tr>';
    }

    var excelHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head>' +
        '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">' +
        '<style>' +
          'body { font-family: Calibri, Arial, sans-serif; font-size: 10pt; }' +
          '.title { font-size: 13pt; font-weight: bold; color: #0f172a; }' +
          '.th-col { background-color: #0f172a; color: #ffffff; font-weight: bold; text-align: center; border: 0.5pt solid #000; }' +
          '.th-total { background-color: #0f172a; color: #00ffaa; font-weight: bold; text-align: right; border: 0.5pt solid #000; }' +
        '</style>' +
      '</head>' +
      '<body>' +
        '<table>' +
          '<tr><td colspan="12" class="title">PEMERINTAH KABUPATEN TOBA</td></tr>' +
          '<tr><td colspan="12" style="font-size:11pt;font-weight:bold;">BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA</td></tr>' +
          '<tr><td colspan="12" style="color:#64748b;font-style:italic;">Matriks Bezetting & Gap Analysis Formasi Jabatan Fungsional (JF) | Tanggal: ' + tanggalCetak + ' WIB</td></tr>' +
          '<tr></tr>' +
          '<tr>' +
            '<th class="th-col" style="width:40px;">NO</th>' +
            '<th class="th-col" style="width:280px;">PERANGKAT DAERAH</th>' +
            '<th class="th-col" style="width:280px;">NOMENKLATUR JABATAN FUNGSIONAL</th>' +
            '<th class="th-col" style="width:110px;">KATEGORI</th>' +
            '<th class="th-col" style="width:110px;">JENJANG</th>' +
            '<th class="th-col" style="width:65px;">PNS</th>' +
            '<th class="th-col" style="width:65px;">PPPK</th>' +
            '<th class="th-col" style="width:65px;">PW</th>' +
            '<th class="th-col" style="width:90px;background:#0284c7;">BEZETTING</th>' +
            '<th class="th-col" style="width:90px;background:#d97706;">ABK</th>' +
            '<th class="th-col" style="width:75px;">GAP</th>' +
            '<th class="th-col" style="width:120px;">STATUS FORMASI</th>' +
          '</tr>' +
          rowsXml +
          '<tr>' +
            '<td colspan="5" style="background-color:#0f172a;color:#ffffff;font-weight:bold;text-align:center;border:0.5pt solid #000;">TOTAL KESELURUHAN</td>' +
            '<td class="th-total" colspan="4" mso-number-format="#,##0">' + stats.totalJF + '</td>' +
            '<td class="th-total" mso-number-format="#,##0">' + stats.totalKebutuhan + '</td>' +
            '<td class="th-total" colspan="2" style="color:#f43f5e;">Defisit: ' + stats.totalGapPegawai + ' Pegawai</td>' +
          '</tr>' +
        '</table>' +
      '</body>' +
      '</html>';

    var blob = Utilities.newBlob(excelHtml, 'application/vnd.ms-excel', 'Bezetting_Gap_Jabatan_Fungsional_Toba_' + dateFileStr + '.xls');
    return {
      success: true,
      filename: blob.getName(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    Logger.log("Error generateExportBezettingJFExcel: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * =========================================================================
 * 2. BEZETTING & KEBUTUHAN FORMASI ABK JABATAN PELAKSANA (JP)
 * =========================================================================
 */
function getBezettingPelaksanaData(filters) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Spreadsheet aktif tidak ditemukan.");

    filters = filters || {};
    var filterSearch = String(filters.search || '').toLowerCase().trim();
    var filterStatus = String(filters.statusKebutuhan || 'ALL').toUpperCase().trim();
    var filterOpd = String(filters.opd || 'ALL').trim();

    var sheetJP = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, 'Data_Master_Kebutuhan_JP') : ss.getSheetByName('Data_Master_Kebutuhan_JP');
    var masterJpList = [];
    var masterJpMap = {};

    if (sheetJP && sheetJP.getLastRow() > 1) {
      var dataMaster = sheetJP.getDataRange().getDisplayValues();
      var hMaster = dataMaster[0];

      var idxMOpd = typeof findColumnIndex === 'function' ? findColumnIndex(hMaster, ["PERANGKAT_DAERAH", "OPD"]) : 1;
      var idxMJab = typeof findColumnIndex === 'function' ? findColumnIndex(hMaster, ["JABATAN_NAMA", "JABATAN"]) : 2;
      var idxMKeb = typeof findColumnIndex === 'function' ? findColumnIndex(hMaster, ["JUMLAH_KEBUTUHAN", "JUMLAH_FORMASI", "KEBUTUHAN"]) : 3;
      var idxMSumber = typeof findColumnIndex === 'function' ? findColumnIndex(hMaster, ["SUMBER_ANALISIS", "SUMBER"]) : 4;
      var idxMKet = typeof findColumnIndex === 'function' ? findColumnIndex(hMaster, ["KETERANGAN", "KET"]) : 5;

      for (var m = 1; m < dataMaster.length; m++) {
        var mRow = dataMaster[m];
        if (mRow.every(function(cell) { return String(cell).trim() === ''; })) continue;

        var rawOpd = idxMOpd !== -1 ? String(mRow[idxMOpd] || '').trim() : '';
        var rawJab = idxMJab !== -1 ? String(mRow[idxMJab] || '').trim() : '';
        var rawKeb = idxMKeb !== -1 ? mRow[idxMKeb] : 0;
        var rawSumber = idxMSumber !== -1 ? String(mRow[idxMSumber] || '').trim() : '';
        var rawKet = idxMKet !== -1 ? String(mRow[idxMKet] || '').trim() : '';

        if (!rawOpd && !rawJab) continue;

        var normOpd = typeof normalizeOpdName === 'function' ? normalizeOpdName(rawOpd) : rawOpd;
        var parsedKeb = parseInt(String(rawKeb || '0').replace(/[^0-9]/g, ''), 10) || 0;

        var item = { opd: normOpd, jabatan: rawJab, kebutuhan: parsedKeb, sumber: rawSumber, keterangan: rawKet };
        masterJpList.push(item);
        masterJpMap[cleanMatchKey(normOpd) + '___' + cleanMatchKey(rawJab)] = item;
      }
    }

    var targetSheets = [
      { name: 'Data_Pegawai_PNS', statusTag: 'PNS' },
      { name: 'Data_Pegawai_PPPK', statusTag: 'PPPK' },
      { name: 'Data_Pegawai_PPPK_Paruh_Waktu', statusTag: 'PW' }
    ];

    var bezettingMap = {};
    for (var s = 0; s < targetSheets.length; s++) {
      var shMeta = targetSheets[s];
      var sheetObj = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, shMeta.name) : ss.getSheetByName(shMeta.name);
      if (!sheetObj || sheetObj.getLastRow() <= 1) continue;

      var rawData = sheetObj.getDataRange().getDisplayValues();
      var headers = rawData[0];

      var idxNip = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["NIP BARU", "NIP"]) : -1;
      var idxNama = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["NAMA", "NAMA LENGKAP"]) : -1;
      var idxGelarD = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["GELAR DEPAN"]) : -1;
      var idxGelarB = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["GELAR BELAKANG"]) : -1;
      var idxGol = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["GOL AKHIR NAMA", "GOLONGAN"]) : -1;
      var idxPend = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["TINGKAT PENDIDIKAN NAMA", "PENDIDIKAN"]) : -1;
      var idxSatker = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["SATUAN KERJA KERJA NAMA", "SATKER"]) : -1;
      var idxSatkerInduk = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK"]) : -1;
      var idxUnor = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["UNOR NAMA", "UNOR"]) : -1;
      var idxJenisJab = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JENIS JABATAN NAMA", "JENIS JABATAN"]) : -1;
      var idxJabNama = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JABATAN NAMA", "JABATAN"]) : -1;
      var idxEselon = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["ESELON NAMA", "ESELON"]) : -1;

      for (var r = 1; r < rawData.length; r++) {
        var row = rawData[r];
        if (row.every(function(cell) { return String(cell).trim() === ''; })) continue;

        var rawCellJenisJab = idxJenisJab !== -1 ? String(row[idxJenisJab] || '').trim() : '';
        var rawCellJabNama = idxJabNama !== -1 ? String(row[idxJabNama] || '').trim() : '';
        var rawCellEselon = idxEselon !== -1 ? String(row[idxEselon] || '').trim() : '';

        var jabMeta = typeof categorizePegawaiJabatanLocal === 'function'
          ? categorizePegawaiJabatanLocal(rawCellJenisJab, rawCellJabNama, rawCellEselon)
          : { type: 'PELAKSANA' };

        if (jabMeta.type !== 'PELAKSANA') continue;

        var sInduk = idxSatkerInduk !== -1 ? String(row[idxSatkerInduk] || '').trim() : '';
        var sKerja = idxSatker !== -1 ? String(row[idxSatker] || '').trim() : '';
        var sUnor = idxUnor !== -1 ? String(row[idxUnor] || '').trim() : '';

        var targetOpd = typeof mapUnorToTarget === 'function' ? mapUnorToTarget(sInduk, sKerja, sUnor) : (sKerja || sInduk || sUnor || 'Perangkat Daerah');
        targetOpd = typeof normalizeOpdName === 'function' ? normalizeOpdName(targetOpd) : targetOpd;

        var nip = idxNip !== -1 ? String(row[idxNip] || '-').trim() : '-';
        var namaMurni = idxNama !== -1 ? String(row[idxNama] || 'Pegawai').trim() : 'Pegawai';
        var gd = idxGelarD !== -1 && row[idxGelarD] ? String(row[idxGelarD]).trim() + ' ' : '';
        var gb = idxGelarB !== -1 && row[idxGelarB] ? ', ' + String(row[idxGelarB]).trim() : '';
        var fullNama = gd + namaMurni + gb;
        var gol = idxGol !== -1 ? String(row[idxGol] || '-').trim() : '-';
        var pend = idxPend !== -1 ? String(row[idxPend] || '-').trim() : '-';

        var key = cleanMatchKey(targetOpd) + '___' + cleanMatchKey(rawCellJabNama);

        if (!bezettingMap[key]) {
          bezettingMap[key] = { opd: targetOpd, jabatan: rawCellJabNama, pns: 0, pppk: 0, p3k_pw: 0, pegawaiList: [] };
        }

        if (shMeta.statusTag === 'PNS') bezettingMap[key].pns++;
        else if (shMeta.statusTag === 'PPPK') bezettingMap[key].pppk++;
        else bezettingMap[key].p3k_pw++;

        bezettingMap[key].pegawaiList.push({ nip: nip, nama: fullNama, golongan: gol, pendidikan: pend, statusAsn: shMeta.statusTag });
      }
    }

    var combinedMap = {};
    for (var i = 0; i < masterJpList.length; i++) {
      var itemM = masterJpList[i];
      var mKey = cleanMatchKey(itemM.opd) + '___' + cleanMatchKey(itemM.jabatan);
      var bez = bezettingMap[mKey];

      combinedMap[mKey] = {
        perangkatDaerah: itemM.opd, jabatanNama: itemM.jabatan,
        pns: bez ? bez.pns : 0, pppk: bez ? bez.pppk : 0, p3k_pw: bez ? bez.p3k_pw : 0,
        totalBezetting: bez ? (bez.pns + bez.pppk + bez.p3k_pw) : 0,
        kebutuhan: itemM.kebutuhan, sumber: itemM.sumber, keterangan: itemM.keterangan,
        pegawaiList: bez ? bez.pegawaiList : []
      };
    }

    for (var bKey in bezettingMap) {
      if (!combinedMap[bKey]) {
        var itemB = bezettingMap[bKey];
        var matchInMaster = masterJpMap[bKey] || null;
        combinedMap[bKey] = {
          perangkatDaerah: itemB.opd, jabatanNama: itemB.jabatan,
          pns: itemB.pns, pppk: itemB.pppk, p3k_pw: itemB.p3k_pw,
          totalBezetting: itemB.pns + itemB.pppk + itemB.p3k_pw,
          kebutuhan: matchInMaster ? matchInMaster.kebutuhan : 0,
          sumber: matchInMaster ? matchInMaster.sumber : '', keterangan: matchInMaster ? matchInMaster.keterangan : '',
          pegawaiList: itemB.pegawaiList
        };
      }
    }

    var resultRows = [];
    var totalPelaksana = 0, totalKebutuhan = 0, totalDefisitPos = 0, totalSurplusPos = 0, totalGapPegawai = 0;

    for (var key in combinedMap) {
      var rObj = combinedMap[key];
      var gap = 0;
      var statusKebutuhan = 'BELUM_DIATUR';

      if (rObj.kebutuhan > 0) {
        gap = rObj.totalBezetting - rObj.kebutuhan;
        if (gap < 0) {
          statusKebutuhan = 'DEFISIT'; totalDefisitPos++; totalGapPegawai += Math.abs(gap);
        } else if (gap > 0) {
          statusKebutuhan = 'SURPLUS'; totalSurplusPos++;
        } else {
          statusKebutuhan = 'IDEAL';
        }
      }

      rObj.gap = gap;
      rObj.statusKebutuhan = statusKebutuhan;
      rObj.statusLabel = statusKebutuhan === 'DEFISIT' ? ('Kurang ' + Math.abs(gap)) : (statusKebutuhan === 'SURPLUS' ? ('Lebih +' + gap) : (statusKebutuhan === 'IDEAL' ? 'Ideal' : 'Belum Diatur'));

      if (filterOpd && filterOpd !== 'ALL') {
        var cleanFilterOpd = cleanMatchKey(filterOpd);
        var cleanRowOpd = cleanMatchKey(rObj.perangkatDaerah);
        if (cleanRowOpd !== cleanFilterOpd && cleanRowOpd.indexOf(cleanFilterOpd) === -1 && cleanFilterOpd.indexOf(cleanRowOpd) === -1) continue;
      }

      if (filterStatus && filterStatus !== 'ALL') {
        if (rObj.statusKebutuhan !== filterStatus) continue;
      }

      if (filterSearch) {
        var match = rObj.jabatanNama.toLowerCase().indexOf(filterSearch) !== -1 || rObj.perangkatDaerah.toLowerCase().indexOf(filterSearch) !== -1;
        if (!match && rObj.pegawaiList.length > 0) {
          match = rObj.pegawaiList.some(function(p) { return p.nama.toLowerCase().indexOf(filterSearch) !== -1 || p.nip.toLowerCase().indexOf(filterSearch) !== -1; });
        }
        if (!match) continue;
      }

      totalPelaksana += rObj.totalBezetting;
      totalKebutuhan += rObj.kebutuhan;
      resultRows.push(rObj);
    }

    resultRows.sort(function(a, b) {
      var cmp = a.perangkatDaerah.localeCompare(b.perangkatDaerah);
      if (cmp !== 0) return cmp;
      return a.jabatanNama.localeCompare(b.jabatanNama);
    });

    return {
      success: true, data: resultRows,
      stats: { totalPelaksana: totalPelaksana, totalKebutuhan: totalKebutuhan, totalDefisitPos: totalDefisitPos, totalSurplusPos: totalSurplusPos, totalGapPegawai: totalGapPegawai }
    };
  } catch (err) {
    Logger.log("Error getBezettingPelaksanaData: " + err.toString());
    return { success: false, data: [], stats: {}, error: err.toString() };
  }
}

function generateExportBezettingPelaksanaExcel(filters) {
  try {
    var res = getBezettingPelaksanaData(filters);
    if (!res || !res.success) throw new Error("Gagal mengambil data bezetting Pelaksana.");

    var rows = res.data || [];
    var stats = res.stats || {};
    var tanggalCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var dateFileStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');

    var rowsXml = '';
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var gapStr = r.gap > 0 ? ('+' + r.gap) : (r.kebutuhan > 0 ? r.gap : '-');
      var statusColor = r.statusKebutuhan === 'DEFISIT' ? '#ef4444' : (r.statusKebutuhan === 'SURPLUS' ? '#f59e0b' : (r.statusKebutuhan === 'IDEAL' ? '#10b981' : '#64748b'));

      rowsXml += '<tr>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;">' + (i + 1) + '</td>' +
        '<td style="border:0.5pt solid #cccccc;font-weight:bold;">' + r.perangkatDaerah + '</td>' +
        '<td style="border:0.5pt solid #cccccc;">' + r.jabatanNama + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + r.pns + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + r.pppk + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + r.p3k_pw + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background:#e0f2fe;" mso-number-format="#,##0">' + r.totalBezetting + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background:#fef3c7;" mso-number-format="#,##0">' + (r.kebutuhan > 0 ? r.kebutuhan : '-') + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;">' + gapStr + '</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;color:' + statusColor + ';font-weight:bold;">' + r.statusLabel + '</td>' +
      '</tr>';
    }

    var excelHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head>' +
        '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">' +
        '<style>body { font-family: Calibri, Arial, sans-serif; font-size: 10pt; } .title { font-size: 13pt; font-weight: bold; } .th-col { background-color: #0f172a; color: #ffffff; font-weight: bold; text-align: center; border: 0.5pt solid #000; } .th-total { background-color: #0f172a; color: #00ffaa; font-weight: bold; text-align: right; border: 0.5pt solid #000; }</style>' +
      '</head>' +
      '<body>' +
        '<table>' +
          '<tr><td colspan="10" class="title">PEMERINTAH KABUPATEN TOBA</td></tr>' +
          '<tr><td colspan="10" style="font-size:11pt;font-weight:bold;">BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA</td></tr>' +
          '<tr><td colspan="10" style="color:#64748b;font-style:italic;">Matriks Bezetting & Gap Analysis Jabatan Pelaksana (JP) | Tanggal: ' + tanggalCetak + ' WIB</td></tr>' +
          '<tr></tr>' +
          '<tr><th class="th-col">NO</th><th class="th-col">PERANGKAT DAERAH</th><th class="th-col">NOMENKLATUR JABATAN PELAKSANA</th><th class="th-col">PNS</th><th class="th-col">PPPK</th><th class="th-col">PW</th><th class="th-col" style="background:#0284c7;">BEZETTING</th><th class="th-col" style="background:#d97706;">ABK</th><th class="th-col">GAP</th><th class="th-col">STATUS FORMASI</th></tr>' +
          rowsXml +
          '<tr><td colspan="3" style="background-color:#0f172a;color:#ffffff;font-weight:bold;text-align:center;">TOTAL KESELURUHAN</td><td class="th-total" colspan="4" mso-number-format="#,##0">' + stats.totalPelaksana + '</td><td class="th-total" mso-number-format="#,##0">' + stats.totalKebutuhan + '</td><td class="th-total" colspan="2" style="color:#f43f5e;">Defisit: ' + stats.totalGapPegawai + ' Staf</td></tr>' +
        '</table>' +
      '</body></html>';

    var blob = Utilities.newBlob(excelHtml, 'application/vnd.ms-excel', 'Bezetting_Gap_Jabatan_Pelaksana_Toba_' + dateFileStr + '.xls');
    return { success: true, filename: blob.getName(), base64: Utilities.base64Encode(blob.getBytes()) };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/**
 * Helper Rank Golongan untuk Deteksi Kualifikasi Pemangku
 */
function getGolonganRank_(gol) {
  if (!gol) return 0;
  var g = String(gol).toUpperCase().trim();
  var ranks = {
    'I/A': 1, 'I/B': 2, 'I/C': 3, 'I/D': 4,
    'II/A': 5, 'II/B': 6, 'II/C': 7, 'II/D': 8,
    'III/A': 9, 'III/B': 10, 'III/C': 11, 'III/D': 12,
    'IV/A': 13, 'IV/B': 14, 'IV/C': 15, 'IV/D': 16, 'IV/E': 17
  };
  var clean = g.replace(/[\.\s]/g, '/').replace(/([I|V|X]+)([A-E])/g, '$1/$2');
  return ranks[clean] || 0;
}

/**
 * =========================================================================
 * 3. GAP ANALYSIS FORMASI JABATAN STRUKTURAL (ESELON II - IV)
 * =========================================================================
 */
function getStrukturData(filters) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Spreadsheet aktif tidak ditemukan.");

    filters = filters || {};
    var filterSearch = String(filters.search || '').toLowerCase().trim();
    var filterOpd = String(filters.opd || 'ALL').trim();
    var filterEselon = String(filters.eselon || 'ALL').toUpperCase().trim();
    var filterStatus = String(filters.status || 'ALL').toUpperCase().trim();

    var sheetStruktur = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, 'Data_Master_Struktur_Jabatan') : ss.getSheetByName('Data_Master_Struktur_Jabatan');
    if (!sheetStruktur || sheetStruktur.getLastRow() <= 1) {
      return {
        success: true,
        data: [],
        stats: {
          totalFormasi: 0, totalTerisi: 0, totalLowong: 0, totalLowongPlt: 0, totalLowongMurni: 0,
          persenTerisi: 0,
          eselonStat: {
            II: { formasi: 0, terisi: 0, lowong: 0, lowongPlt: 0, lowongMurni: 0 },
            III: { formasi: 0, terisi: 0, lowong: 0, lowongPlt: 0, lowongMurni: 0 },
            IV: { formasi: 0, terisi: 0, lowong: 0, lowongPlt: 0, lowongMurni: 0 }
          }
        }
      };
    }

    var dataStruktur = sheetStruktur.getDataRange().getDisplayValues();
    var hStr = dataStruktur[0];

    var idxStrId = typeof findColumnIndex === 'function' ? findColumnIndex(hStr, ["ID_JABATAN", "ID"]) : 0;
    var idxStrOpd = typeof findColumnIndex === 'function' ? findColumnIndex(hStr, ["PERANGKAT_DAERAH", "OPD"]) : 1;
    var idxStrUnor = typeof findColumnIndex === 'function' ? findColumnIndex(hStr, ["UNOR_NAMA", "UNOR"]) : 2;
    var idxStrJab = typeof findColumnIndex === 'function' ? findColumnIndex(hStr, ["JABATAN_NAMA", "JABATAN"]) : 3;
    var idxStrEsl = typeof findColumnIndex === 'function' ? findColumnIndex(hStr, ["ESELON", "ESELON_NAMA"]) : 4;
    var idxStrFormasi = typeof findColumnIndex === 'function' ? findColumnIndex(hStr, ["JUMLAH_FORMASI", "JUMLAH", "FORMASI"]) : 5;
    var idxStrGolMin = typeof findColumnIndex === 'function' ? findColumnIndex(hStr, ["GOL_MINIMAL", "GOLONGAN"]) : 6;

    var sheetPns = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, 'Data_Pegawai_PNS') : ss.getSheetByName('Data_Pegawai_PNS');
    var pnsMap = {};
    var pltMap = {};
    var pnsByOpd = {};
    var pltByOpd = {};
    var pnsByJabatan = {};
    var pltByJabatan = {};

    if (sheetPns && sheetPns.getLastRow() > 1) {
      var dataPns = sheetPns.getDataRange().getDisplayValues();
      var hPns = dataPns[0];

      var idxPnsNip = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["NIP BARU", "NIP"]) : -1;
      var idxPnsNama = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["NAMA", "NAMA LENGKAP"]) : -1;
      var idxPnsGd = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["GELAR DEPAN"]) : -1;
      var idxPnsGb = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["GELAR BELAKANG"]) : -1;
      var idxPnsJab = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["JABATAN NAMA", "JABATAN"]) : -1;
      var idxPnsGol = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["GOL AKHIR NAMA", "GOLONGAN"]) : -1;
      var idxPnsSatker = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["SATUAN KERJA KERJA NAMA", "SATKER"]) : -1;
      var idxPnsInduk = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK"]) : -1;
      var idxPnsUnor = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["UNOR NAMA", "UNOR"]) : -1;

      for (var p = 1; p < dataPns.length; p++) {
        var rowP = dataPns[p];
        if (rowP.every(function(cell) { return String(cell).trim() === ''; })) continue;

        var sInduk = idxPnsInduk !== -1 ? String(rowP[idxPnsInduk] || '').trim() : '';
        var sKerja = idxPnsSatker !== -1 ? String(rowP[idxPnsSatker] || '').trim() : '';
        var sUnor = idxPnsUnor !== -1 ? String(rowP[idxPnsUnor] || '').trim() : '';
        var jabP = idxPnsJab !== -1 ? String(rowP[idxPnsJab] || '').trim() : '';

        var targetOpd = typeof mapUnorToTarget === 'function' ? mapUnorToTarget(sInduk, sKerja, sUnor, jabP) : (sKerja || sInduk || sUnor);
        targetOpd = typeof normalizeOpdName === 'function' ? normalizeOpdName(targetOpd) : targetOpd;

        var nip = idxPnsNip !== -1 ? String(rowP[idxPnsNip] || '-').trim() : '-';
        var namaMurni = idxPnsNama !== -1 ? String(rowP[idxPnsNama] || '').trim() : '';
        var gd = idxPnsGd !== -1 && rowP[idxPnsGd] ? String(rowP[idxPnsGd]).trim() + ' ' : '';
        var gb = idxPnsGb !== -1 && rowP[idxPnsGb] ? ', ' + String(rowP[idxPnsGb]).trim() : '';
        var gol = idxPnsGol !== -1 ? String(rowP[idxPnsGol] || '-').trim() : '-';
        var fullNama = gd + namaMurni + gb;

        var isPltOfficer = jabP.toUpperCase().indexOf('PLT') !== -1 || jabP.toUpperCase().indexOf('PLH') !== -1 || jabP.toUpperCase().indexOf('PENJABAT') !== -1;
        var cleanJabP = jabP.replace(/^(PLT\.|PLH\.|PJ\.|PLT\s|PLH\s|PJ\s)\s*/i, '').trim();

        var pegawaiObj = {
          nip: nip,
          nama: fullNama,
          golongan: gol,
          opd: targetOpd,
          unor: sUnor,
          satker: sKerja,
          jabatan: cleanJabP
        };

        var keyJab = cleanMatchKey(cleanJabP);
        var jabPStripped = typeof stripOpdFromJabatan === 'function' ? stripOpdFromJabatan(cleanJabP, targetOpd) : cleanJabP;
        var keyJabStripped = cleanMatchKey(jabPStripped);
        var keyJabWithOpd = cleanMatchKey(cleanJabP + ' ' + targetOpd);

        var targetMap = isPltOfficer ? pltMap : pnsMap;
        var targetJabMap = isPltOfficer ? pltByJabatan : pnsByJabatan;
        var targetOpdMap = isPltOfficer ? pltByOpd : pnsByOpd;

        var kTargetOpd = cleanMatchKey(targetOpd);
        if (!targetOpdMap[kTargetOpd]) targetOpdMap[kTargetOpd] = [];
        targetOpdMap[kTargetOpd].push(pegawaiObj);

        // 1. Index OPD + Jabatan
        var keyOpdJab = kTargetOpd + '___' + keyJab;
        if (!targetMap[keyOpdJab]) targetMap[keyOpdJab] = pegawaiObj;

        if (keyJabStripped && keyJabStripped !== keyJab) {
          var keyOpdJabStr = kTargetOpd + '___' + keyJabStripped;
          if (!targetMap[keyOpdJabStr]) targetMap[keyOpdJabStr] = pegawaiObj;
        }

        if (keyJabWithOpd && keyJabWithOpd !== keyJab) {
          var keyOpdJabWith = kTargetOpd + '___' + keyJabWithOpd;
          if (!targetMap[keyOpdJabWith]) targetMap[keyOpdJabWith] = pegawaiObj;
        }

        // Khusus Sekretaris Dinas / Badan / Kecamatan:
        var isSekretarisOffice = (keyJab === 'SEKRETARIS' || keyJabStripped === 'SEKRETARIS' || 
          (keyJab.indexOf('SEKRETARIS') === 0 && keyJab.indexOf('SEKRETARISDAERAH') === -1 && keyJab.indexOf('SEKRETARISDEWAN') === -1 && keyJab.indexOf('SEKRETARISDESA') === -1));
        if (isSekretarisOffice) {
          targetMap[kTargetOpd + '___SEKRETARIS'] = pegawaiObj;
          targetMap[kTargetOpd + '___SEKRETARIS' + kTargetOpd] = pegawaiObj;
          targetMap[kTargetOpd + '___' + cleanMatchKey('SEKRETARIS ' + targetOpd)] = pegawaiObj;
        }

        // 2. Index Satker / Induk + Jabatan jika ada
        if (sKerja) {
          var kKerja = cleanMatchKey(sKerja);
          var keyKerjaJab = kKerja + '___' + keyJab;
          if (!targetMap[keyKerjaJab]) targetMap[keyKerjaJab] = pegawaiObj;
          if (keyJabStripped && keyJabStripped !== keyJab) {
            targetMap[kKerja + '___' + keyJabStripped] = pegawaiObj;
          }
        }
        if (sInduk) {
          var kInduk = cleanMatchKey(sInduk);
          var keyIndukJab = kInduk + '___' + keyJab;
          if (!targetMap[keyIndukJab]) targetMap[keyIndukJab] = pegawaiObj;
          if (keyJabStripped && keyJabStripped !== keyJab) {
            targetMap[kInduk + '___' + keyJabStripped] = pegawaiObj;
          }
        }

        // 3. Index UNOR + Jabatan jika ada
        if (sUnor) {
          var kUnorP = cleanMatchKey(sUnor);
          var keyUnorJab = kUnorP + '___' + keyJab;
          if (!targetMap[keyUnorJab]) targetMap[keyUnorJab] = pegawaiObj;
          if (keyJabStripped && keyJabStripped !== keyJab) {
            targetMap[kUnorP + '___' + keyJabStripped] = pegawaiObj;
          }
        }

        // Khusus Jabatan UPTD / Sub-unit (misal: Kepala Subbagian Tata Usaha UPTD Balai Benih Ikan Lumban Pea)
        var jabShortTu = cleanJabP.replace(/\s+(UPTD?|BALAI|PUSKESWAN)\s+.*$/i, '').trim();
        if (jabShortTu && jabShortTu !== cleanJabP) {
          var keyJabShortTu = cleanMatchKey(jabShortTu);
          if (sUnor) {
            targetMap[cleanMatchKey(sUnor) + '___' + keyJabShortTu] = pegawaiObj;
          }
          if (sKerja) {
            targetMap[cleanMatchKey(sKerja) + '___' + keyJabShortTu] = pegawaiObj;
          }
        }

        // 4. Index Jabatan Tunggal
        if (!targetJabMap[keyJab]) {
          targetJabMap[keyJab] = [];
        }
        targetJabMap[keyJab].push(pegawaiObj);

        if (keyJabStripped && keyJabStripped !== keyJab) {
          if (!targetJabMap[keyJabStripped]) {
            targetJabMap[keyJabStripped] = [];
          }
          targetJabMap[keyJabStripped].push(pegawaiObj);
        }
      }
    }

    var resultList = [];
    var totFormasi = 0, totTerisi = 0, totLowong = 0, totLowongPlt = 0, totLowongMurni = 0;
    var eselonStat = {
      II: { formasi: 0, terisi: 0, lowong: 0, lowongPlt: 0, lowongMurni: 0 },
      III: { formasi: 0, terisi: 0, lowong: 0, lowongPlt: 0, lowongMurni: 0 },
      IV: { formasi: 0, terisi: 0, lowong: 0, lowongPlt: 0, lowongMurni: 0 }
    };

    var usedPnsNips = {};
    var usedPltNips = {};

    for (var i = 1; i < dataStruktur.length; i++) {
      var row = dataStruktur[i];
      if (row.every(function(cell) { return String(cell).trim() === ''; })) continue;

      var rawOpd = idxStrOpd !== -1 ? String(row[idxStrOpd] || '').trim() : '';
      var rawUnor = idxStrUnor !== -1 ? String(row[idxStrUnor] || '').trim() : '';
      var rawJab = idxStrJab !== -1 ? String(row[idxStrJab] || '').trim() : '';
      var rawEsl = idxStrEsl !== -1 ? String(row[idxStrEsl] || '').trim() : '-';
      var rawFormasi = idxStrFormasi !== -1 ? parseInt(row[idxStrFormasi], 10) || 1 : 1;
      var rawGolMin = idxStrGolMin !== -1 ? String(row[idxStrGolMin] || '-').trim() : '-';

      var eslUpper = rawEsl.toUpperCase();
      var eslKey = null;
      if ((eslUpper.indexOf('II') !== -1 && eslUpper.indexOf('III') === -1) || eslUpper.indexOf('JPT') !== -1) {
        eslKey = 'II';
      } else if (eslUpper.indexOf('III') !== -1 || eslUpper.indexOf('ADMINISTRATOR') !== -1) {
        eslKey = 'III';
      } else if (eslUpper.indexOf('IV') !== -1 || eslUpper.indexOf('PENGAWAS') !== -1) {
        eslKey = 'IV';
      }

      var normOpd = typeof normalizeOpdName === 'function' ? normalizeOpdName(rawOpd) : rawOpd;
      var kOpd = cleanMatchKey(normOpd);
      var kUnor = cleanMatchKey(rawUnor);
      var kJab = cleanMatchKey(rawJab);

      var jabStripped = typeof stripOpdFromJabatan === 'function' ? stripOpdFromJabatan(rawJab, normOpd) : rawJab;
      var kJabStripped = cleanMatchKey(jabStripped);
      var kJabWithOpd = cleanMatchKey(rawJab + ' ' + normOpd);

      var keyStr = kOpd + '___' + kJab;
      var keyStrStripped = kOpd + '___' + kJabStripped;
      var keyStrWithOpd = kOpd + '___' + kJabWithOpd;
      var keyUnorStr = kUnor ? (kUnor + '___' + kJab) : '';
      var keyUnorStripped = kUnor ? (kUnor + '___' + kJabStripped) : '';

      var isSekretarisJab = (kJab === 'SEKRETARIS' || kJabStripped === 'SEKRETARIS' ||
        (kJab.indexOf('SEKRETARIS') === 0 && kJab.indexOf('SEKRETARISDAERAH') === -1 && kJab.indexOf('SEKRETARISDEWAN') === -1 && kJab.indexOf('SEKRETARISDESA') === -1));

      var isInspekturJab = (kJab === 'INSPEKTUR' || kJab === 'INSPEKTURINSPEKTORAT' || kJab === 'INSPEKTURDAERAH');

      // -------------------------------------------------------------
      // Cari Pejabat Definitif (Pemangku):
      // -------------------------------------------------------------
      var pemangku = null;

      // Level 1: Pencocokan Kunci Langsung Berdasarkan OPD
      var candP = pnsMap[keyStr] || pnsMap[keyStrStripped] || pnsMap[keyStrWithOpd];
      if (!candP && (keyUnorStr || keyUnorStripped)) {
        var candUnorCandidate = (keyUnorStr ? pnsMap[keyUnorStr] : null) || (keyUnorStripped ? pnsMap[keyUnorStripped] : null);
        if (candUnorCandidate) {
          var candOpdCheck = cleanMatchKey(candUnorCandidate.opd);
          var isPerkimUnor = (kOpd.indexOf('PERUMAHAN') !== -1 || kOpd.indexOf('PERKIM') !== -1) &&
                             (candOpdCheck.indexOf('PERUMAHAN') !== -1 || candOpdCheck.indexOf('PERKIM') !== -1 || candOpdCheck.indexOf('PEKERJAANUMUM') !== -1);
          if (candOpdCheck === kOpd || candOpdCheck.indexOf(kOpd) !== -1 || kOpd.indexOf(candOpdCheck) !== -1 || isPerkimUnor) {
            candP = candUnorCandidate;
          }
        }
      }

      if (candP && (!candP.nip || candP.nip === '-' || !usedPnsNips[candP.nip])) {
        var candJabUpper = String(candP.jabatan || '').toUpperCase();
        if (isInspekturJab && (candJabUpper.indexOf('PEMBANTU') !== -1 || candJabUpper.indexOf('IRBAN') !== -1)) {
          // Proteksi: Jangan pernah cocokkan Inspektur Pembantu dengan Inspektur
        } else {
          pemangku = candP;
        }
      }

      // Level 2: Khusus Jabatan Sekretaris OPD
      if (!pemangku && isSekretarisJab) {
        var candSek = pnsMap[kOpd + '___SEKRETARIS'] || pnsMap[kOpd + '___SEKRETARIS' + kOpd] || pnsMap[kOpd + '___' + cleanMatchKey('SEKRETARIS ' + normOpd)];
        if (candSek && (!candSek.nip || candSek.nip === '-' || !usedPnsNips[candSek.nip])) {
          pemangku = candSek;
        } else if (pnsByOpd[kOpd]) {
          var opdStaff = pnsByOpd[kOpd];
          for (var s = 0; s < opdStaff.length; s++) {
            var st = opdStaff[s];
            if (st.nip && st.nip !== '-' && usedPnsNips[st.nip]) continue;
            var stJab = cleanMatchKey(st.jabatan);
            if (stJab === 'SEKRETARIS' || (stJab.indexOf('SEKRETARIS') === 0 && stJab.indexOf('SEKRETARISDAERAH') === -1 && stJab.indexOf('SEKRETARISDEWAN') === -1 && stJab.indexOf('SEKRETARISDESA') === -1)) {
              pemangku = st;
              break;
            }
          }
        }
      }

      // Level 3: Pencocokan Presisi dalam OPD yang Sama (Khusus Eselon IV/III: Kasubbag, Kasi, Kasubbid)
      // JPT Pratama (Eselon II) dan Inspektur tidak boleh menggunakan pencocokan substring Level 3
      if (!pemangku && pnsByOpd[kOpd] && eslKey !== 'II' && !isInspekturJab) {
        var opdList = pnsByOpd[kOpd];
        var coreMaster = (kJabStripped || kJab).replace(/^(KEPALASUBBAGIAN|KEPALASUBBIDANG|KEPALASEKSI|KEPALABIDANG|KEPALABAGIAN)/, '');
        for (var o = 0; o < opdList.length; o++) {
          var emp = opdList[o];
          if (emp.nip && emp.nip !== '-' && usedPnsNips[emp.nip]) continue;
          var empCleanJab = cleanMatchKey(emp.jabatan);
          var empStripJab = cleanMatchKey(typeof stripOpdFromJabatan === 'function' ? stripOpdFromJabatan(emp.jabatan, normOpd) : emp.jabatan);
          var empJabRaw = String(emp.jabatan || '').toUpperCase();
          if (empJabRaw.indexOf('PEMBANTU') !== -1 || empJabRaw.indexOf('IRBAN') !== -1) continue;

          if (empCleanJab === kJab || empCleanJab === kJabStripped || empStripJab === kJab || empStripJab === kJabStripped) {
            pemangku = emp;
            break;
          }

          if (coreMaster && coreMaster.length >= 4) {
            var coreEmp = (empStripJab || empCleanJab).replace(/^(KEPALASUBBAGIAN|KEPALASUBBIDANG|KEPALASEKSI|KEPALABIDANG|KEPALABAGIAN)/, '');
            if (coreEmp && (coreMaster === coreEmp || coreMaster.indexOf(coreEmp) !== -1 || coreEmp.indexOf(coreMaster) !== -1)) {
              var requiresUnitMatch = coreEmp.indexOf('LUMBANPEA') !== -1 || coreMaster.indexOf('LUMBANPEA') !== -1 ||
                                      coreEmp.indexOf('PUSKESWAN') !== -1 || coreMaster.indexOf('PUSKESWAN') !== -1;
              if (requiresUnitMatch) {
                var empAll = (emp.jabatan + ' ' + emp.unor + ' ' + emp.satker).toUpperCase();
                var masterAll = (rawJab + ' ' + rawUnor).toUpperCase();
                if ((empAll.indexOf('LUMBAN PEA') !== -1 && masterAll.indexOf('LUMBAN PEA') !== -1) ||
                    (empAll.indexOf('PUSKESWAN') !== -1 && masterAll.indexOf('PUSKESWAN') !== -1)) {
                  pemangku = emp;
                  break;
                }
              } else {
                pemangku = emp;
                break;
              }
            }
          }
        }
      }

      // Level 4: Fallback Berdasarkan Jabatan (Wajib dalam Batasan OPD / UNOR yang Beririsan)
      if (!pemangku) {
        var jabList = pnsByJabatan[kJab] || (kJabStripped ? pnsByJabatan[kJabStripped] : null);
        if (jabList) {
          for (var c = 0; c < jabList.length; c++) {
            var cand = jabList[c];
            if (cand.nip && cand.nip !== '-' && usedPnsNips[cand.nip]) continue;
            var candJabRaw = String(cand.jabatan || '').toUpperCase();
            if (isInspekturJab && (candJabRaw.indexOf('PEMBANTU') !== -1 || candJabRaw.indexOf('IRBAN') !== -1)) continue;

            var candOpd = cleanMatchKey(cand.opd);
            var candUnor = cleanMatchKey(cand.unor);
            var candCleanJab = cleanMatchKey(cand.jabatan);

            // Verifikasi mutlak batasan OPD agar pejabat dari OPD lain tidak terambil
            var isPerkimMatch = (kOpd.indexOf('PERUMAHAN') !== -1 || kOpd.indexOf('PERKIM') !== -1) &&
                                (candOpd.indexOf('PERUMAHAN') !== -1 || candOpd.indexOf('PERKIM') !== -1 || candUnor.indexOf('PERUMAHAN') !== -1 || candUnor.indexOf('PERKIM') !== -1 || candCleanJab.indexOf('PERUMAHAN') !== -1 || candOpd.indexOf('PEKERJAANUMUM') !== -1);

            var isOpdMatch = (candOpd === kOpd || candUnor === kUnor || candOpd.indexOf(kOpd) !== -1 || kOpd.indexOf(candOpd) !== -1 || (kUnor && (candUnor.indexOf(kUnor) !== -1 || kUnor.indexOf(candUnor) !== -1)) || isPerkimMatch);

            if (isOpdMatch) {
              if (eslKey === 'II') {
                if (candCleanJab === kJab || candCleanJab === kJabStripped) {
                  pemangku = cand;
                  break;
                }
              } else {
                pemangku = cand;
                break;
              }
            }
          }
        }
      }

      if (pemangku && pemangku.nip && pemangku.nip !== '-') {
        usedPnsNips[pemangku.nip] = true;
      }

      // -------------------------------------------------------------
      // Cari Pelaksana Tugas (Plt / Plh):
      // -------------------------------------------------------------
      var pltInfo = null;

      // Level 1: Pencocokan Kunci Langsung Berdasarkan OPD
      var candPlt = pltMap[keyStr] || pltMap[keyStrStripped] || pltMap[keyStrWithOpd];
      if (!candPlt && (keyUnorStr || keyUnorStripped)) {
        var candPltUnorCandidate = (keyUnorStr ? pltMap[keyUnorStr] : null) || (keyUnorStripped ? pltMap[keyUnorStripped] : null);
        if (candPltUnorCandidate) {
          var candPltOpdCheck = cleanMatchKey(candPltUnorCandidate.opd);
          var isPerkimPltUnor = (kOpd.indexOf('PERUMAHAN') !== -1 || kOpd.indexOf('PERKIM') !== -1) &&
                                (candPltOpdCheck.indexOf('PERUMAHAN') !== -1 || candPltOpdCheck.indexOf('PERKIM') !== -1 || candPltOpdCheck.indexOf('PEKERJAANUMUM') !== -1);
          if (candPltOpdCheck === kOpd || candPltOpdCheck.indexOf(kOpd) !== -1 || kOpd.indexOf(candPltOpdCheck) !== -1 || isPerkimPltUnor) {
            candPlt = candPltUnorCandidate;
          }
        }
      }

      if (candPlt && (!candPlt.nip || candPlt.nip === '-' || !usedPltNips[candPlt.nip])) {
        var candPltJabUpper = String(candPlt.jabatan || '').toUpperCase();
        if (isInspekturJab && (candPltJabUpper.indexOf('PEMBANTU') !== -1 || candPltJabUpper.indexOf('IRBAN') !== -1)) {
          // Proteksi
        } else {
          pltInfo = candPlt;
        }
      }

      // Level 2: Khusus Jabatan Sekretaris OPD
      if (!pltInfo && isSekretarisJab) {
        var candPltSek = pltMap[kOpd + '___SEKRETARIS'] || pltMap[kOpd + '___SEKRETARIS' + kOpd] || pltMap[kOpd + '___' + cleanMatchKey('SEKRETARIS ' + normOpd)];
        if (candPltSek && (!candPltSek.nip || candPltSek.nip === '-' || !usedPltNips[candPltSek.nip])) {
          pltInfo = candPltSek;
        } else if (pltByOpd[kOpd]) {
          var opdPltStaff = pltByOpd[kOpd];
          for (var sp = 0; sp < opdPltStaff.length; sp++) {
            var stp = opdPltStaff[sp];
            if (stp.nip && stp.nip !== '-' && usedPltNips[stp.nip]) continue;
            var stpJab = cleanMatchKey(stp.jabatan);
            if (stpJab === 'SEKRETARIS' || (stpJab.indexOf('SEKRETARIS') === 0 && stpJab.indexOf('SEKRETARISDAERAH') === -1 && stpJab.indexOf('SEKRETARISDEWAN') === -1 && stpJab.indexOf('SEKRETARISDESA') === -1)) {
              pltInfo = stp;
              break;
            }
          }
        }
      }

      // Level 3: Pencocokan Presisi dalam OPD yang Sama (Khusus Eselon IV/III)
      if (!pltInfo && pltByOpd[kOpd] && eslKey !== 'II' && !isInspekturJab) {
        var opdPltList = pltByOpd[kOpd];
        var coreMasterPlt = (kJabStripped || kJab).replace(/^(KEPALASUBBAGIAN|KEPALASUBBIDANG|KEPALASEKSI|KEPALABIDANG|KEPALABAGIAN)/, '');
        for (var op = 0; op < opdPltList.length; op++) {
          var empPlt = opdPltList[op];
          if (empPlt.nip && empPlt.nip !== '-' && usedPltNips[empPlt.nip]) continue;
          var empCleanJabPlt = cleanMatchKey(empPlt.jabatan);
          var empStripJabPlt = cleanMatchKey(typeof stripOpdFromJabatan === 'function' ? stripOpdFromJabatan(empPlt.jabatan, normOpd) : empPlt.jabatan);
          var empPltJabRaw = String(empPlt.jabatan || '').toUpperCase();
          if (empPltJabRaw.indexOf('PEMBANTU') !== -1 || empPltJabRaw.indexOf('IRBAN') !== -1) continue;
          
          if (empCleanJabPlt === kJab || empCleanJabPlt === kJabStripped || empStripJabPlt === kJab || empStripJabPlt === kJabStripped) {
            pltInfo = empPlt;
            break;
          }

          if (coreMasterPlt && coreMasterPlt.length >= 4) {
            var coreEmpPlt = (empStripJabPlt || empCleanJabPlt).replace(/^(KEPALASUBBAGIAN|KEPALASUBBIDANG|KEPALASEKSI|KEPALABIDANG|KEPALABAGIAN)/, '');
            if (coreEmpPlt && (coreMasterPlt === coreEmpPlt || coreMasterPlt.indexOf(coreEmpPlt) !== -1 || coreEmpPlt.indexOf(coreMasterPlt) !== -1)) {
              var requiresUnitMatchPlt = coreEmpPlt.indexOf('LUMBANPEA') !== -1 || coreMasterPlt.indexOf('LUMBANPEA') !== -1 ||
                                         coreEmpPlt.indexOf('PUSKESWAN') !== -1 || coreMasterPlt.indexOf('PUSKESWAN') !== -1;
              if (requiresUnitMatchPlt) {
                var empAllPlt = (empPlt.jabatan + ' ' + empPlt.unor + ' ' + empPlt.satker).toUpperCase();
                var masterAllPlt = (rawJab + ' ' + rawUnor).toUpperCase();
                if ((empAllPlt.indexOf('LUMBAN PEA') !== -1 && masterAllPlt.indexOf('LUMBAN PEA') !== -1) ||
                    (empAllPlt.indexOf('PUSKESWAN') !== -1 && masterAllPlt.indexOf('PUSKESWAN') !== -1)) {
                  pltInfo = empPlt;
                  break;
                }
              } else {
                pltInfo = empPlt;
                break;
              }
            }
          }
        }
      }

      // Level 4: Fallback Global Berdasarkan Jabatan Tunggal
      if (!pltInfo) {
        var pltJabList = pltByJabatan[kJab] || (kJabStripped ? pltByJabatan[kJabStripped] : null);
        if (pltJabList) {
          for (var cp = 0; cp < pltJabList.length; cp++) {
            var candPltItem = pltJabList[cp];
            if (candPltItem.nip && candPltItem.nip !== '-' && usedPltNips[candPltItem.nip]) continue;
            var candPltJabRaw = String(candPltItem.jabatan || '').toUpperCase();
            if (isInspekturJab && (candPltJabRaw.indexOf('PEMBANTU') !== -1 || candPltJabRaw.indexOf('IRBAN') !== -1)) continue;

            var candPltOpd = cleanMatchKey(candPltItem.opd);
            var candPltUnor = cleanMatchKey(candPltItem.unor);
            var candPltCleanJab = cleanMatchKey(candPltItem.jabatan);

            var isPerkimPltMatch = (kOpd.indexOf('PERUMAHAN') !== -1 || kOpd.indexOf('PERKIM') !== -1) &&
                                   (candPltOpd.indexOf('PERUMAHAN') !== -1 || candPltOpd.indexOf('PERKIM') !== -1 || candPltUnor.indexOf('PERUMAHAN') !== -1 || candPltUnor.indexOf('PERKIM') !== -1 || candPltCleanJab.indexOf('PERUMAHAN') !== -1 || candPltOpd.indexOf('PEKERJAANUMUM') !== -1);

            var isPltOpdMatch = (candPltOpd === kOpd || candPltUnor === kUnor || candPltOpd.indexOf(kOpd) !== -1 || kOpd.indexOf(candPltOpd) !== -1 || (kUnor && (candPltUnor.indexOf(kUnor) !== -1 || kUnor.indexOf(candPltUnor) !== -1)) || isPerkimPltMatch);

            if (isPltOpdMatch) {
              if (eslKey === 'II') {
                if (candPltCleanJab === kJab || candPltCleanJab === kJabStripped) {
                  pltInfo = candPltItem;
                  break;
                }
              } else {
                pltInfo = candPltItem;
                break;
              }
            }
          }
        }
      }

      if (pltInfo && pltInfo.nip && pltInfo.nip !== '-') {
        usedPltNips[pltInfo.nip] = true;
      }

      var status = 'LOWONG_MURNI';
      var statusLabel = 'Lowong Murni';
      var isTerisi = false;
      var isPlt = false;
      var isUnderQualified = false;
      var qualificationNote = '';

      var pejNama = '-';
      var pejNip = '-';
      var pejGol = '-';
      var pltNama = '';
      var pltNip = '';
      var pltGol = '';

      if (pemangku) {
        status = 'TERISI';
        statusLabel = 'Terisi';
        isTerisi = true;
        pejNama = pemangku.nama;
        pejNip = pemangku.nip;
        pejGol = pemangku.golongan;

        var minRank = getGolonganRank_(rawGolMin);
        var currRank = getGolonganRank_(pejGol);
        if (minRank > 0 && currRank > 0 && currRank < minRank) {
          isUnderQualified = true;
          qualificationNote = 'Golongan pemangku (' + pejGol + ') di bawah syarat minimal (' + rawGolMin + ')';
        }
      } else if (pltInfo) {
        status = 'LOWONG_PLT';
        statusLabel = 'Lowong (Diisi Plt)';
        isPlt = true;
        pltNama = pltInfo.nama;
        pltNip = pltInfo.nip;
        pltGol = pltInfo.golongan;
      }

      // Akumulasi KPI Keseluruhan
      totFormasi += rawFormasi;
      if (status === 'TERISI') {
        totTerisi += rawFormasi;
      } else {
        totLowong += rawFormasi;
        if (status === 'LOWONG_PLT') totLowongPlt += rawFormasi;
        else totLowongMurni += rawFormasi;
      }

      if (eslKey && eselonStat[eslKey]) {
        eselonStat[eslKey].formasi += rawFormasi;
        if (status === 'TERISI') {
          eselonStat[eslKey].terisi += rawFormasi;
        } else {
          eselonStat[eslKey].lowong += rawFormasi;
          if (status === 'LOWONG_PLT') eselonStat[eslKey].lowongPlt += rawFormasi;
          else eselonStat[eslKey].lowongMurni += rawFormasi;
        }
      }

      // Filter OPD
      if (filterOpd && filterOpd !== 'ALL' && filterOpd !== '') {
        var cleanFilterOpd = cleanMatchKey(filterOpd);
        var cleanRowOpd = cleanMatchKey(normOpd);
        if (cleanRowOpd !== cleanFilterOpd && cleanRowOpd.indexOf(cleanFilterOpd) === -1 && cleanFilterOpd.indexOf(cleanRowOpd) === -1) {
          continue;
        }
      }

      // Filter Eselon Presisi (Eselon II hanya memunculkan Eselon II)
      if (filterEselon && filterEselon !== 'ALL' && filterEselon !== '') {
        if (filterEselon === 'II') {
          if (eslKey !== 'II') continue;
        } else if (filterEselon === 'III') {
          if (eslKey !== 'III') continue;
        } else if (filterEselon === 'IV') {
          if (eslKey !== 'IV') continue;
        } else {
          if (eslUpper.indexOf(filterEselon) === -1) continue;
        }
      }

      // Filter Status (3-Tier & Under-Qualified)
      if (filterStatus && filterStatus !== 'ALL' && filterStatus !== '') {
        if (filterStatus === 'TERISI' && status !== 'TERISI') continue;
        if (filterStatus === 'LOWONG' && status === 'TERISI') continue;
        if (filterStatus === 'LOWONG_PLT' && status !== 'LOWONG_PLT') continue;
        if (filterStatus === 'LOWONG_MURNI' && status !== 'LOWONG_MURNI') continue;
        if (filterStatus === 'UNDER_QUALIFIED' && !isUnderQualified) continue;
      }

      // Filter Search
      if (filterSearch) {
        var match = rawJab.toLowerCase().indexOf(filterSearch) !== -1 ||
                    normOpd.toLowerCase().indexOf(filterSearch) !== -1 ||
                    rawUnor.toLowerCase().indexOf(filterSearch) !== -1 ||
                    pejNama.toLowerCase().indexOf(filterSearch) !== -1 ||
                    pejNip.indexOf(filterSearch) !== -1 ||
                    pltNama.toLowerCase().indexOf(filterSearch) !== -1 ||
                    pltNip.indexOf(filterSearch) !== -1;
        if (!match) continue;
      }

      resultList.push({
        id: idxStrId !== -1 ? String(row[idxStrId] || '').trim() : ('STR-' + i),
        perangkatDaerah: normOpd,
        unorNama: rawUnor,
        jabatanNama: rawJab,
        eselon: rawEsl,
        jumlahFormasi: rawFormasi,
        golMinimal: rawGolMin,
        isTerisi: isTerisi,
        isPlt: isPlt,
        status: status,
        statusLabel: statusLabel,
        isUnderQualified: isUnderQualified,
        qualificationNote: qualificationNote,
        pejabatNama: pejNama,
        pejabatNip: pejNip,
        pejabatGol: pejGol,
        pltNama: pltNama,
        pltNip: pltNip,
        pltGol: pltGol
      });
    }

    var persenTerisi = totFormasi > 0 ? Math.round((totTerisi / totFormasi) * 100) : 0;

    return {
      success: true,
      data: resultList,
      stats: {
        totalFormasi: totFormasi,
        totalTerisi: totTerisi,
        totalLowong: totLowong,
        totalLowongPlt: totLowongPlt,
        totalLowongMurni: totLowongMurni,
        persenTerisi: persenTerisi,
        eselonStat: eselonStat
      }
    };
  } catch (err) {
    Logger.log("Error getStrukturData: " + err.toString());
    return {
      success: false,
      data: [],
      stats: {
        totalFormasi: 0, totalTerisi: 0, totalLowong: 0, totalLowongPlt: 0, totalLowongMurni: 0,
        persenTerisi: 0,
        eselonStat: {
          II: { formasi: 0, terisi: 0, lowong: 0, lowongPlt: 0, lowongMurni: 0 },
          III: { formasi: 0, terisi: 0, lowong: 0, lowongPlt: 0, lowongMurni: 0 },
          IV: { formasi: 0, terisi: 0, lowong: 0, lowongPlt: 0, lowongMurni: 0 }
        }
      },
      error: err.toString()
    };
  }
}

/**
 * Sinkronisasi Otomatis & Deduplikasi Master SOTK Jabatan Struktural dari Data PNS
 */
function syncMasterStrukturFromUI() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Spreadsheet aktif tidak ditemukan.");

    var sheetPns = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, 'Data_Pegawai_PNS') : ss.getSheetByName('Data_Pegawai_PNS');
    if (!sheetPns || sheetPns.getLastRow() <= 1) {
      throw new Error("Data PNS tidak ditemukan atau masih kosong.");
    }

    var sheetStruktur = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, 'Data_Master_Struktur_Jabatan') : ss.getSheetByName('Data_Master_Struktur_Jabatan');
    if (!sheetStruktur) {
      sheetStruktur = ss.insertSheet('Data_Master_Struktur_Jabatan');
      sheetStruktur.getRange(1, 1, 1, 7).setValues([
        ['ID_JABATAN', 'PERANGKAT_DAERAH', 'UNOR_NAMA', 'JABATAN_NAMA', 'ESELON', 'JUMLAH_FORMASI', 'GOL_MINIMAL']
      ]);
    }

    // 1. Baca master eksisting untuk mempertahankan formasi & gol minimal kustom
    var existingMasterMap = {};
    if (sheetStruktur.getLastRow() > 1) {
      var dMaster = sheetStruktur.getDataRange().getDisplayValues();
      var hM = dMaster[0];
      var idxMOpd = typeof findColumnIndex === 'function' ? findColumnIndex(hM, ["PERANGKAT_DAERAH", "OPD"]) : 1;
      var idxMUnor = typeof findColumnIndex === 'function' ? findColumnIndex(hM, ["UNOR_NAMA", "UNOR"]) : 2;
      var idxMJab = typeof findColumnIndex === 'function' ? findColumnIndex(hM, ["JABATAN_NAMA", "JABATAN"]) : 3;
      var idxMEsl = typeof findColumnIndex === 'function' ? findColumnIndex(hM, ["ESELON"]) : 4;
      var idxMFormasi = typeof findColumnIndex === 'function' ? findColumnIndex(hM, ["JUMLAH_FORMASI", "JUMLAH", "FORMASI"]) : 5;
      var idxMGolMin = typeof findColumnIndex === 'function' ? findColumnIndex(hM, ["GOL_MINIMAL", "GOLONGAN"]) : 6;

      for (var m = 1; m < dMaster.length; m++) {
        var rowM = dMaster[m];
        if (rowM.every(function(c) { return String(c).trim() === ''; })) continue;
        var mOpd = typeof normalizeOpdName === 'function' ? normalizeOpdName(rowM[idxMOpd]) : rowM[idxMOpd];
        var mUnor = String(rowM[idxMUnor] || '').trim();
        var mJab = String(rowM[idxMJab] || '').trim();
        var mEsl = String(rowM[idxMEsl] || '').trim();
        var mFormasi = parseInt(rowM[idxMFormasi], 10) || 1;
        var mGolMin = String(rowM[idxMGolMin] || '').trim();

        var mKey = cleanMatchKey(mOpd) + '___' + cleanMatchKey(mJab);
        if (!existingMasterMap[mKey]) {
          existingMasterMap[mKey] = {
            opd: mOpd, unor: mUnor || mOpd, jabatan: mJab, eselon: mEsl, formasi: mFormasi, golMinimal: mGolMin
          };
        }
      }
    }

    // 2. Pindai seluruh data PNS untuk mendeteksi jabatan struktural riil
    var dataPns = sheetPns.getDataRange().getDisplayValues();
    var hPns = dataPns[0];
    var idxInduk = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK"]) : -1;
    var idxSatker = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["SATUAN KERJA KERJA NAMA", "SATKER"]) : -1;
    var idxUnor = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["UNOR NAMA", "UNOR"]) : -1;
    var idxJab = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["JABATAN NAMA", "JABATAN"]) : -1;
    var idxEsl = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["ESELON NAMA", "ESELON"]) : -1;
    var idxJenis = typeof findColumnIndex === 'function' ? findColumnIndex(hPns, ["JENIS JABATAN NAMA", "JENIS JABATAN"]) : -1;

    for (var p = 1; p < dataPns.length; p++) {
      var rowP = dataPns[p];
      if (rowP.every(function(c) { return String(c).trim() === ''; })) continue;

      var rawEsl = idxEsl !== -1 ? String(rowP[idxEsl] || '').trim() : '';
      var rawJenis = idxJenis !== -1 ? String(rowP[idxJenis] || '').trim().toUpperCase() : '';
      var rawJab = idxJab !== -1 ? String(rowP[idxJab] || '').trim() : '';

      var isStruktural = (rawEsl && rawEsl !== '-' && rawEsl !== 'NON ESELON' && rawEsl !== '0') ||
                         rawJenis.indexOf('STRUKTUR') !== -1 || rawJenis.indexOf('PIMPINAN') !== -1;

      if (!isStruktural || !rawJab) continue;

      var cleanJab = rawJab.replace(/^(PLT\.|PLH\.|PJ\.|PLT\s|PLH\s|PJ\s)\s*/i, '').trim();

      var sInduk = idxInduk !== -1 ? String(rowP[idxInduk] || '').trim() : '';
      var sKerja = idxSatker !== -1 ? String(rowP[idxSatker] || '').trim() : '';
      var sUnor = idxUnor !== -1 ? String(rowP[idxUnor] || '').trim() : '';
      var targetOpd = typeof mapUnorToTarget === 'function' ? mapUnorToTarget(sInduk, sKerja, sUnor, cleanJab) : (sKerja || sInduk);
      targetOpd = typeof normalizeOpdName === 'function' ? normalizeOpdName(targetOpd) : targetOpd;

      var key = cleanMatchKey(targetOpd) + '___' + cleanMatchKey(cleanJab);
      if (!existingMasterMap[key]) {
        var defaultGol = 'III/c';
        var eslUpper = rawEsl.toUpperCase();
        if (eslUpper.indexOf('II.A') !== -1) defaultGol = 'IV/c';
        else if (eslUpper.indexOf('II.B') !== -1 || eslUpper === 'II') defaultGol = 'IV/b';
        else if (eslUpper.indexOf('III.A') !== -1) defaultGol = 'IV/a';
        else if (eslUpper.indexOf('III.B') !== -1 || eslUpper === 'III') defaultGol = 'III/d';
        else if (eslUpper.indexOf('IV.A') !== -1) defaultGol = 'III/c';
        else if (eslUpper.indexOf('IV.B') !== -1 || eslUpper === 'IV') defaultGol = 'III/b';

        existingMasterMap[key] = {
          opd: targetOpd,
          unor: sUnor || targetOpd,
          jabatan: cleanJab,
          eselon: rawEsl || 'IV.a',
          formasi: 1,
          golMinimal: defaultGol
        };
      }
    }

    var sortedRows = Object.keys(existingMasterMap).map(function(k) {
      return existingMasterMap[k];
    });

    sortedRows.sort(function(a, b) {
      var cmpOpd = a.opd.localeCompare(b.opd);
      if (cmpOpd !== 0) return cmpOpd;
      return a.jabatan.localeCompare(b.jabatan);
    });

    var finalValues = [
      ['ID_JABATAN', 'PERANGKAT_DAERAH', 'UNOR_NAMA', 'JABATAN_NAMA', 'ESELON', 'JUMLAH_FORMASI', 'GOL_MINIMAL']
    ];

    for (var s = 0; s < sortedRows.length; s++) {
      var item = sortedRows[s];
      var idStr = 'STR-' + String(s + 1).padStart(3, '0');
      finalValues.push([
        idStr,
        item.opd,
        item.unor || item.opd,
        item.jabatan,
        item.eselon || '-',
        item.formasi || 1,
        item.golMinimal || 'III/c'
      ]);
    }

    sheetStruktur.clear();
    sheetStruktur.getRange(1, 1, finalValues.length, 7).setValues(finalValues);
    sheetStruktur.getRange(1, 1, 1, 7)
      .setFontWeight('bold')
      .setBackground('#0f172a')
      .setFontColor('#ffffff');
    sheetStruktur.setFrozenRows(1);
    SpreadsheetApp.flush();

    return {
      success: true,
      count: sortedRows.length,
      message: "Sinkronisasi Master SOTK berhasil! " + sortedRows.length + " formasi jabatan struktural diselaraskan."
    };
  } catch (err) {
    Logger.log("Error syncMasterStrukturFromUI: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

function generateExportStrukturExcel(filters) {
  try {
    var res = getStrukturData(filters);
    if (!res || !res.success) throw new Error("Gagal mengambil data struktur: " + ((res && res.error) || ''));

    var rows = res.data || [];
    var tanggalCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var dateFileStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');

    var rowsXml = '';
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var statusColor = r.status === 'TERISI' ? '#10b981' : (r.status === 'LOWONG_PLT' ? '#f59e0b' : '#ef4444');
      var pemangkuName = r.status === 'TERISI' ? r.pejabatNama : (r.status === 'LOWONG_PLT' ? ('Plt: ' + r.pltNama) : '(Lowong Murni)');
      var nipStr = r.status === 'TERISI' ? r.pejabatNip : (r.status === 'LOWONG_PLT' ? r.pltNip : '-');
      var golStr = r.status === 'TERISI' ? r.pejabatGol : (r.status === 'LOWONG_PLT' ? r.pltGol : '-');

      rowsXml += '<tr>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;">' + (i + 1) + '</td>' +
        '<td style="border:0.5pt solid #cccccc;font-weight:bold;">' + r.perangkatDaerah + '</td>' +
        '<td style="border:0.5pt solid #cccccc;">' + r.unorNama + '</td>' +
        '<td style="border:0.5pt solid #cccccc;font-weight:bold;">' + r.jabatanNama + '</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;">' + r.eselon + '</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;font-weight:bold;color:' + statusColor + ';">' + r.statusLabel + '</td>' +
        '<td style="border:0.5pt solid #cccccc;">' + pemangkuName + (r.isUnderQualified ? ' (Under-Qualified)' : '') + '</td>' +
        '<td style="border:0.5pt solid #cccccc;text-align:center;" mso-number-format="\\@">' + nipStr + '</td>' +
        '<td style="border:0.5pt solid #cccccc;text-align:center;">' + golStr + '</td>' +
      '</tr>';
    }

    var excelHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><style>body { font-family: Calibri, Arial, sans-serif; font-size: 10pt; } .title { font-size: 13pt; font-weight: bold; } .th-col { background-color: #0f172a; color: #ffffff; font-weight: bold; text-align: center; border: 0.5pt solid #000; }</style></head>' +
      '<body><table><tr><td colspan="9" class="title">PEMERINTAH KABUPATEN TOBA</td></tr><tr><td colspan="9" style="font-size:11pt;font-weight:bold;">BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA</td></tr><tr><td colspan="9" style="color:#64748b;font-style:italic;">Peta Jabatan Struktural & Gap Analysis | Tanggal: ' + tanggalCetak + ' WIB</td></tr><tr></tr>' +
      '<tr><th class="th-col">NO</th><th class="th-col">PERANGKAT DAERAH</th><th class="th-col">UNIT KERJA / UNOR</th><th class="th-col">NOMENKLATUR JABATAN</th><th class="th-col">ESELON</th><th class="th-col">STATUS</th><th class="th-col">PEMANGKU / PLT</th><th class="th-col">NIP</th><th class="th-col">GOL</th></tr>' +
      rowsXml + '</table></body></html>';

    var blob = Utilities.newBlob(excelHtml, 'application/vnd.ms-excel', 'Peta_Jabatan_Struktural_Toba_' + dateFileStr + '.xls');
    return { success: true, filename: blob.getName(), base64: Utilities.base64Encode(blob.getBytes()) };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}