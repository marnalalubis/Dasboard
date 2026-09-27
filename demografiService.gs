/**
 * =========================================================================
 * ACRU EMPLOYEE ANALYTICS SYSTEM - KABUPATEN TOBA
 * File: demografiService.gs
 * Deskripsi: Backend Service Demografi ASN, Piramida Usia Gender,
 * Klasifikasi 4 Generasi (Boomer, Gen X, Millennial, Gen Z),
 * Analisis Rasio Pensiun vs Pengganti (Replacement Ratio 3 Tahun),
 * dan Generator Export Excel Demografi Daerah.
 * =========================================================================
 */

/**
 * Helper parsing tanggal lahir fleksibel dari Date object, String, atau NIP
 */
function parseTanggalLahirDemografi_(rawTgl, nip) {
  if (rawTgl instanceof Date && !isNaN(rawTgl.getTime())) {
    return rawTgl;
  }
  if (rawTgl) {
    var s = String(rawTgl).trim();
    var dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dmy) {
      var d = new Date(parseInt(dmy[3], 10), parseInt(dmy[2], 10) - 1, parseInt(dmy[1], 10));
      if (!isNaN(d.getTime())) return d;
    }
    var ymd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (ymd) {
      var d2 = new Date(parseInt(ymd[1], 10), parseInt(ymd[2], 10) - 1, parseInt(ymd[3], 10));
      if (!isNaN(d2.getTime())) return d2;
    }
  }
  var cleanNip = String(nip || '').replace(/[^0-9]/g, '');
  if (cleanNip.length >= 8) {
    var y = parseInt(cleanNip.substring(0, 4), 10);
    var m = parseInt(cleanNip.substring(4, 6), 10) - 1;
    var dy = parseInt(cleanNip.substring(6, 8), 10);
    if (y >= 1940 && y <= 2015 && m >= 0 && m <= 11 && dy >= 1 && dy <= 31) {
      var d3 = new Date(y, m, dy);
      if (!isNaN(d3.getTime())) return d3;
    }
  }
  return null;
}

/**
 * Mengelompokkan generasi berdasarkan tahun kelahiran
 */
function getGenerasiByBirthYear_(birthYear) {
  if (birthYear >= 1997) {
    return { code: 'GEN_Z', label: 'Gen Z (≤ 29 Thn)', color: '#10b981', badgeClass: 'badge-gen-z' };
  } else if (birthYear >= 1981) {
    return { code: 'MILLENNIAL', label: 'Millennials (30 - 45 Thn)', color: '#0284c7', badgeClass: 'badge-millennial' };
  } else if (birthYear >= 1965) {
    return { code: 'GEN_X', label: 'Gen X (46 - 61 Thn)', color: '#f59e0b', badgeClass: 'badge-gen-x' };
  } else {
    return { code: 'BOOMER', label: 'Baby Boomers (≥ 62 Thn)', color: '#ef4444', badgeClass: 'badge-boomer' };
  }
}

/**
 * Controller Utama Pengambilan Data Demografi & Piramida Usia ASN
 */
function getDemografiData(filters) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Spreadsheet aktif tidak ditemukan.");

    filters = filters || {};
    var filterSearch = String(filters.search || '').toLowerCase().trim();
    var filterOpd = String(filters.opd || '').trim();
    var filterAsn = String(filters.jenisAsn || 'ALL').toUpperCase().trim();
    var filterJabatan = String(filters.jenisJabatan || 'ALL').toUpperCase().trim();
    var filterGenerasi = String(filters.generasi || 'ALL').toUpperCase().trim();

    var targetSheets = [];
    if (filterAsn === 'ALL' || filterAsn === 'PNS') targetSheets.push({ name: 'Data_Pegawai_PNS', tag: 'PNS' });
    if (filterAsn === 'ALL' || filterAsn === 'PPPK') targetSheets.push({ name: 'Data_Pegawai_PPPK', tag: 'PPPK' });
    if (filterAsn === 'ALL' || filterAsn === 'PPPK_PARUH_WAKTU') targetSheets.push({ name: 'Data_Pegawai_PPPK_Paruh_Waktu', tag: 'PPPK_PW' });

    var today = new Date();
    var curYear = today.getFullYear();

    // Inisialisasi Kelompok Piramida Usia (Pria vs Wanita)
    var pyramidBrackets = [
      { key: 'UNDER_30', label: '< 30 Tahun', min: 0, max: 29, male: 0, female: 0, total: 0 },
      { key: '30_39', label: '30 - 39 Tahun', min: 30, max: 39, male: 0, female: 0, total: 0 },
      { key: '40_49', label: '40 - 49 Tahun', min: 40, max: 49, male: 0, female: 0, total: 0 },
      { key: '50_54', label: '50 - 54 Tahun', min: 50, max: 54, male: 0, female: 0, total: 0 },
      { key: '55_57', label: '55 - 57 Tahun', min: 55, max: 57, male: 0, female: 0, total: 0 },
      { key: '58_PLUS', label: '≥ 58 Tahun', min: 58, max: 120, male: 0, female: 0, total: 0 }
    ];

    var genStats = {
      GEN_Z: { count: 0, label: 'Gen Z', desc: 'Lahir 1997 ke atas', color: '#10b981' },
      MILLENNIAL: { count: 0, label: 'Millennials (Gen Y)', desc: 'Lahir 1981 - 1996', color: '#0284c7' },
      GEN_X: { count: 0, label: 'Gen X', desc: 'Lahir 1965 - 1980', color: '#f59e0b' },
      BOOMER: { count: 0, label: 'Baby Boomers', desc: 'Lahir s.d 1964', color: '#ef4444' }
    };

    var opdMap = {};
    var allAges = [];
    var retiringCount3Years = 0; // ASN yang akan pensiun dalam 3 tahun ke depan (usia >= 55)
    var juniorWorkforceCount = 0; // ASN generasi baru (usia <= 32)
    var totalProcessed = 0;

    for (var s = 0; s < targetSheets.length; s++) {
      var shMeta = targetSheets[s];
      var sh = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, shMeta.name) : ss.getSheetByName(shMeta.name);
      if (!sh || sh.getLastRow() <= 1) continue;

      var rawData = sh.getDataRange().getDisplayValues();
      var headers = rawData[0];

      var idxNip = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["NIP BARU", "NIP"]) : 1;
      var idxNama = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["NAMA", "NAMA LENGKAP"]) : 2;
      var idxTglLahir = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["TANGGAL LAHIR", "TGL LAHIR"]) : 6;
      var idxGender = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JENIS KELAMIN", "KELAMIN", "GENDER", "JK"]) : 7;
      var idxGol = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["GOL AKHIR NAMA", "GOLONGAN"]) : 9;
      var idxJenisJab = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JENIS JABATAN NAMA", "JENIS JABATAN"]) : 14;
      var idxJabNama = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JABATAN NAMA", "JABATAN"]) : 15;
      var idxEselon = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["ESELON NAMA", "ESELON"]) : 17;
      var idxSatkerInduk = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK"]) : 19;
      var idxSatker = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["SATUAN KERJA KERJA NAMA", "SATKER"]) : 20;
      var idxUnor = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["UNOR NAMA", "UNOR"]) : 21;

      for (var r = 1; r < rawData.length; r++) {
        var row = rawData[r];
        if (row.every(function(cell) { return String(cell).trim() === ''; })) continue;

        var nip = idxNip !== -1 ? String(row[idxNip] || '').trim() : '';
        var rawTgl = idxTglLahir !== -1 ? row[idxTglLahir] : '';
        var birthDate = parseTanggalLahirDemografi_(rawTgl, nip);
        if (!birthDate) continue;

        var birthYear = birthDate.getFullYear();
        var ageYears = curYear - birthYear;
        if (today.getMonth() < birthDate.getMonth() || (today.getMonth() === birthDate.getMonth() && today.getDate() < birthDate.getDate())) {
          ageYears--;
        }
        if (ageYears < 18 || ageYears > 75) continue;

        var rawGender = idxGender !== -1 ? String(row[idxGender] || '').toUpperCase().trim() : '';
        var isFemale = rawGender.indexOf('PEREMPUAN') !== -1 || rawGender.indexOf('WANITA') !== -1 || rawGender === 'P' || rawGender === 'F';
        var genderCode = isFemale ? 'PEREMPUAN' : 'LAKI-LAKI';

        var sInduk = idxSatkerInduk !== -1 ? String(row[idxSatkerInduk] || '').trim() : '';
        var sKerja = idxSatker !== -1 ? String(row[idxSatker] || '').trim() : '';
        var sUnor = idxUnor !== -1 ? String(row[idxUnor] || '').trim() : '';
        var targetOpd = typeof mapUnorToTarget === 'function' ? mapUnorToTarget(sInduk, sKerja, sUnor) : (sKerja || sInduk || sUnor);
        targetOpd = typeof normalizeOpdName === 'function' ? normalizeOpdName(targetOpd) : targetOpd;

        var rawJenisJab = idxJenisJab !== -1 ? String(row[idxJenisJab] || '').trim() : '';
        var rawJab = idxJabNama !== -1 ? String(row[idxJabNama] || '').trim() : '';
        var rawEsl = idxEselon !== -1 ? String(row[idxEselon] || '').trim() : '';
        var catMeta = typeof categorizePegawaiJabatanLocal === 'function'
          ? categorizePegawaiJabatanLocal(rawJenisJab, rawJab, rawEsl)
          : { type: (rawEsl && rawEsl !== '-' && rawEsl !== 'NON ESELON') ? 'STRUKTURAL' : 'PELAKSANA' };

        var genInfo = getGenerasiByBirthYear_(birthYear);

        // Filter Pengecekan
        if (filterOpd && filterOpd !== 'ALL' && filterOpd !== '') {
          var cleanFilter = typeof cleanMatchKey === 'function' ? cleanMatchKey(filterOpd) : filterOpd.toUpperCase().replace(/[^A-Z0-9]/g, '');
          var cleanRow = typeof cleanMatchKey === 'function' ? cleanMatchKey(targetOpd) : targetOpd.toUpperCase().replace(/[^A-Z0-9]/g, '');
          if (cleanRow.indexOf(cleanFilter) === -1 && cleanFilter.indexOf(cleanRow) === -1) continue;
        }

        if (filterJabatan && filterJabatan !== 'ALL' && filterJabatan !== '') {
          if (catMeta.type !== filterJabatan) continue;
        }

        if (filterGenerasi && filterGenerasi !== 'ALL' && filterGenerasi !== '') {
          if (genInfo.code !== filterGenerasi) continue;
        }

        if (filterSearch) {
          var namaStr = idxNama !== -1 ? String(row[idxNama] || '').toLowerCase() : '';
          var jabStr = rawJab.toLowerCase();
          var opdStr = targetOpd.toLowerCase();
          if (namaStr.indexOf(filterSearch) === -1 && nip.indexOf(filterSearch) === -1 && jabStr.indexOf(filterSearch) === -1 && opdStr.indexOf(filterSearch) === -1) {
            continue;
          }
        }

        totalProcessed++;
        allAges.push(ageYears);

        // Akumulasi Generasi
        if (genStats[genInfo.code]) {
          genStats[genInfo.code].count++;
        }

        // Akumulasi Piramida Usia
        for (var b = 0; b < pyramidBrackets.length; b++) {
          var brk = pyramidBrackets[b];
          if (ageYears >= brk.min && ageYears <= brk.max) {
            if (isFemale) brk.female++; else brk.male++;
            brk.total++;
            break;
          }
        }

        // Akumulasi Rasio Pensiun vs Pengganti (Senior vs Junior)
        if (ageYears >= 55) {
          retiringCount3Years++;
        }
        if (ageYears <= 32) {
          juniorWorkforceCount++;
        }

        // Akumulasi per OPD
        if (!opdMap[targetOpd]) {
          opdMap[targetOpd] = {
            nama: targetOpd,
            total: 0,
            sumAge: 0,
            genZ: 0,
            millennial: 0,
            genX: 0,
            boomer: 0
          };
        }
        var opdObj = opdMap[targetOpd];
        opdObj.total++;
        opdObj.sumAge += ageYears;
        if (genInfo.code === 'GEN_Z') opdObj.genZ++;
        else if (genInfo.code === 'MILLENNIAL') opdObj.millennial++;
        else if (genInfo.code === 'GEN_X') opdObj.genX++;
        else if (genInfo.code === 'BOOMER') opdObj.boomer++;
      }
    }

    // Hitung Rata-rata Usia Keseluruhan
    var avgAgeTotal = allAges.length > 0 ? (allAges.reduce(function(acc, val) { return acc + val; }, 0) / allAges.length).toFixed(1) : '0';

    // Hitung Rasio Pengganti (Replacement Ratio)
    // Formula: (Tenaga Junior Masuk / Tenaga Senior yang Akan Pensiun 3 Tahun ke Depan) * 100%
    var replacementRatio = retiringCount3Years > 0 ? Math.round((juniorWorkforceCount / retiringCount3Years) * 100) : 100;
    var replacementStatus = 'IDEAL';
    var replacementNote = 'Regenerasi birokrasi seimbang dan sehat.';

    if (replacementRatio < 60) {
      replacementStatus = 'DEFISIT_KRITIS';
      replacementNote = 'Gelombang pensiun 3 tahun ke depan melebihi jumlah tenaga muda masuk. Diperlukan penambahan kuota formasi CASN/PPPK.';
    } else if (replacementRatio < 90) {
      replacementStatus = 'WASPADA';
      replacementNote = 'Peremajaan birokrasi moderat. Pertahankan pengusulan formasi berkala.';
    } else if (replacementRatio > 140) {
      replacementStatus = 'SURPLUS';
      replacementNote = 'Jumlah tenaga muda sangat tinggi, regenerasi jangka panjang sangat aman.';
    }

    // Rekomendasi Tambahan Rekrutmen Tahunan
    var deficitCount = Math.max(0, retiringCount3Years - juniorWorkforceCount);
    var annualRecruitmentAdvice = Math.ceil(deficitCount / 3) + Math.ceil(retiringCount3Years / 3);

    // Format Data Tabel OPD
    var opdTableRows = [];
    var orderedList = [];
    var seenOpd = {};

    if (typeof TARGET_ORDER_PD !== 'undefined' && Array.isArray(TARGET_ORDER_PD)) {
      for (var tIdx = 0; tIdx < TARGET_ORDER_PD.length; tIdx++) {
        var canonicalName = TARGET_ORDER_PD[tIdx];
        if (opdMap[canonicalName]) {
          orderedList.push(canonicalName);
          seenOpd[canonicalName] = true;
        }
      }
    }

    var allKeys = Object.keys(opdMap).sort();
    for (var ok = 0; ok < allKeys.length; ok++) {
      if (!seenOpd[allKeys[ok]]) {
        orderedList.push(allKeys[ok]);
      }
    }

    for (var k = 0; k < orderedList.length; k++) {
      var item = opdMap[orderedList[k]];
      var avgOpd = item.total > 0 ? (item.sumAge / item.total).toFixed(1) : '0';
      
      // Deteksi Generasi Dominan di OPD
      var dominantGen = 'Millennials';
      var maxCnt = item.millennial;
      if (item.genX > maxCnt) { dominantGen = 'Gen X'; maxCnt = item.genX; }
      if (item.genZ > maxCnt) { dominantGen = 'Gen Z'; maxCnt = item.genZ; }
      if (item.boomer > maxCnt) { dominantGen = 'Baby Boomers'; maxCnt = item.boomer; }

      opdTableRows.push({
        no: k + 1,
        nama: item.nama,
        total: item.total,
        avgAge: avgOpd,
        genZ: item.genZ,
        millennial: item.millennial,
        genX: item.genX,
        boomer: item.boomer,
        dominantGen: dominantGen
      });
    }

    return {
      success: true,
      totalPegawai: totalProcessed,
      avgAgeTotal: avgAgeTotal,
      replacement: {
        ratio: replacementRatio,
        status: replacementStatus,
        note: replacementNote,
        retiring3Years: retiringCount3Years,
        juniorWorkforce: juniorWorkforceCount,
        annualAdvice: annualRecruitmentAdvice
      },
      generations: genStats,
      pyramid: pyramidBrackets,
      opdBreakdown: opdTableRows
    };
  } catch (err) {
    Logger.log("Error getDemografiData: " + err.toString());
    return {
      success: false,
      error: err.toString(),
      totalPegawai: 0,
      avgAgeTotal: '0',
      replacement: { ratio: 0, status: 'ERROR', retiring3Years: 0, juniorWorkforce: 0, annualAdvice: 0 },
      generations: {},
      pyramid: [],
      opdBreakdown: []
    };
  }
}

/**
 * Export Ringkasan Demografi & Generasi ke Excel (.xls)
 */
function generateExportDemografiExcel(filters) {
  try {
    var res = getDemografiData(filters);
    if (!res || !res.success) throw new Error("Gagal mengambil data demografi: " + ((res && res.error) || 'Tidak ada data'));

    var tanggalCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var dateFileStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');
    var rows = res.opdBreakdown || [];

    var rowsXml = '';
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      rowsXml += '<tr>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;">' + r.no + '</td>' +
        '<td style="border:0.5pt solid #cccccc;font-weight:bold;">' + r.nama + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + r.total + '</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;font-weight:bold;color:#0284c7;">' + r.avgAge + ' Thn</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + r.genZ + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + r.millennial + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + r.genX + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + r.boomer + '</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;font-weight:bold;">' + r.dominantGen + '</td>' +
      '</tr>';
    }

    var excelHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><style>body { font-family: Calibri, Arial, sans-serif; font-size: 10pt; } .title { font-size: 13pt; font-weight: bold; } .th-col { background-color: #0f172a; color: #ffffff; font-weight: bold; text-align: center; border: 0.5pt solid #000; } .th-total { background-color: #0f172a; color: #00ffaa; font-weight: bold; text-align: right; border: 0.5pt solid #000; }</style></head>' +
      '<body><table><tr><td colspan="9" class="title">PEMERINTAH KABUPATEN TOBA</td></tr><tr><td colspan="9" style="font-size:11pt;font-weight:bold;">BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA</td></tr><tr><td colspan="9" style="color:#64748b;font-style:italic;">Analisis Demografi & Generasi ASN | Tanggal Cetak: ' + tanggalCetak + ' WIB</td></tr><tr></tr>' +
      '<tr><td colspan="9" style="font-weight:bold;background:#f1f5f9;border:0.5pt solid #cccccc;">Rasio Pengganti (Replacement Ratio): ' + res.replacement.ratio + '% (' + res.replacement.note + ') | Rekomendasi Rekrutmen: ~' + res.replacement.annualAdvice + ' ASN/Tahun</td></tr>' +
      '<tr></tr>' +
      '<tr><th class="th-col">NO</th><th class="th-col">PERANGKAT DAERAH / KECAMATAN</th><th class="th-col">TOTAL PEGAWAI</th><th class="th-col">USIA RATA-RATA</th><th class="th-col">GEN Z (&le;29 Thn)</th><th class="th-col">MILLENNIALS (30-45 Thn)</th><th class="th-col">GEN X (46-61 Thn)</th><th class="th-col">BOOMERS (&ge;62 Thn)</th><th class="th-col">GENERASI DOMINAN</th></tr>' +
      rowsXml +
      '<tr><td colspan="2" style="background-color:#0f172a;color:#ffffff;font-weight:bold;text-align:center;">TOTAL KESELURUHAN</td><td class="th-total" mso-number-format="#,##0">' + res.totalPegawai + '</td><td class="th-total" style="text-align:center;">' + res.avgAgeTotal + ' Thn</td><td class="th-total" mso-number-format="#,##0">' + ((res.generations && res.generations.GEN_Z && res.generations.GEN_Z.count) || 0) + '</td><td class="th-total" mso-number-format="#,##0">' + ((res.generations && res.generations.MILLENNIAL && res.generations.MILLENNIAL.count) || 0) + '</td><td class="th-total" mso-number-format="#,##0">' + ((res.generations && res.generations.GEN_X && res.generations.GEN_X.count) || 0) + '</td><td class="th-total" mso-number-format="#,##0">' + ((res.generations && res.generations.BOOMER && res.generations.BOOMER.count) || 0) + '</td><td class="th-total" style="text-align:center;">-</td></tr>' +
      '</table></body></html>';

    var blob = Utilities.newBlob(excelHtml, 'application/vnd.ms-excel', 'Analisis_Demografi_Generasi_ASN_Toba_' + dateFileStr + '.xls');
    return { success: true, filename: blob.getName(), base64: Utilities.base64Encode(blob.getBytes()) };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}