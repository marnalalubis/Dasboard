/**
 * =========================================================================
 * ACRU EMPLOYEE ANALYTICS SYSTEM - KABUPATEN TOBA
 * File: pensiunService.gs
 * Deskripsi: Backend Service Batas Usia Pensiun (BUP) ASN Pemkab Toba.
 * Menghitung usia riil, batas usia pensiun (58/60/65 tahun), TMT pensiun (BKN),
 * sisa masa kerja countdown, multi-filter proyeksi pensiun, serta Export Excel.
 * Bebas dari Date Object serialization crash (100% Primitive Safe).
 * =========================================================================
 */

/**
 * Helper Pembersih Kunci Filter OPD Mandiri (Self-Contained)
 */
function cleanMatchKeyPensiun_(str) {
  if (!str) return '';
  var s = String(str).toUpperCase().trim();
  s = s.replace(/\bSDM\b/g, 'SUMBER DAYA MANUSIA');
  s = s.replace(/\bBKPSDM\b/g, 'BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA');
  s = s.replace(/\bBAPPERIDA\b/g, 'BADAN PERENCANAAN PEMBANGUNAN RISET DAN INOVASI DAERAH');
  s = s.replace(/\bDINKES\b/g, 'DINAS KESEHATAN');
  s = s.replace(/\bDISDIKPORA\b/g, 'DINAS PENDIDIKAN PEMUDA DAN OLAHRAGA');
  return s.replace(/[^A-Z0-9]/g, '');
}

/**
 * Logika Penentuan Batas Usia Pensiun (BUP) berdasarkan Regulasi ASN
 */
function determineBupYears(jenisJabatan, jabatanNama, eselonNama) {
  var jab = String(jabatanNama || '').toUpperCase().trim();
  var esl = String(eselonNama || '').toUpperCase().trim();
  var jns = String(jenisJabatan || '').toUpperCase().trim();

  // 1. BUP 65 Tahun: JF Ahli Utama
  if (jab.indexOf('AHLI UTAMA') !== -1 || (jab.indexOf('UTAMA') !== -1 && jns.indexOf('FUNG') !== -1)) {
    return { bup: 65, kategori: 'BUP 65 Tahun (JF Ahli Utama)' };
  }

  // 2. BUP 60 Tahun:
  // - Pejabat Pimpinan Tinggi (JPT / Eselon II)
  var isEselon2 = (esl.indexOf('II.') !== -1 || esl.indexOf('II ') !== -1 || esl === 'II' || esl === 'ESELON II' || esl === 'ESELON II.A' || esl === 'ESELON II.B');
  // - JF Ahli Madya
  var isAhliMadya = (jab.indexOf('AHLI MADYA') !== -1 || jab.indexOf('MADYA') !== -1);
  // - Guru & Pengawas Sekolah
  var isGuru = (jab.indexOf('GURU') !== -1 || jab.indexOf('PENGAWAS SEKOLAH') !== -1);
  // - Dokter Spesialis / Medik Madya
  var isDokter = (jab.indexOf('DOKTER') !== -1 && (jab.indexOf('MADYA') !== -1 || jab.indexOf('SPESIALIS') !== -1));

  if (isEselon2 || isAhliMadya || isGuru || isDokter) {
    var label = isEselon2 ? 'BUP 60 Tahun (JPT / Eselon II)' : (isGuru ? 'BUP 60 Tahun (Guru / JF Madya)' : 'BUP 60 Tahun (JF Ahli Madya)');
    return { bup: 60, kategori: label };
  }

  // 3. BUP 58 Tahun: Administrator (Eselon III), Pengawas (Eselon IV), Pelaksana, JF Keterampilan & Ahli Pertama/Muda
  var label58 = 'BUP 58 Tahun (Pelaksana / JF Pertama-Muda)';
  if (esl.indexOf('III') !== -1) label58 = 'BUP 58 Tahun (Eselon III - Administrator)';
  else if (esl.indexOf('IV') !== -1) label58 = 'BUP 58 Tahun (Eselon IV - Pengawas)';

  return { bup: 58, kategori: label58 };
}

/**
 * Parser Tanggal Lahir Fleksibel (Mendukung Date Object, String, & Fallback NIP 18 Digit)
 */
function parseTanggalLahirPegawai(rawTanggalLahir, nipBaru) {
  var d = null;

  // 1. Jika sudah bertipe Date object
  if (rawTanggalLahir instanceof Date && !isNaN(rawTanggalLahir.getTime())) {
    return rawTanggalLahir;
  }

  // 2. Parse dari String Tanggal
  if (rawTanggalLahir) {
    var s = String(rawTanggalLahir).trim();
    // Format DD/MM/YYYY atau DD-MM-YYYY
    var dmyMatch = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (dmyMatch) {
      var day = parseInt(dmyMatch[1], 10);
      var month = parseInt(dmyMatch[2], 10) - 1;
      var year = parseInt(dmyMatch[3], 10);
      d = new Date(year, month, day);
      if (!isNaN(d.getTime())) return d;
    }

    // Format YYYY-MM-DD atau YYYY/MM/DD
    var ymdMatch = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (ymdMatch) {
      var y = parseInt(ymdMatch[1], 10);
      var m = parseInt(ymdMatch[2], 10) - 1;
      var dy = parseInt(ymdMatch[3], 10);
      d = new Date(y, m, dy);
      if (!isNaN(d.getTime())) return d;
    }
  }

  // 3. Fallback Cerdas dari 8 Digit Pertama NIP 18 Digit (YYYYMMDD)
  var cleanNip = String(nipBaru || '').replace(/[^0-9]/g, '');
  if (cleanNip.length >= 8) {
    var nYear = parseInt(cleanNip.substring(0, 4), 10);
    var nMonth = parseInt(cleanNip.substring(4, 6), 10) - 1;
    var nDay = parseInt(cleanNip.substring(6, 8), 10);
    if (nYear >= 1940 && nYear <= 2015 && nMonth >= 0 && nMonth <= 11 && nDay >= 1 && nDay <= 31) {
      d = new Date(nYear, nMonth, nDay);
      if (!isNaN(d.getTime())) return d;
    }
  }

  return null;
}

/**
 * Menghitung TMT Pensiun Resmi BKN:
 * Efektif tanggal 1 pada bulan berikutnya setelah tanggal lahir mencapai BUP
 * (Kecuali jika lahir tepat tanggal 1, TMT pensiun berlaku tanggal 1 bulan kelahirannya).
 */
function calculateTmtPensiun(birthDate, bupYears) {
  var bYear = birthDate.getFullYear();
  var bMonth = birthDate.getMonth(); // 0 - 11
  var bDay = birthDate.getDate();

  var retYear = bYear + bupYears;
  var retMonth = bMonth;

  if (bDay === 1) {
    retMonth = bMonth;
  } else {
    retMonth = bMonth + 1;
    if (retMonth > 11) {
      retMonth = 0;
      retYear += 1;
    }
  }

  return new Date(retYear, retMonth, 1);
}

/**
 * Controller Utama Pengambilan Data Pensiun BUP dengan Multi-Filtering
 * Menghasilkan data yang 100% Primitive-Safe (Tanpa objek Date mentah yang merusak RPC)
 */
function getPensiunData(filters) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Spreadsheet aktif tidak ditemukan.");

    filters = filters || {};
    var filterSearch = String(filters.search || '').toLowerCase().trim();
    var filterHorizon = String(filters.horizon || 'TAHUN_INI').toUpperCase().trim();
    var filterTahun = String(filters.tahun || '').trim();
    var filterBup = String(filters.bup || 'ALL').trim();
    var filterJabatan = String(filters.jenisJabatan || 'ALL').toUpperCase().trim();
    var filterOpd = String(filters.opd || '').trim();

    var today = new Date();
    var curYear = today.getFullYear();

    var targetSheets = [
      { name: 'Data_Pegawai_PNS', statusTag: 'PNS' },
      { name: 'Data_Pegawai_PPPK', statusTag: 'PPPK' }
    ];

    var rawList = [];
    var availableYearsSet = {};

    var totalTahunIni = 0;
    var totalKritis = 0;
    var strukturalTahunIni = 0;
    var guruNakesTahunIni = 0;

    for (var s = 0; s < targetSheets.length; s++) {
      var shMeta = targetSheets[s];
      var sheetObj = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, shMeta.name) : ss.getSheetByName(shMeta.name);
      if (!sheetObj || sheetObj.getLastRow() <= 1) continue;

      var rawData = sheetObj.getDataRange().getDisplayValues();
      var headers = rawData[0];

      var idxNip = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["NIP BARU", "NIP"]) : 1;
      var idxNama = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["NAMA", "NAMA LENGKAP"]) : 2;
      var idxGelarD = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["GELAR DEPAN"]) : -1;
      var idxGelarB = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["GELAR BELAKANG"]) : -1;
      var idxTglLahir = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["TANGGAL LAHIR", "TGL LAHIR"]) : 6;
      var idxGol = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["GOL AKHIR NAMA", "GOLONGAN"]) : 9;
      var idxJenisJab = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JENIS JABATAN NAMA", "JENIS JABATAN"]) : 14;
      var idxJabNama = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["JABATAN NAMA", "JABATAN"]) : 15;
      var idxEselon = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["ESELON NAMA", "ESELON"]) : 17;
      var idxSatker = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["SATUAN KERJA KERJA NAMA", "SATKER"]) : 20;
      var idxSatkerInduk = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK"]) : 19;
      var idxUnor = typeof findColumnIndex === 'function' ? findColumnIndex(headers, ["UNOR NAMA", "UNOR"]) : 21;

      for (var r = 1; r < rawData.length; r++) {
        var row = rawData[r];
        if (row.every(function(cell) { return String(cell).trim() === ''; })) continue;

        var nip = idxNip !== -1 ? String(row[idxNip] || '').trim() : '';
        var rawTgl = idxTglLahir !== -1 ? row[idxTglLahir] : '';
        var birthDate = parseTanggalLahirPegawai(rawTgl, nip);
        if (!birthDate) continue;

        var rawJenisJab = idxJenisJab !== -1 ? String(row[idxJenisJab] || '').trim() : '';
        var rawJab = idxJabNama !== -1 ? String(row[idxJabNama] || '').trim() : '';
        var rawEsl = idxEselon !== -1 ? String(row[idxEselon] || '').trim() : '';

        var bupInfo = determineBupYears(rawJenisJab, rawJab, rawEsl);
        var tmtPensiun = calculateTmtPensiun(birthDate, bupInfo.bup);

        var tmtYear = tmtPensiun.getFullYear();
        availableYearsSet[tmtYear] = true;

        // Hitung selisih waktu & bulan ke TMT pensiun
        var diffTime = tmtPensiun.getTime() - today.getTime();
        var diffMonths = (tmtYear - today.getFullYear()) * 12 + (tmtPensiun.getMonth() - today.getMonth());
        var isLewat = diffTime < 0;

        // Hitung Usia Saat Ini (Tahun & Bulan)
        var ageYears = today.getFullYear() - birthDate.getFullYear();
        var ageMonths = today.getMonth() - birthDate.getMonth();
        if (today.getDate() < birthDate.getDate()) {
          ageMonths--;
        }
        if (ageMonths < 0) {
          ageYears--;
          ageMonths += 12;
        }

        var namaMurni = idxNama !== -1 ? String(row[idxNama] || 'Pegawai').trim() : 'Pegawai';
        var gd = idxGelarD !== -1 && row[idxGelarD] ? String(row[idxGelarD]).trim() + ' ' : '';
        var gb = idxGelarB !== -1 && row[idxGelarB] ? ', ' + String(row[idxGelarB]).trim() : '';
        var fullNama = gd + namaMurni + gb;

        var gol = idxGol !== -1 ? String(row[idxGol] || '-').trim() : '-';
        var sInduk = idxSatkerInduk !== -1 ? String(row[idxSatkerInduk] || '').trim() : '';
        var sKerja = idxSatker !== -1 ? String(row[idxSatker] || '').trim() : '';
        var sUnor = idxUnor !== -1 ? String(row[idxUnor] || '').trim() : '';

        var targetOpd = typeof mapUnorToTarget === 'function' ? mapUnorToTarget(sInduk, sKerja, sUnor) : (sKerja || sInduk || sUnor);
        targetOpd = typeof normalizeOpdName === 'function' ? normalizeOpdName(targetOpd) : targetOpd;

        var catMeta = typeof categorizePegawaiJabatanLocal === 'function'
          ? categorizePegawaiJabatanLocal(rawJenisJab, rawJab, rawEsl)
          : { type: (rawEsl && rawEsl !== '-' && rawEsl !== 'NON ESELON') ? 'STRUKTURAL' : 'PELAKSANA' };

        var isGuruOrNakes = rawJab.toUpperCase().indexOf('GURU') !== -1 || rawJab.toUpperCase().indexOf('DOKTER') !== -1 || rawJab.toUpperCase().indexOf('PERAWAT') !== -1 || rawJab.toUpperCase().indexOf('BIDAN') !== -1;

        // Akumulasi KPI
        if (tmtYear === curYear) {
          totalTahunIni++;
          if (catMeta.type === 'STRUKTURAL') strukturalTahunIni++;
          if (isGuruOrNakes) guruNakesTahunIni++;
        }
        if (!isLewat && diffMonths <= 6 && diffMonths >= 0) {
          totalKritis++;
        }

        // Teks Sisa Masa Kerja
        var sisaMasaKerjaText = '';
        if (isLewat) {
          sisaMasaKerjaText = 'Sudah Lewat BUP';
        } else if (diffMonths <= 0) {
          sisaMasaKerjaText = '< 1 Bulan lagi';
        } else if (diffMonths < 12) {
          sisaMasaKerjaText = diffMonths + ' Bulan lagi';
        } else {
          var remYears = Math.floor(diffMonths / 12);
          var remM = diffMonths % 12;
          sisaMasaKerjaText = remYears + ' Thn ' + (remM > 0 ? remM + ' Bln ' : '') + 'lagi';
        }

        // PENTING: Jangan masukkan Date object mentah ke array return (hanya gunakan primitive number & string)
        rawList.push({
          nip: nip,
          nama: fullNama,
          golongan: gol,
          statusAsn: shMeta.statusTag,
          jabatanNama: rawJab,
          eselon: rawEsl || '-',
          kategoriJabatan: catMeta.type || 'PELAKSANA',
          kategoriJabatanLabel: catMeta.label || bupInfo.kategori,
          perangkatDaerah: targetOpd,
          unorNama: sUnor || targetOpd,
          tanggalLahirText: Utilities.formatDate(birthDate, 'Asia/Jakarta', 'dd/MM/yyyy'),
          usiaText: ageYears + ' Thn ' + ageMonths + ' Bln',
          bup: bupInfo.bup,
          bupLabel: bupInfo.kategori,
          tmtPensiunText: Utilities.formatDate(tmtPensiun, 'Asia/Jakarta', 'dd/MM/yyyy'),
          tmtYear: tmtYear,
          tmtTimestamp: tmtPensiun.getTime(), // Primitive number untuk sorting aman
          diffMonths: diffMonths,
          isLewat: isLewat,
          sisaMasaKerjaText: sisaMasaKerjaText,
          isGuruOrNakes: isGuruOrNakes
        });
      }
    }

    // Urutkan daftar tahun pensiun untuk opsi dropdown
    var sortedYears = Object.keys(availableYearsSet).map(function(y) { return parseInt(y, 10); }).sort(function(a, b) { return a - b; });

    // Penyaringan Hasil Sesuai Filter
    var filteredList = rawList.filter(function(item) {
      // 1. Filter Horizon Waktu
      if (filterHorizon === 'KRITIS') {
        if (item.isLewat || item.diffMonths > 6 || item.diffMonths < 0) return false;
      } else if (filterHorizon === 'TAHUN_INI') {
        if (item.tmtYear !== curYear) return false;
      } else if (filterHorizon === 'TAHUN_DEPAN') {
        if (item.tmtYear !== (curYear + 1)) return false;
      } else if (filterHorizon === '2_TAHUN') {
        if (item.isLewat || item.tmtYear > (curYear + 2)) return false;
      } else if (filterHorizon === '3_TAHUN') {
        if (item.isLewat || item.tmtYear > (curYear + 3)) return false;
      } else if (filterHorizon === '5_TAHUN') {
        if (item.isLewat || item.tmtYear > (curYear + 5)) return false;
      } else if (filterHorizon === 'SUDAH_LEWAT') {
        if (!item.isLewat) return false;
      }

      // 2. Filter Tahun Spesifik
      if (filterTahun && String(item.tmtYear) !== filterTahun) {
        return false;
      }

      // 3. Filter BUP
      if (filterBup && filterBup !== 'ALL' && String(item.bup) !== filterBup) {
        return false;
      }

      // 4. Filter Kelompok Jabatan
      if (filterJabatan && filterJabatan !== 'ALL') {
        if (item.kategoriJabatan !== filterJabatan) return false;
      }

      // 5. Filter OPD
      if (filterOpd && filterOpd !== 'ALL' && filterOpd !== '') {
        var cleanFilterOpd = cleanMatchKeyPensiun_(filterOpd);
        var cleanRowOpd = cleanMatchKeyPensiun_(item.perangkatDaerah);
        if (cleanRowOpd !== cleanFilterOpd && cleanRowOpd.indexOf(cleanFilterOpd) === -1 && cleanFilterOpd.indexOf(cleanRowOpd) === -1) {
          return false;
        }
      }

      // 6. Filter Search Query
      if (filterSearch) {
        var mSearch = item.nama.toLowerCase().indexOf(filterSearch) !== -1 ||
                      item.nip.indexOf(filterSearch) !== -1 ||
                      item.jabatanNama.toLowerCase().indexOf(filterSearch) !== -1 ||
                      item.perangkatDaerah.toLowerCase().indexOf(filterSearch) !== -1;
        if (!mSearch) return false;
      }

      return true;
    });

    // Urutkan pegawai dari yang paling dekat TMT pensiunnya (menggunakan timestamp primitif)
    filteredList.sort(function(a, b) {
      return a.tmtTimestamp - b.tmtTimestamp;
    });

    return {
      success: true,
      data: filteredList,
      stats: {
        totalTahunIni: totalTahunIni,
        totalKritis: totalKritis,
        strukturalTahunIni: strukturalTahunIni,
        guruNakesTahunIni: guruNakesTahunIni,
        currentYear: curYear,
        availableYears: sortedYears,
        selectedYear: filterTahun
      }
    };
  } catch (err) {
    Logger.log("Error getPensiunData: " + err.toString());
    return {
      success: false,
      data: [],
      stats: { totalTahunIni: 0, totalKritis: 0, strukturalTahunIni: 0, guruNakesTahunIni: 0, currentYear: new Date().getFullYear(), availableYears: [] },
      error: err.toString()
    };
  }
}

/**
 * Export Daftar Nominatif Pensiun BUP ke Excel (.xls)
 */
function generateExportPensiunExcel(filters) {
  try {
    var res = getPensiunData(filters);
    if (!res || !res.success) throw new Error("Gagal mengambil data pensiun: " + ((res && res.error) || 'Tidak ada data'));

    var rows = res.data || [];
    var tanggalCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    var dateFileStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');

    var rowsXml = '';
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var rowColor = r.diffMonths <= 6 && !r.isLewat ? '#fff1f2' : (i % 2 === 1 ? '#f8fafc' : '#ffffff');
      var sisaColor = r.diffMonths <= 6 && !r.isLewat ? '#ef4444' : (r.diffMonths <= 12 && !r.isLewat ? '#f59e0b' : '#10b981');

      rowsXml += '<tr style="background:' + rowColor + ';">' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;">' + (i + 1) + '</td>' +
        '<td style="border:0.5pt solid #cccccc;font-weight:bold;">' + r.nama + '</td>' +
        '<td style="border:0.5pt solid #cccccc;text-align:center;" mso-number-format="\\@">' + r.nip + '</td>' +
        '<td style="border:0.5pt solid #cccccc;text-align:center;">' + r.golongan + '</td>' +
        '<td style="border:0.5pt solid #cccccc;">' + r.jabatanNama + '</td>' +
        '<td style="border:0.5pt solid #cccccc;text-align:center;">' + r.eselon + '</td>' +
        '<td style="border:0.5pt solid #cccccc;font-weight:bold;">' + r.perangkatDaerah + '</td>' +
        '<td style="border:0.5pt solid #cccccc;">' + r.unorNama + '</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;" mso-number-format="\\@">' + r.tanggalLahirText + '</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;">' + r.usiaText + '</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;font-weight:bold;">' + r.bup + ' Tahun</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;font-weight:bold;color:#0284c7;" mso-number-format="\\@">' + r.tmtPensiunText + '</td>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;font-weight:bold;color:' + sisaColor + ';">' + r.sisaMasaKerjaText + '</td>' +
      '</tr>';
    }

    var excelHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><style>body { font-family: Calibri, Arial, sans-serif; font-size: 10pt; } .title { font-size: 13pt; font-weight: bold; } .th-col { background-color: #0f172a; color: #ffffff; font-weight: bold; text-align: center; border: 0.5pt solid #000; } .th-total { background-color: #0f172a; color: #00ffaa; font-weight: bold; text-align: right; border: 0.5pt solid #000; }</style></head>' +
      '<body><table><tr><td colspan="13" class="title">PEMERINTAH KABUPATEN TOBA</td></tr><tr><td colspan="13" style="font-size:11pt;font-weight:bold;">BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA</td></tr><tr><td colspan="13" style="color:#64748b;font-style:italic;">Daftar Nominatif Pegawai Menjelang Batas Usia Pensiun (BUP) | Tanggal Cetak: ' + tanggalCetak + ' WIB</td></tr><tr></tr>' +
      '<tr><th class="th-col">NO</th><th class="th-col">NAMA PEGAWAI</th><th class="th-col">NIP</th><th class="th-col">GOL</th><th class="th-col">JABATAN</th><th class="th-col">ESELON</th><th class="th-col">PERANGKAT DAERAH</th><th class="th-col">UNIT KERJA (UNOR)</th><th class="th-col">TGL LAHIR</th><th class="th-col">USIA SAAT INI</th><th class="th-col">BUP</th><th class="th-col">TMT PENSIUN</th><th class="th-col">SISA MASA KERJA</th></tr>' +
      rowsXml +
      '<tr><td colspan="4" style="background-color:#0f172a;color:#ffffff;font-weight:bold;text-align:center;">TOTAL NOMINATIF</td><td class="th-total" colspan="9" mso-number-format="#,##0">' + rows.length + ' Pegawai Terdaftar</td></tr>' +
      '</table></body></html>';

    var blob = Utilities.newBlob(excelHtml, 'application/vnd.ms-excel', 'Daftar_Nominatif_Pensiun_BUP_Toba_' + dateFileStr + '.xls');
    return { success: true, filename: blob.getName(), base64: Utilities.base64Encode(blob.getBytes()) };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}