/**
 * =========================================================================
 * ACRU EMPLOYEE ANALYTICS SYSTEM - KABUPATEN TOBA
 * File: rekapJenisJabatanService.gs
 * Deskripsi: Kategorisasi posisi (Struktural, Fungsional, Pelaksana),
 * pemetaan jenjang karier JF (Keahlian & Keterampilan), serta Cetak PDF & Excel.
 * Pasangan Komponen Frontend: subRekapJenisJabatan.html
 * =========================================================================
 */

function categorizeJenisJabatan(rawJenis, rawJabatan, rawEselon) {
  var j = String(rawJenis || '').toUpperCase().trim();
  var jab = String(rawJabatan || '').toUpperCase().trim();
  var esl = String(rawEselon || '').toUpperCase().trim();

  // 1. Deteksi Jabatan Struktural (Manajerial)
  var hasEselon = esl && esl !== '-' && esl !== 'NON ESELON' && esl !== 'NON' && esl !== '0';
  var isStrukturalText = j.indexOf('STRUKTURAL') !== -1 || j.indexOf('PIMPINAN') !== -1 || j.indexOf('ADMINISTRATOR') !== -1 || j.indexOf('PENGAWAS') !== -1;
  
  if (hasEselon || isStrukturalText) {
    return {
      type: 'STRUKTURAL',
      jfCategory: '-',
      jfLevel: '-'
    };
  }

  // 2. Deteksi Jabatan Fungsional (JF)
  var isFungsionalText = j.indexOf('FUNGSIONAL') !== -1 || j.indexOf('JFT') !== -1 || j.indexOf('TERTENTU') !== -1;
  
  // Kata kunci khas jenjang Jabatan Fungsional
  var isJfLevelKeahlian = jab.indexOf('AHLI UTAMA') !== -1 || jab.indexOf('AHLI MADYA') !== -1 || jab.indexOf('AHLI MUDA') !== -1 || jab.indexOf('AHLI PERTAMA') !== -1;
  var isJfLevelKeterampilan = jab.indexOf('PENYELIA') !== -1 || jab.indexOf('MAHIR') !== -1 || jab.indexOf('PELAKSANA LANJUTAN') !== -1 || jab.indexOf('TERAMPIL') !== -1 || jab.indexOf('PELAKSANA') !== -1 && isFungsionalText || jab.indexOf('PEMULA') !== -1;
  var isJfProfesi = jab.indexOf('GURU') !== -1 || jab.indexOf('DOKTER') !== -1 || jab.indexOf('BIDAN') !== -1 || jab.indexOf('PERAWAT') !== -1 || jab.indexOf('AUDITOR') !== -1 || jab.indexOf('PRANATA') !== -1 || jab.indexOf('PENGAWAS SEKOLAH') !== -1 || jab.indexOf('PAMONG') !== -1 || jab.indexOf('PENYULUH') !== -1 || jab.indexOf('EPIDEMIOLOG') !== -1 || jab.indexOf('SANITARIAN') !== -1 || jab.indexOf('NUTRISIONIS') !== -1 || jab.indexOf('APOTEKER') !== -1 || jab.indexOf('ASISTEN APOTEKER') !== -1;

  if (isFungsionalText || isJfLevelKeahlian || isJfLevelKeterampilan || isJfProfesi) {
    var jfCat = 'KEAHLIAN';
    var jfLvl = 'Ahli Pertama';

    // Sub-Klasifikasi Jenjang Keahlian
    if (jab.indexOf('AHLI UTAMA') !== -1 || jab.indexOf('UTAMA') !== -1 && isFungsionalText) {
      jfCat = 'KEAHLIAN'; jfLvl = 'Ahli Utama';
    } else if (jab.indexOf('AHLI MADYA') !== -1 || jab.indexOf('MADYA') !== -1) {
      jfCat = 'KEAHLIAN'; jfLvl = 'Ahli Madya';
    } else if (jab.indexOf('AHLI MUDA') !== -1 || jab.indexOf('MUDA') !== -1) {
      jfCat = 'KEAHLIAN'; jfLvl = 'Ahli Muda';
    } else if (jab.indexOf('AHLI PERTAMA') !== -1 || jab.indexOf('PERTAMA') !== -1) {
      jfCat = 'KEAHLIAN'; jfLvl = 'Ahli Pertama';
    }
    // Sub-Klasifikasi Jenjang Keterampilan / Vokasi
    else if (jab.indexOf('PENYELIA') !== -1) {
      jfCat = 'KETERAMPILAN'; jfLvl = 'Penyelia';
    } else if (jab.indexOf('MAHIR') !== -1 || jab.indexOf('PELAKSANA LANJUTAN') !== -1) {
      jfCat = 'KETERAMPILAN'; jfLvl = 'Mahir';
    } else if (jab.indexOf('TERAMPIL') !== -1) {
      jfCat = 'KETERAMPILAN'; jfLvl = 'Terampil';
    } else if (jab.indexOf('PEMULA') !== -1) {
      jfCat = 'KETERAMPILAN'; jfLvl = 'Pemula';
    } else {
      jfCat = 'KEAHLIAN'; jfLvl = 'Ahli Pertama';
    }

    return {
      type: 'FUNGSIONAL',
      jfCategory: jfCat,
      jfLevel: jfLvl
    };
  }

  // 3. Sisanya adalah Jabatan Pelaksana (JFU / Administrasi)
  return {
    type: 'PELAKSANA',
    jfCategory: '-',
    jfLevel: '-'
  };
}

function getRekapJenisJabatanData(filterAsn) {
  try {
    if (typeof ensureDatabaseReady === 'function') ensureDatabaseReady();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    filterAsn = filterAsn || 'ALL';

    const sheets = [];
    if (filterAsn === 'ALL' || filterAsn === 'PNS') sheets.push({ name: 'Data_Pegawai_PNS', type: 'PNS' });
    if (filterAsn === 'ALL' || filterAsn === 'PPPK') sheets.push({ name: 'Data_Pegawai_PPPK', type: 'PPPK' });
    if (filterAsn === 'ALL' || filterAsn === 'PPPK_PARUH_WAKTU') sheets.push({ name: 'Data_Pegawai_PPPK_Paruh_Waktu', type: 'PPPK_PW' });

    let rekapMap = {};
    TARGET_ORDER_PD.forEach(function(pd) {
      var norm = normalizeOpdName(pd);
      rekapMap[norm] = {
        nama: norm,
        struktural: 0,
        fungsional: 0,
        pelaksana: 0,
        jf_utama: 0,
        jf_madya: 0,
        jf_muda: 0,
        jf_pertama: 0,
        jf_penyelia: 0,
        jf_mahir: 0,
        jf_terampil: 0,
        jf_pemula: 0
      };
    });

    sheets.forEach(function(shInfo) {
      const sheet = ss.getSheetByName(shInfo.name);
      if (!sheet || sheet.getLastRow() <= 1) return;

      const data = sheet.getDataRange().getDisplayValues();
      if (data.length <= 1) return;

      const headers = data[0];
      const idxJenis = findColumnIndex(headers, ["JENIS JABATAN NAMA", "JENIS JABATAN", "JENIS_JABATAN"]);
      const idxJabatan = findColumnIndex(headers, ["JABATAN NAMA", "JABATAN", "NAMA JABATAN"]);
      const idxEselon = findColumnIndex(headers, ["ESELON NAMA", "ESELON", "ESELON_NAMA", "TINGKAT ESELON"]);
      const idxUnor = findColumnIndex(headers, ["UNOR NAMA", "UNOR", "UNIT ORGANISASI", "UNIT KERJA"]);
      const idxSatker = findColumnIndex(headers, ["SATUAN KERJA KERJA NAMA", "SATKER", "SATUAN KERJA", "SATUAN KERJA NAMA"]);
      const idxSatkerInduk = findColumnIndex(headers, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK", "INDUK"]);

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.every(function(cell) { return String(cell).trim() === ''; })) continue;

        let satkerIndukVal = idxSatkerInduk !== -1 ? row[idxSatkerInduk] : "";
        let satkerVal = idxSatker !== -1 ? row[idxSatker] : "";
        let unorVal = idxUnor !== -1 ? row[idxUnor] : "";
        let rawJenisVal = idxJenis !== -1 ? row[idxJenis] : "";
        let rawJabVal = idxJabatan !== -1 ? row[idxJabatan] : "";
        let rawEslVal = idxEselon !== -1 ? row[idxEselon] : "";

        let targetCat = normalizeOpdName(mapUnorToTarget(satkerIndukVal, satkerVal, unorVal));
        if (targetCat && rekapMap[targetCat]) {
          let pos = categorizeJenisJabatan(rawJenisVal, rawJabVal, rawEslVal);
          let targetObj = rekapMap[targetCat];

          if (pos.type === 'STRUKTURAL') {
            targetObj.struktural++;
          } else if (pos.type === 'FUNGSIONAL') {
            targetObj.fungsional++;

            switch (pos.jfLevel) {
              case 'Ahli Utama': targetObj.jf_utama++; break;
              case 'Ahli Madya': targetObj.jf_madya++; break;
              case 'Ahli Muda': targetObj.jf_muda++; break;
              case 'Ahli Pertama': targetObj.jf_pertama++; break;
              case 'Penyelia': targetObj.jf_penyelia++; break;
              case 'Mahir': targetObj.jf_mahir++; break;
              case 'Terampil': targetObj.jf_terampil++; break;
              case 'Pemula': targetObj.jf_pemula++; break;
              default: targetObj.jf_pertama++; break;
            }
          } else {
            targetObj.pelaksana++;
          }
        }
      }
    });

    let finalData = [];
    let counter = 1;
    let totals = {
      struktural: 0,
      fungsional: 0,
      pelaksana: 0,
      total_all: 0,
      jf_levels: {
        utama: 0,
        madya: 0,
        muda: 0,
        pertama: 0,
        penyelia: 0,
        mahir: 0,
        terampil: 0,
        pemula: 0,
        total_keahlian: 0,
        total_keterampilan: 0
      }
    };

    TARGET_ORDER_PD.forEach(function(cat) {
      let cleanCat = normalizeOpdName(cat);
      let r = rekapMap[cleanCat];
      let no = cleanCat === "RSUD Porsea" ? "" : counter++;
      let jumlah = r.struktural + r.fungsional + r.pelaksana;

      totals.struktural += r.struktural;
      totals.fungsional += r.fungsional;
      totals.pelaksana += r.pelaksana;
      totals.total_all += jumlah;

      totals.jf_levels.utama += r.jf_utama;
      totals.jf_levels.madya += r.jf_madya;
      totals.jf_levels.muda += r.jf_muda;
      totals.jf_levels.pertama += r.jf_pertama;
      totals.jf_levels.penyelia += r.jf_penyelia;
      totals.jf_levels.mahir += r.jf_mahir;
      totals.jf_levels.terampil += r.jf_terampil;
      totals.jf_levels.pemula += r.jf_pemula;

      finalData.push({
        no: no,
        nama: cleanCat,
        struktural: r.struktural,
        fungsional: r.fungsional,
        pelaksana: r.pelaksana,
        jumlah: jumlah,
        jf_detail: {
          utama: r.jf_utama,
          madya: r.jf_madya,
          muda: r.jf_muda,
          pertama: r.jf_pertama,
          penyelia: r.jf_penyelia,
          mahir: r.jf_mahir,
          terampil: r.jf_terampil,
          pemula: r.jf_pemula
        }
      });
    });

    totals.jf_levels.total_keahlian = totals.jf_levels.utama + totals.jf_levels.madya + totals.jf_levels.muda + totals.jf_levels.pertama;
    totals.jf_levels.total_keterampilan = totals.jf_levels.penyelia + totals.jf_levels.mahir + totals.jf_levels.terampil + totals.jf_levels.pemula;

    return {
      success: true,
      data: finalData,
      totals: totals
    };
  } catch (e) {
    Logger.log("Error getRekapJenisJabatanData: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function generateRekapJenisJabatanPDF(filterAsn) {
  try {
    const res = getRekapJenisJabatanData(filterAsn);
    if (!res || !res.success) {
      throw new Error("Gagal mengambil data rekapitulasi jenis jabatan: " + (res ? res.error : ""));
    }

    const t = res.totals || {};
    const jf = t.jf_levels || {};
    const rows = res.data || [];
    const tanggalCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');
    const waktuCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'HH:mm:ss');
    const dateFileStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');

    let filterLabel = "Semua Status ASN (PNS, PPPK & PPPK Paruh Waktu)";
    if (filterAsn === 'PNS') filterLabel = "Hanya Pegawai Negeri Sipil (PNS)";
    else if (filterAsn === 'PPPK') filterLabel = "Hanya PPPK Penuh Waktu";
    else if (filterAsn === 'PPPK_PARUH_WAKTU') filterLabel = "Hanya PPPK Paruh Waktu";

    let tableRowsHtml = '';
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const displayName = normalizeOpdName(r.nama);
      const isRsud = displayName === 'RSUD Porsea';
      const bgStyle = isRsud ? 'background-color: #f0fdf4;' : (i % 2 === 1 ? 'background-color: #f8fafc;' : 'background-color: #ffffff;');

      tableRowsHtml += '<tr style="' + bgStyle + '">' +
        '<td style="text-align: center; border: 1px solid #cbd5e1; padding: 4px 3px; font-size: 7.5pt;">' + (r.no || '-') + '</td>' +
        '<td style="border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; font-weight: bold; color: #1e293b;">' + displayName + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; color: #7c3aed; font-weight: 600;">' + (r.struktural || '-') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; color: #0284c7; font-weight: 600;">' + (r.fungsional || '-') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; color: #d97706; font-weight: 600;">' + (r.pelaksana || '-') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; font-weight: bold; background-color: #e2e8f0; color: #0f172a;">' + Number(r.jumlah || 0).toLocaleString('id-ID') + '</td>' +
      '</tr>';
    }

    const htmlContent = '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
        '<meta charset="UTF-8">' +
        '<style>' +
          '@page { size: A4 landscape; margin: 8mm 10mm 10mm 10mm; }' +
          'body { font-family: Arial, Helvetica, sans-serif; color: #0f172a; margin: 0; padding: 0; }' +
          '.kop-wrapper { text-align: center; margin-bottom: 6px; }' +
          '.kop-line-1 { font-size: 11pt; font-weight: bold; letter-spacing: 0.5px; text-transform: uppercase; margin: 0; }' +
          '.kop-line-2 { font-size: 13pt; font-weight: bold; letter-spacing: 0.5px; text-transform: uppercase; margin: 2px 0; }' +
          '.kop-line-3 { font-size: 8.5pt; color: #334155; margin: 0; }' +
          '.kop-divider-double { border: 0; border-top: 2px solid #000; border-bottom: 0.5px solid #000; height: 3px; margin: 6px 0 10px 0; }' +
          '.doc-title { text-align: center; margin-bottom: 8px; }' +
          '.doc-title h3 { margin: 0; font-size: 11pt; font-weight: bold; text-transform: uppercase; text-decoration: underline; }' +
          '.doc-title p { margin: 3px 0 0 0; font-size: 8pt; color: #475569; }' +
          'table.rekap { width: 100%; border-collapse: collapse; margin-bottom: 12px; }' +
          'table.rekap th { background-color: #0f172a; color: #ffffff; border: 1px solid #0f172a; padding: 5px 4px; font-size: 8pt; text-transform: uppercase; text-align: center; }' +
          '.total-row td { background-color: #0f172a !important; color: #ffffff !important; font-weight: bold; padding: 5px 4px; border: 1px solid #0f172a; }' +
          '.ttd-container { width: 100%; margin-top: 10px; page-break-inside: avoid; }' +
          '.ttd-box { float: right; width: 240px; text-align: center; font-size: 8pt; line-height: 1.35; }' +
          '.ttd-space { height: 42px; }' +
          '.footer-info { font-size: 6.5pt; color: #94a3b8; font-style: italic; margin-top: 15px; clear: both; }' +
        '</style>' +
      '</head>' +
      '<body>' +
        '<div class="kop-wrapper">' +
          '<div class="kop-line-1">PEMERINTAH KABUPATEN TOBA</div>' +
          '<div class="kop-line-2">BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA</div>' +
          '<div class="kop-line-3">Jln.AB. Silalahi, Desa Sianipar Sihail-hail, Balige (22311), Sumatera Utara</div>' +
          '<div class="kop-divider-double"></div>' +
        '</div>' +
        '<div class="doc-title">' +
          '<h3>MATRIKS REKAPITULASI SEBARAN JENIS JABATAN ASN</h3>' +
          '<p>44 Perangkat Daerah & Kecamatan se-Kabupaten Toba | Filter: ' + filterLabel + ' | Posisi Per: ' + tanggalCetak + '</p>' +
        '</div>' +
        '<table class="rekap">' +
          '<thead>' +
            '<tr>' +
              '<th style="width: 25px;">No</th>' +
              '<th style="width: 320px;">Nama Perangkat Daerah / Instansi</th>' +
              '<th style="width: 80px; text-align: right; background: #7c3aed;">Struktural</th>' +
              '<th style="width: 80px; text-align: right; background: #0284c7;">Fungsional</th>' +
              '<th style="width: 80px; text-align: right; background: #d97706;">Pelaksana</th>' +
              '<th style="width: 90px; text-align: right; background: #059669;">Total ASN</th>' +
            '</tr>' +
          '</thead>' +
          '<tbody>' +
            tableRowsHtml +
            '<tr class="total-row">' +
              '<td colspan="2" style="text-align: center; font-size: 8.5pt;">TOTAL KESELURUHAN PEMKAB TOBA</td>' +
              '<td style="text-align: right; font-size: 8.5pt; color: #ddd6fe;">' + Number(t.struktural || 0).toLocaleString('id-ID') + '</td>' +
              '<td style="text-align: right; font-size: 8.5pt; color: #bae6fd;">' + Number(t.fungsional || 0).toLocaleString('id-ID') + '</td>' +
              '<td style="text-align: right; font-size: 8.5pt; color: #fde68a;">' + Number(t.pelaksana || 0).toLocaleString('id-ID') + '</td>' +
              '<td style="text-align: right; font-size: 9pt; background-color: #00ffaa; color: #0f172a;">' + Number(t.total_all || 0).toLocaleString('id-ID') + '</td>' +
            '</tr>' +
          '</tbody>' +
        '</table>' +
        '<table style="width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 7.5pt; background: #f8fafc; border: 1px solid #cbd5e1;">' +
          '<tr>' +
            '<td style="width: 50%; vertical-align: top; padding: 6px 10px; border-right: 1px solid #cbd5e1;">' +
              '<div style="font-weight: bold; color: #0284c7; margin-bottom: 4px; text-transform: uppercase;">🎓 Peta Jenjang JF Keahlian (Total: ' + (jf.total_keahlian || 0) + ')</div>' +
              '<div>• Ahli Utama: <b>' + (jf.utama || 0) + '</b> pegawai</div>' +
              '<div>• Ahli Madya: <b>' + (jf.madya || 0) + '</b> pegawai</div>' +
              '<div>• Ahli Muda: <b>' + (jf.muda || 0) + '</b> pegawai</div>' +
              '<div>• Ahli Pertama: <b>' + (jf.pertama || 0) + '</b> pegawai</div>' +
            '</td>' +
            '<td style="width: 50%; vertical-align: top; padding: 6px 10px;">' +
              '<div style="font-weight: bold; color: #d97706; margin-bottom: 4px; text-transform: uppercase;">🛠️ Peta Jenjang JF Keterampilan / Vokasi (Total: ' + (jf.total_keterampilan || 0) + ')</div>' +
              '<div>• Penyelia: <b>' + (jf.penyelia || 0) + '</b> pegawai</div>' +
              '<div>• Mahir / Pelaksana Lanjutan: <b>' + (jf.mahir || 0) + '</b> pegawai</div>' +
              '<div>• Terampil / Pelaksana: <b>' + (jf.terampil || 0) + '</b> pegawai</div>' +
              '<div>• Pemula: <b>' + (jf.pemula || 0) + '</b> pegawai</div>' +
            '</td>' +
          '</tr>' +
        '</table>' +
        '<div class="ttd-container">' +
          '<div class="ttd-box">' +
            '<div>Balige, ' + tanggalCetak + '</div>' +
            '<div style="font-weight: bold; margin-top: 2px;">KEPALA BKPSDM KABUPATEN TOBA</div>' +
            '<div class="ttd-space"></div>' +
            '<div style="font-weight: bold; text-decoration: underline;">DONALD SIMANJUNTAK, SH</div>' +
            '<div>Pembina Utama Muda (IV/c)</div>' +
            '<div>NIP. 197011252006041002</div>' +
          '</div>' +
        '</div>' +
        '<div class="footer-info">' +
          'Dicetak secara otomatis melalui Dasboard Analytics System Bidang: PPDI Kab. Toba pada ' + tanggalCetak + ' pukul ' + waktuCetak + ' WIB.' +
        '</div>' +
      '</body>' +
      '</html>';

    const blob = HtmlService.createHtmlOutput(htmlContent)
      .getAs('application/pdf')
      .setName('Rekapitulasi_Jenis_Jabatan_ASN_Toba_' + dateFileStr + '.pdf');

    return {
      success: true,
      filename: blob.getName(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    Logger.log("Error generateRekapJenisJabatanPDF: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

function generateRekapJenisJabatanExcel(filterAsn) {
  try {
    const res = getRekapJenisJabatanData(filterAsn);
    if (!res || !res.success) {
      throw new Error("Gagal mengambil data rekapitulasi jenis jabatan: " + (res ? res.error : ""));
    }

    const t = res.totals || {};
    const jf = t.jf_levels || {};
    const rows = res.data || [];
    const tanggalCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    const dateFileStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');

    let rowsXml = '';
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const displayName = normalizeOpdName(r.nama);
      rowsXml += '<tr>' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;">' + (r.no || '') + '</td>' +
        '<td style="border:0.5pt solid #cccccc;font-weight:bold;">' + displayName + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.struktural || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.fungsional || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.pelaksana || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background-color:#e6f4ea;" mso-number-format="#,##0">' + (r.jumlah || 0) + '</td>' +
      '</tr>';
    }

    const excelHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head>' +
        '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">' +
        '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>' +
        '<x:Name>Rekap Jenis Jabatan 44 OPD</x:Name>' +
        '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
        '<style>' +
          'body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; }' +
          '.header-title { font-size: 14pt; font-weight: bold; color: #1e293b; text-align: left; }' +
          '.header-sub { font-size: 10pt; color: #64748b; font-style: italic; }' +
          '.table-header th { background-color: #0f172a; color: #ffffff; font-weight: bold; border: 0.5pt solid #000000; text-align: center; }' +
          '.total-cell { background-color: #0f172a; color: #00ffaa; font-weight: bold; border: 0.5pt solid #000000; text-align: right; }' +
        '</style>' +
      '</head>' +
      '<body>' +
        '<table>' +
          '<tr><td colspan="6" class="header-title">PEMERINTAH KABUPATEN TOBA</td></tr>' +
          '<tr><td colspan="6" style="font-size:12pt;font-weight:bold;">BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA</td></tr>' +
          '<tr><td colspan="6" class="header-sub">Matriks Rekapitulasi Sebaran ASN Berdasarkan Jenis Jabatan (Struktural, Fungsional, & Pelaksana) di Organisasi Perangkat Daerah / Kecamatan</td></tr>' +
          '<tr><td colspan="6" class="header-sub">Tanggal Ekspor: ' + tanggalCetak + ' WIB</td></tr>' +
          '<tr></tr>' +
          '<tr class="table-header">' +
            '<th style="width:40px;">NO</th>' +
            '<th style="width:360px;">NAMA PERANGKAT DAERAH / INSTANSI</th>' +
            '<th style="width:110px;">STRUKTURAL</th>' +
            '<th style="width:110px;">FUNGSIONAL</th>' +
            '<th style="width:110px;">PELAKSANA</th>' +
            '<th style="width:120px;">TOTAL ASN</th>' +
          '</tr>' +
          rowsXml +
          '<tr>' +
            '<td colspan="2" style="background-color:#0f172a;color:#ffffff;font-weight:bold;text-align:center;border:0.5pt solid #000000;">TOTAL KESELURUHAN PEMKAB TOBA</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.struktural || 0) + '</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.fungsional || 0) + '</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.pelaksana || 0) + '</td>' +
            '<td class="total-cell" style="font-size:12pt;color:#00ffaa;" mso-number-format="#,##0">' + (t.total_all || 0) + '</td>' +
          '</tr>' +
          '<tr></tr>' +
          '<tr><td colspan="6" style="font-weight:bold;background-color:#e2e8f0;">DISTRIBUSI JENJANG JABATAN FUNGSIONAL (JF) PEMKAB TOBA</td></tr>' +
          '<tr><td colspan="2">JF Ahli Utama</td><td mso-number-format="#,##0">' + (jf.utama || 0) + '</td><td colspan="3" style="color:#64748b;">Kategori Keahlian</td></tr>' +
          '<tr><td colspan="2">JF Ahli Madya</td><td mso-number-format="#,##0">' + (jf.madya || 0) + '</td><td colspan="3" style="color:#64748b;">Kategori Keahlian</td></tr>' +
          '<tr><td colspan="2">JF Ahli Muda</td><td mso-number-format="#,##0">' + (jf.muda || 0) + '</td><td colspan="3" style="color:#64748b;">Kategori Keahlian</td></tr>' +
          '<tr><td colspan="2">JF Ahli Pertama</td><td mso-number-format="#,##0">' + (jf.pertama || 0) + '</td><td colspan="3" style="color:#64748b;">Kategori Keahlian</td></tr>' +
          '<tr><td colspan="2">JF Penyelia</td><td mso-number-format="#,##0">' + (jf.penyelia || 0) + '</td><td colspan="3" style="color:#64748b;">Kategori Keterampilan</td></tr>' +
          '<tr><td colspan="2">JF Mahir / Pelaksana Lanjutan</td><td mso-number-format="#,##0">' + (jf.mahir || 0) + '</td><td colspan="3" style="color:#64748b;">Kategori Keterampilan</td></tr>' +
          '<tr><td colspan="2">JF Terampil / Pelaksana</td><td mso-number-format="#,##0">' + (jf.terampil || 0) + '</td><td colspan="3" style="color:#64748b;">Kategori Keterampilan</td></tr>' +
          '<tr><td colspan="2">JF Pemula</td><td mso-number-format="#,##0">' + (jf.pemula || 0) + '</td><td colspan="3" style="color:#64748b;">Kategori Keterampilan</td></tr>' +
          '<tr><td colspan="2" style="font-weight:bold;">Total Seluruh Pejabat Fungsional</td><td style="font-weight:bold;" mso-number-format="#,##0">' + (t.fungsional || 0) + '</td><td></td></tr>' +
        '</table>' +
      '</body>' +
      '</html>';

    const blob = Utilities.newBlob(excelHtml, 'application/vnd.ms-excel', 'Rekapitulasi_Jenis_Jabatan_ASN_Toba_' + dateFileStr + '.xls');

    return {
      success: true,
      filename: blob.getName(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    Logger.log("Error generateRekapJenisJabatanExcel: " + err.toString());
    return { success: false, error: err.toString() };
  }
}