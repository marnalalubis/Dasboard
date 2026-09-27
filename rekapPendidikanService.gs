/**
 * =========================================================================
 * ACRU EMPLOYEE ANALYTICS SYSTEM - KABUPATEN TOBA
 * File: rekapPendidikanService.gs
 * Deskripsi: Standarisasi kualifikasi pendidikan (S3 s.d SD), agregasi sebaran
 * tingkat pendidikan 44 OPD, serta Generator Dokumen Cetak (PDF & Excel Landscape).
 * Pasangan Komponen Frontend: subRekapPendidikan.html
 * =========================================================================
 */

function normalizePendidikan(rawPend) {
  var p = String(rawPend || '').toUpperCase().trim();
  if (!p || p === '-' || p === '0' || p === 'NON PENDIDIKAN') return 'LAINNYA';

  if (p.indexOf('S-3') !== -1 || p.indexOf('S3') !== -1 || p.indexOf('DOKTOR') !== -1 || p.indexOf('STRATA 3') !== -1) {
    return 'S3';
  }
  if (p.indexOf('S-2') !== -1 || p.indexOf('S2') !== -1 || p.indexOf('MAGISTER') !== -1 || p.indexOf('MASTER') !== -1 || p.indexOf('PASCA') !== -1 || p.indexOf('STRATA 2') !== -1) {
    return 'S2';
  }
  if (p.indexOf('D-III') !== -1 || p.indexOf('D3') !== -1 || p.indexOf('D-3') !== -1 || p.indexOf('DIPLOMA III') !== -1 || p.indexOf('DIPLOMA 3') !== -1 || p.indexOf('SARJANA MUDA') !== -1 || p.indexOf('A.MD') !== -1 || p.indexOf('AMD') !== -1 || p.indexOf('AKADEMI') !== -1) {
    return 'D3';
  }
  if (p.indexOf('D-II') !== -1 || p.indexOf('D2') !== -1 || p.indexOf('D-2') !== -1 || p.indexOf('DIPLOMA II') !== -1 || p.indexOf('DIPLOMA 2') !== -1 || p.indexOf('D-I') !== -1 || p.indexOf('D1') !== -1 || p.indexOf('D-1') !== -1 || p.indexOf('DIPLOMA I') !== -1 || p.indexOf('DIPLOMA 1') !== -1) {
    return 'D1_D2';
  }
  if (p.indexOf('S-1') !== -1 || p.indexOf('S1') !== -1 || p.indexOf('STRATA 1') !== -1 || (p.indexOf('SARJANA') !== -1 && p.indexOf('SARJANA MUDA') === -1) || p.indexOf('D-IV') !== -1 || p.indexOf('D4') !== -1 || p.indexOf('D-4') !== -1 || p.indexOf('DIPLOMA IV') !== -1 || p.indexOf('DIPLOMA 4') !== -1 || p.indexOf('PROFESI') !== -1) {
    return 'S1';
  }
  if (p.indexOf('SMA') !== -1 || p.indexOf('SMK') !== -1 || p.indexOf('SLTA') !== -1 || p.indexOf('STM') !== -1 || p.indexOf('SMEA') !== -1 || p.indexOf('SPK') !== -1 || p.indexOf('ALIYAH') !== -1 || p.indexOf('SMU') !== -1 || p.indexOf('MENENGAH ATAS') !== -1 || p.indexOf('PGA') !== -1 || p.indexOf('KPAA') !== -1) {
    return 'SLTA';
  }
  if (p.indexOf('SMP') !== -1 || p.indexOf('SLTP') !== -1 || p.indexOf('MTS') !== -1 || p.indexOf('TSANAWIYAH') !== -1 || p.indexOf('MENENGAH PERTAMA') !== -1) {
    return 'SLTP';
  }
  if (p.indexOf('SD') !== -1 || p.indexOf('SEKOLAH DASAR') !== -1 || p.indexOf('IBTIDAIYAH') !== -1 || p.indexOf('MI') !== -1) {
    return 'SD';
  }

  return 'LAINNYA';
}

function getRekapPendidikanData(filterAsn) {
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
        s3: 0, s2: 0, s1: 0, d3: 0, d1_d2: 0,
        slta: 0, sltp: 0, sd: 0, lainnya: 0
      };
    });

    sheets.forEach(function(shInfo) {
      const sheet = ss.getSheetByName(shInfo.name);
      if (!sheet || sheet.getLastRow() <= 1) return;

      const data = sheet.getDataRange().getDisplayValues();
      if (data.length <= 1) return;

      const headers = data[0];
      const idxPend = findColumnIndex(headers, ["TINGKAT PENDIDIKAN NAMA", "PENDIDIKAN", "TINGKAT PENDIDIKAN", "PENDIDIKAN TERAKHIR"]);
      const idxUnor = findColumnIndex(headers, ["UNOR NAMA", "UNOR", "UNIT ORGANISASI", "UNIT KERJA"]);
      const idxSatker = findColumnIndex(headers, ["SATUAN KERJA KERJA NAMA", "SATKER", "SATUAN KERJA", "SATUAN KERJA NAMA"]);
      const idxSatkerInduk = findColumnIndex(headers, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK", "INDUK"]);

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.every(function(cell) { return String(cell).trim() === ''; })) continue;

        let satkerIndukVal = idxSatkerInduk !== -1 ? row[idxSatkerInduk] : "";
        let satkerVal = idxSatker !== -1 ? row[idxSatker] : "";
        let unorVal = idxUnor !== -1 ? row[idxUnor] : "";
        let pendVal = idxPend !== -1 ? row[idxPend] : "";

        let targetCat = normalizeOpdName(mapUnorToTarget(satkerIndukVal, satkerVal, unorVal));
        if (targetCat && rekapMap[targetCat]) {
          let code = normalizePendidikan(pendVal);
          switch (code) {
            case 'S3': rekapMap[targetCat].s3++; break;
            case 'S2': rekapMap[targetCat].s2++; break;
            case 'S1': rekapMap[targetCat].s1++; break;
            case 'D3': rekapMap[targetCat].d3++; break;
            case 'D1_D2': rekapMap[targetCat].d1_d2++; break;
            case 'SLTA': rekapMap[targetCat].slta++; break;
            case 'SLTP': rekapMap[targetCat].sltp++; break;
            case 'SD': rekapMap[targetCat].sd++; break;
            default: rekapMap[targetCat].lainnya++; break;
          }
        }
      }
    });

    let finalData = [];
    let counter = 1;
    let totals = {
      s3: 0, s2: 0, s1: 0, d3: 0, d1_d2: 0,
      slta: 0, sltp: 0, sd: 0, lainnya: 0,
      total_all: 0
    };

    TARGET_ORDER_PD.forEach(function(cat) {
      let cleanCat = normalizeOpdName(cat);
      let r = rekapMap[cleanCat];
      let no = cleanCat === "RSUD Porsea" ? "" : counter++;
      let jumlah = r.s3 + r.s2 + r.s1 + r.d3 + r.d1_d2 + r.slta + r.sltp + r.sd + r.lainnya;

      totals.s3 += r.s3;
      totals.s2 += r.s2;
      totals.s1 += r.s1;
      totals.d3 += r.d3;
      totals.d1_d2 += r.d1_d2;
      totals.slta += r.slta;
      totals.sltp += r.sltp;
      totals.sd += r.sd;
      totals.lainnya += r.lainnya;
      totals.total_all += jumlah;

      finalData.push({
        no: no,
        nama: cleanCat,
        s3: r.s3,
        s2: r.s2,
        s1: r.s1,
        d3: r.d3,
        d1_d2: r.d1_d2,
        slta: r.slta,
        sltp: r.sltp,
        sd: r.sd,
        lainnya: r.lainnya,
        jumlah: jumlah
      });
    });

    return {
      success: true,
      data: finalData,
      totals: totals
    };
  } catch (e) {
    Logger.log("Error getRekapPendidikanData: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function generateRekapPendidikanPDF(filterAsn) {
  try {
    const res = getRekapPendidikanData(filterAsn);
    if (!res || !res.success) {
      throw new Error("Gagal mengambil data rekapitulasi pendidikan: " + (res ? res.error : ""));
    }

    const t = res.totals || {};
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
        '<td style="text-align: center; border: 1px solid #cbd5e1; padding: 4px 2px; font-size: 7.5pt;">' + (r.no || '-') + '</td>' +
        '<td style="border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; font-weight: bold; color: #1e293b;">' + displayName + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 4px; font-size: 7.5pt;">' + (r.s3 || '-') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 4px; font-size: 7.5pt;">' + (r.s2 || '-') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 4px; font-size: 7.5pt;">' + (r.s1 || '-') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 4px; font-size: 7.5pt;">' + (r.d3 || '-') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 4px; font-size: 7.5pt;">' + (r.d1_d2 || '-') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 4px; font-size: 7.5pt;">' + (r.slta || '-') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 4px; font-size: 7.5pt;">' + (r.sltp || '-') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 4px; font-size: 7.5pt;">' + (r.sd || '-') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 4px; font-size: 7.5pt;">' + (r.lainnya || '-') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; font-weight: bold; background-color: #e2e8f0;">' + Number(r.jumlah || 0).toLocaleString('id-ID') + '</td>' +
      '</tr>';
    }

    const htmlContent = '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
        '<meta charset="UTF-8">' +
        '<style>' +
          '@page { size: A4 landscape; margin: 10mm 12mm 12mm 12mm; }' +
          'body { font-family: Arial, Helvetica, sans-serif; color: #0f172a; margin: 0; padding: 0; }' +
          '.kop-wrapper { text-align: center; margin-bottom: 6px; }' +
          '.kop-line-1 { font-size: 11pt; font-weight: bold; letter-spacing: 0.5px; text-transform: uppercase; margin: 0; }' +
          '.kop-line-2 { font-size: 13pt; font-weight: bold; letter-spacing: 0.5px; text-transform: uppercase; margin: 2px 0; }' +
          '.kop-line-3 { font-size: 8.5pt; color: #334155; margin: 0; }' +
          '.kop-divider-double { border: 0; border-top: 2px solid #000; border-bottom: 0.5px solid #000; height: 3px; margin: 6px 0 10px 0; }' +
          '.doc-title { text-align: center; margin-bottom: 10px; }' +
          '.doc-title h3 { margin: 0; font-size: 11pt; font-weight: bold; text-transform: uppercase; text-decoration: underline; }' +
          '.doc-title p { margin: 3px 0 0 0; font-size: 8pt; color: #475569; }' +
          'table.rekap { width: 100%; border-collapse: collapse; margin-bottom: 12px; }' +
          'table.rekap th { background-color: #0f172a; color: #ffffff; border: 1px solid #0f172a; padding: 5px 3px; font-size: 7.5pt; text-transform: uppercase; text-align: center; }' +
          '.total-row td { background-color: #0f172a !important; color: #ffffff !important; font-weight: bold; padding: 5px 3px; border: 1px solid #0f172a; }' +
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
          '<div class="kop-line-3">JJln.AB. Silalahi, Desa Sianipar Sihail-hail, Balige (22311), Sumatera Utara</div>' +
          '<div class="kop-divider-double"></div>' +
        '</div>' +
        '<div class="doc-title">' +
          '<h3>MATRIKS REKAPITULASI SEBARAN PENDIDIKAN PEGAWAI ASN</h3>' +
          '<p>Organisasi Perangkat Daerah & Kecamatan se-Kabupaten Toba | Filter: ' + filterLabel + ' | Posisi Per: ' + tanggalCetak + '</p>' +
        '</div>' +
        '<table class="rekap">' +
          '<thead>' +
            '<tr>' +
              '<th style="width: 25px;">No</th>' +
              '<th style="width: 260px;">Nama Perangkat Daerah / Instansi</th>' +
              '<th style="width: 38px;">S-3</th>' +
              '<th style="width: 38px;">S-2</th>' +
              '<th style="width: 60px;">S-1/D-IV</th>' +
              '<th style="width: 45px;">D-III</th>' +
              '<th style="width: 50px;">D-I/D-II</th>' +
              '<th style="width: 45px;">SLTA</th>' +
              '<th style="width: 40px;">SLTP</th>' +
              '<th style="width: 35px;">SD</th>' +
              '<th style="width: 45px;">Lainnya</th>' +
              '<th style="width: 65px;">Total</th>' +
            '</tr>' +
          '</thead>' +
          '<tbody>' +
            tableRowsHtml +
            '<tr class="total-row">' +
              '<td colspan="2" style="text-align: center; font-size: 8pt;">TOTAL KESELURUHAN PEMKAB TOBA</td>' +
              '<td style="text-align: right; font-size: 7.5pt;">' + (t.s3 || 0) + '</td>' +
              '<td style="text-align: right; font-size: 7.5pt;">' + (t.s2 || 0) + '</td>' +
              '<td style="text-align: right; font-size: 7.5pt;">' + (t.s1 || 0) + '</td>' +
              '<td style="text-align: right; font-size: 7.5pt;">' + (t.d3 || 0) + '</td>' +
              '<td style="text-align: right; font-size: 7.5pt;">' + (t.d1_d2 || 0) + '</td>' +
              '<td style="text-align: right; font-size: 7.5pt;">' + (t.slta || 0) + '</td>' +
              '<td style="text-align: right; font-size: 7.5pt;">' + (t.sltp || 0) + '</td>' +
              '<td style="text-align: right; font-size: 7.5pt;">' + (t.sd || 0) + '</td>' +
              '<td style="text-align: right; font-size: 7.5pt;">' + (t.lainnya || 0) + '</td>' +
              '<td style="text-align: right; font-size: 8.5pt; background-color: #00ffaa; color: #0f172a;">' + Number(t.total_all || 0).toLocaleString('id-ID') + '</td>' +
            '</tr>' +
          '</tbody>' +
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
      .setName('Rekapitulasi_Pendidikan_ASN_Toba_' + dateFileStr + '.pdf');

    return {
      success: true,
      filename: blob.getName(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    Logger.log("Error generateRekapPendidikanPDF: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

function generateRekapPendidikanExcel(filterAsn) {
  try {
    const res = getRekapPendidikanData(filterAsn);
    if (!res || !res.success) {
      throw new Error("Gagal mengambil data rekapitulasi pendidikan: " + (res ? res.error : ""));
    }

    const t = res.totals || {};
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
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.s3 || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.s2 || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.s1 || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.d3 || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.d1_d2 || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.slta || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.sltp || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.sd || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.lainnya || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background-color:#e6f4ea;" mso-number-format="#,##0">' + (r.jumlah || 0) + '</td>' +
      '</tr>';
    }

    const excelHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head>' +
        '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">' +
        '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>' +
        '<x:Name>Rekap Pendidikan 44 OPD</x:Name>' +
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
          '<tr><td colspan="12" class="header-title">PEMERINTAH KABUPATEN TOBA</td></tr>' +
          '<tr><td colspan="12" style="font-size:12pt;font-weight:bold;">BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA</td></tr>' +
          '<tr><td colspan="12" class="header-sub">Matriks Rekapitulasi Sebaran Pegawai ASN Berdasarkan Tingkat Pendidikan di Organisasi Perangkat Daerah / Kecamatan</td></tr>' +
          '<tr><td colspan="12" class="header-sub">Tanggal Ekspor: ' + tanggalCetak + ' WIB</td></tr>' +
          '<tr></tr>' +
          '<tr class="table-header">' +
            '<th style="width:40px;">NO</th>' +
            '<th style="width:340px;">NAMA PERANGKAT DAERAH / INSTANSI</th>' +
            '<th style="width:50px;">S-3</th>' +
            '<th style="width:50px;">S-2</th>' +
            '<th style="width:75px;">S-1 / D-IV</th>' +
            '<th style="width:60px;">D-III</th>' +
            '<th style="width:65px;">D-I / D-II</th>' +
            '<th style="width:65px;">SLTA</th>' +
            '<th style="width:55px;">SLTP</th>' +
            '<th style="width:50px;">SD</th>' +
            '<th style="width:60px;">LAINNYA</th>' +
            '<th style="width:90px;">TOTAL</th>' +
          '</tr>' +
          rowsXml +
          '<tr>' +
            '<td colspan="2" style="background-color:#0f172a;color:#ffffff;font-weight:bold;text-align:center;border:0.5pt solid #000000;">TOTAL KESELURUHAN PEMKAB TOBA</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.s3 || 0) + '</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.s2 || 0) + '</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.s1 || 0) + '</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.d3 || 0) + '</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.d1_d2 || 0) + '</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.slta || 0) + '</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.sltp || 0) + '</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.sd || 0) + '</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.lainnya || 0) + '</td>' +
            '<td class="total-cell" style="font-size:12pt;color:#00ffaa;" mso-number-format="#,##0">' + (t.total_all || 0) + '</td>' +
          '</tr>' +
        '</table>' +
      '</body>' +
      '</html>';

    const blob = Utilities.newBlob(excelHtml, 'application/vnd.ms-excel', 'Rekapitulasi_Pendidikan_ASN_Toba_' + dateFileStr + '.xls');

    return {
      success: true,
      filename: blob.getName(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    Logger.log("Error generateRekapPendidikanExcel: " + err.toString());
    return { success: false, error: err.toString() };
  }
}