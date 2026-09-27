/**
 * =========================================================================
 * ACRU EMPLOYEE ANALYTICS SYSTEM - KABUPATEN TOBA
 * File: rekapStatusService.gs
 * Deskripsi: Agregasi status kepegawaian (PNS, PPPK, PPPK PW), Audit UNOR Kosong,
 * serta Generator Export Resmi (PDF & Excel Landscape).
 * Pasangan Komponen Frontend: subRekapStatus.html
 * =========================================================================
 */

function getRekapitulasiData() {
  try {
    if (typeof ensureDatabaseReady === 'function') ensureDatabaseReady();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const sheets = [
      { name: 'Data_Pegawai_PNS', key: 'pns', label: 'PNS' },
      { name: 'Data_Pegawai_PPPK', key: 'pppk', label: 'PPPK' },
      { name: 'Data_Pegawai_PPPK_Paruh_Waktu', key: 'p3k_pw', label: 'PPPK Paruh Waktu' }
    ];

    let rekapMap = {};
    TARGET_ORDER_PD.forEach(function(pd) {
      var norm = normalizeOpdName(pd);
      rekapMap[norm] = { nama: norm, pns: 0, pppk: 0, p3k_pw: 0 };
    });

    let unmappedCounts = { pns: 0, pppk: 0, p3k_pw: 0 };
    let unmappedList = [];

    sheets.forEach(function(shInfo) {
      const sheet = ss.getSheetByName(shInfo.name);
      if (!sheet || sheet.getLastRow() <= 1) return;

      const data = sheet.getDataRange().getDisplayValues();
      if (data.length <= 1) return;

      const headers = data[0];
      const idxUnor = findColumnIndex(headers, ["UNOR NAMA", "UNOR", "UNIT ORGANISASI", "UNIT KERJA"]);
      const idxSatker = findColumnIndex(headers, ["SATUAN KERJA KERJA NAMA", "SATKER", "SATUAN KERJA", "SATUAN KERJA NAMA"]);
      const idxSatkerInduk = findColumnIndex(headers, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK", "INDUK"]);
      const idxNip = findColumnIndex(headers, ["NIP BARU", "NIP", "NIP_BARU", "NOMOR INDUK"]);
      const idxNama = findColumnIndex(headers, ["NAMA", "NAMA LENGKAP", "NAMA PEGAWAI"]);
      const idxGelarD = findColumnIndex(headers, ["GELAR DEPAN", "GLR_DEPAN"]);
      const idxGelarB = findColumnIndex(headers, ["GELAR BELAKANG", "GLR_BELAKANG"]);
      const idxJab = findColumnIndex(headers, ["JABATAN NAMA", "JABATAN", "NAMA JABATAN"]);
      const idxGol = findColumnIndex(headers, ["GOL AKHIR NAMA", "GOLONGAN", "GOL", "PANGKAT/GOLONGAN", "GOL RUANG"]);

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.every(function(cell) { return String(cell).trim() === ''; })) continue;

        let satkerIndukVal = idxSatkerInduk !== -1 ? String(row[idxSatkerInduk] || '').trim() : "";
        let satkerVal = idxSatker !== -1 ? String(row[idxSatker] || '').trim() : "";
        let unorVal = idxUnor !== -1 ? String(row[idxUnor] || '').trim() : "";
        let jabVal = idxJab !== -1 ? String(row[idxJab] || '').trim() : "";

        let targetCat = normalizeOpdName(mapUnorToTarget(satkerIndukVal, satkerVal, unorVal, jabVal));

        if (targetCat && targetCat !== "Lainnya" && rekapMap[targetCat]) {
          rekapMap[targetCat][shInfo.key]++;
        } else {
          unmappedCounts[shInfo.key]++;

          let rawNama = idxNama !== -1 ? String(row[idxNama] || '').trim() : '';
          let gd = idxGelarD !== -1 && row[idxGelarD] ? String(row[idxGelarD]).trim() + ' ' : '';
          let gb = idxGelarB !== -1 && row[idxGelarB] ? ', ' + String(row[idxGelarB]).trim() : '';
          let fullNama = gd + rawNama + gb;

          let reason = 'UNOR NAMA Kosong';
          if (unorVal && unorVal !== '-') {
            reason = 'UNOR Tidak Dikenali di 44 OPD';
          } else if (!satkerVal && !satkerIndukVal) {
            reason = 'Satker & UNOR Belum Diisi';
          }

          unmappedList.push({
            sheet: shInfo.name,
            row: i + 1,
            nip: idxNip !== -1 ? String(row[idxNip] || '-').trim() : '-',
            nama: fullNama || 'Tanpa Nama',
            statusAsn: shInfo.label,
            golongan: idxGol !== -1 ? String(row[idxGol] || '-').trim() : '-',
            jabatan: idxJab !== -1 ? String(row[idxJab] || '-').trim() : '-',
            satkerInduk: satkerIndukVal || '-',
            satker: satkerVal || '-',
            unor: unorVal || '(KOSONG)',
            alasan: reason
          });
        }
      }
    });

    let finalData = [];
    let counter = 1;
    let totalPns = 0, totalPppk = 0, totalP3kpw = 0;

    TARGET_ORDER_PD.forEach(function(cat) {
      let cleanCat = normalizeOpdName(cat);
      let r = rekapMap[cleanCat] || { nama: cleanCat, pns: 0, pppk: 0, p3k_pw: 0 };
      let no = cleanCat === "RSUD Porsea" ? "" : counter++;
      let jumlah = r.pns + r.pppk + r.p3k_pw;

      totalPns += r.pns;
      totalPppk += r.pppk;
      totalP3kpw += r.p3k_pw;

      finalData.push({
        no: no,
        nama: cleanCat,
        pns: r.pns,
        pppk: r.pppk,
        p3k_pw: r.p3k_pw,
        jumlah: jumlah,
        isUnmapped: false
      });
    });

    let totalUnmapped = unmappedCounts.pns + unmappedCounts.pppk + unmappedCounts.p3k_pw;
    if (totalUnmapped > 0) {
      finalData.push({
        no: "⚠️",
        nama: "Data Belum Terpetakan / UNOR Kosong",
        pns: unmappedCounts.pns,
        pppk: unmappedCounts.pppk,
        p3k_pw: unmappedCounts.p3k_pw,
        jumlah: totalUnmapped,
        isUnmapped: true
      });

      totalPns += unmappedCounts.pns;
      totalPppk += unmappedCounts.pppk;
      totalP3kpw += unmappedCounts.p3k_pw;
    }

    return {
      success: true,
      data: finalData,
      totals: {
        pns: totalPns,
        pppk: totalPppk,
        p3k_pw: totalP3kpw,
        total_all: totalPns + totalPppk + totalP3kpw,
        unmapped_total: totalUnmapped,
        unmapped_pns: unmappedCounts.pns,
        unmapped_pppk: unmappedCounts.pppk,
        unmapped_p3k_pw: unmappedCounts.p3k_pw
      },
      unmappedData: unmappedList
    };
  } catch (e) {
    Logger.log("Error getRekapitulasiData: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function getDetailPegawaiUnorKosong() {
  try {
    const res = getRekapitulasiData();
    if (!res || !res.success) {
      throw new Error(res ? res.error : "Gagal mengambil rincian UNOR kosong");
    }
    return {
      success: true,
      unmappedData: res.unmappedData || [],
      totals: res.totals || {}
    };
  } catch (err) {
    Logger.log("Error getDetailPegawaiUnorKosong: " + err.toString());
    return { success: false, error: err.toString(), unmappedData: [] };
  }
}

function generateRekapPDF() {
  try {
    const rekapRes = getRekapitulasiData();
    if (!rekapRes || !rekapRes.success) {
      throw new Error("Gagal mengambil data rekapitulasi: " + (rekapRes ? rekapRes.error : ""));
    }

    const t = rekapRes.totals || {};
    const rows = rekapRes.data || [];
    const tanggalCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');
    const waktuCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'HH:mm:ss');
    const dateFileStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');

    let tableRowsHtml = '';
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const displayName = normalizeOpdName(r.nama);
      const isRsud = displayName === 'RSUD Porsea';
      const isUnmapped = !!r.isUnmapped;
      
      let bgStyle = isRsud ? 'background-color: #f0fdf4;' : (i % 2 === 1 ? 'background-color: #f8fafc;' : 'background-color: #ffffff;');
      let fontColor = '#1e293b';

      if (isUnmapped) {
        bgStyle = 'background-color: #fff1f2; font-weight: bold;';
        fontColor = '#be123c';
      }

      tableRowsHtml += '<tr style="' + bgStyle + '">' +
        '<td style="text-align: center; border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; color:' + fontColor + ';">' + (r.no || '-') + '</td>' +
        '<td style="border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; font-weight: bold; color:' + fontColor + ';">' + displayName + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; color:' + fontColor + ';">' + Number(r.pns || 0).toLocaleString('id-ID') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; color:' + fontColor + ';">' + Number(r.pppk || 0).toLocaleString('id-ID') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; color:' + fontColor + ';">' + Number(r.p3k_pw || 0).toLocaleString('id-ID') + '</td>' +
        '<td style="text-align: right; border: 1px solid #cbd5e1; padding: 4px 6px; font-size: 8pt; font-weight: bold; background-color: #e2e8f0; color:' + fontColor + ';">' + Number(r.jumlah || 0).toLocaleString('id-ID') + '</td>' +
      '</tr>';
    }

    const htmlContent = '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
        '<meta charset="UTF-8">' +
        '<style>' +
          '@page { size: A4 portrait; margin: 12mm 15mm 15mm 15mm; }' +
          'body { font-family: Arial, Helvetica, sans-serif; color: #0f172a; margin: 0; padding: 0; }' +
          '.kop-wrapper { text-align: center; margin-bottom: 8px; position: relative; }' +
          '.kop-line-1 { font-size: 11pt; font-weight: bold; letter-spacing: 0.5px; text-transform: uppercase; margin: 0; }' +
          '.kop-line-2 { font-size: 13pt; font-weight: bold; letter-spacing: 0.5px; text-transform: uppercase; margin: 2px 0; }' +
          '.kop-line-3 { font-size: 9pt; color: #334155; margin: 0; }' +
          '.kop-divider-double { border: 0; border-top: 2px solid #000; border-bottom: 0.5px solid #000; height: 3px; margin: 8px 0 12px 0; }' +
          '.doc-title { text-align: center; margin-bottom: 12px; }' +
          '.doc-title h3 { margin: 0; font-size: 11pt; font-weight: bold; text-transform: uppercase; text-decoration: underline; }' +
          '.doc-title p { margin: 3px 0 0 0; font-size: 8.5pt; color: #475569; }' +
          'table.rekap { width: 100%; border-collapse: collapse; margin-bottom: 14px; }' +
          'table.rekap th { background-color: #0f172a; color: #ffffff; border: 1px solid #0f172a; padding: 6px 5px; font-size: 8pt; text-transform: uppercase; }' +
          '.total-row td { background-color: #0f172a !important; color: #ffffff !important; font-weight: bold; font-size: 8.5pt; padding: 6px 5px; border: 1px solid #0f172a; }' +
          '.ttd-container { width: 100%; margin-top: 15px; page-break-inside: avoid; }' +
          '.ttd-box { float: right; width: 240px; text-align: center; font-size: 8.5pt; line-height: 1.4; }' +
          '.ttd-space { height: 50px; }' +
          '.footer-info { font-size: 7pt; color: #94a3b8; font-style: italic; margin-top: 20px; clear: both; }' +
        '</style>' +
      '</head>' +
      '<body>' +
        '<div class="kop-wrapper">' +
          '<div class="kop-line-1">PEMERINTAH KABUPATEN TOBA</div>' +
          '<div class="kop-line-2">BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA</div>' +
          '<div class="kop-line-3">Jln.AB. Silalahi, Desa Sianipar Sihail-hail, Balige (22311), Sumatera Utara </div>' +
          '<div class="kop-divider-double"></div>' +
        '</div>' +
        '<div class="doc-title">' +
          '<h3>MATRIKS REKAPITULASI SEBARAN PEGAWAI ASN</h3>' +
          '<p>Organisasi Perangkat Daerah & Kecamatan se-Kabupaten Toba | Posisi Per: ' + tanggalCetak + '</p>' +
        '</div>' +
        '<table class="rekap">' +
          '<thead>' +
            '<tr>' +
              '<th style="width: 25px;">No</th>' +
              '<th>Nama Perangkat Daerah / Instansi</th>' +
              '<th style="width: 60px; text-align: right;">PNS</th>' +
              '<th style="width: 60px; text-align: right;">PPPK Penuh</th>' +
              '<th style="width: 75px; text-align: right;">PPPK Paruh Waktu</th>' +
              '<th style="width: 75px; text-align: right;">Total ASN</th>' +
            '</tr>' +
          '</thead>' +
          '<tbody>' +
            tableRowsHtml +
            '<tr class="total-row">' +
              '<td colspan="2" style="text-align: center;">TOTAL KESELURUHAN PEMKAB TOBA</td>' +
              '<td style="text-align: right;">' + Number(t.pns || 0).toLocaleString('id-ID') + '</td>' +
              '<td style="text-align: right;">' + Number(t.pppk || 0).toLocaleString('id-ID') + '</td>' +
              '<td style="text-align: right;">' + Number(t.p3k_pw || 0).toLocaleString('id-ID') + '</td>' +
              '<td style="text-align: right;">' + Number(t.total_all || 0).toLocaleString('id-ID') + '</td>' +
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
      .setName('Rekapitulasi_ASN_Pemkab_Toba_' + dateFileStr + '.pdf');

    return {
      success: true,
      filename: blob.getName(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    Logger.log("Error generateRekapPDF: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

function generateRekapExcel() {
  try {
    const rekapRes = getRekapitulasiData();
    if (!rekapRes || !rekapRes.success) {
      throw new Error("Gagal mengambil data rekapitulasi: " + (rekapRes ? rekapRes.error : ""));
    }

    const t = rekapRes.totals || {};
    const rows = rekapRes.data || [];
    const tanggalCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    const dateFileStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');

    let rowsXml = '';
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const displayName = normalizeOpdName(r.nama);
      const isUnmapped = !!r.isUnmapped;
      const rowStyle = isUnmapped ? 'background-color:#ffe4e6;color:#be123c;' : '';

      rowsXml += '<tr style="' + rowStyle + '">' +
        '<td style="text-align:center;border:0.5pt solid #cccccc;">' + (r.no || '') + '</td>' +
        '<td style="border:0.5pt solid #cccccc;font-weight:bold;">' + displayName + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.pns || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.pppk || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.p3k_pw || 0) + '</td>' +
        '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background-color:#e6f4ea;" mso-number-format="#,##0">' + (r.jumlah || 0) + '</td>' +
      '</tr>';
    }

    const excelHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head>' +
        '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">' +
        '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>' +
        '<x:Name>Rekapitulasi 44 OPD</x:Name>' +
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
          '<tr><td colspan="6" class="header-sub">Matriks Rekapitulasi Sebaran ASN (PNS & PPPK) di Organisasi Perangkat Daerah / Kecamatan</td></tr>' +
          '<tr><td colspan="6" class="header-sub">Tanggal Ekspor: ' + tanggalCetak + ' WIB</td></tr>' +
          '<tr></tr>' +
          '<tr class="table-header">' +
            '<th style="width:50px;">NO</th>' +
            '<th style="width:380px;">NAMA PERANGKAT DAERAH / INSTANSI</th>' +
            '<th style="width:100px;">PNS</th>' +
            '<th style="width:110px;">PPPK PENUH</th>' +
            '<th style="width:140px;">PPPK PARUH WAKTU</th>' +
            '<th style="width:120px;">TOTAL PEGAWAI</th>' +
          '</tr>' +
          rowsXml +
          '<tr>' +
            '<td colspan="2" style="background-color:#0f172a;color:#ffffff;font-weight:bold;text-align:center;border:0.5pt solid #000000;">TOTAL KESELURUHAN PEMKAB TOBA</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.pns || 0) + '</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.pppk || 0) + '</td>' +
            '<td class="total-cell" mso-number-format="#,##0">' + (t.p3k_pw || 0) + '</td>' +
            '<td class="total-cell" style="font-size:12pt;color:#00ffaa;" mso-number-format="#,##0">' + (t.total_all || 0) + '</td>' +
          '</tr>' +
        '</table>' +
      '</body>' +
      '</html>';

    const blob = Utilities.newBlob(excelHtml, 'application/vnd.ms-excel', 'Rekapitulasi_Pegawai_Pemkab_Toba_' + dateFileStr + '.xls');

    return {
      success: true,
      filename: blob.getName(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    Logger.log("Error generateRekapExcel: " + err.toString());
    return { success: false, error: err.toString() };
  }
}