/**
 * =========================================================================
 * ACRU EMPLOYEE ANALYTICS SYSTEM - KABUPATEN TOBA
 * File: rekapGolonganService.gs
 * Deskripsi: Normalisasi pangkat/golongan PNS & PPPK, agregasi matriks sebaran,
 * serta Generator Dokumen PDF & Excel Landscape Golongan.
 * Pasangan Komponen Frontend: subRekapGolongan.html
 * =========================================================================
 */

function normalizeGolRuang(rawGol, isPns) {
  var g = String(rawGol || '').toUpperCase().trim();
  if (!g || g === '-') return isPns ? 'LAINNYA' : 'PPPK';
  
  if (g.indexOf('IV/E') !== -1 || g === 'IV E' || g === 'IV.E') return 'IV/e';
  if (g.indexOf('IV/D') !== -1 || g === 'IV D' || g === 'IV.D') return 'IV/d';
  if (g.indexOf('IV/C') !== -1 || g === 'IV C' || g === 'IV.C') return 'IV/c';
  if (g.indexOf('IV/B') !== -1 || g === 'IV B' || g === 'IV.B') return 'IV/b';
  if (g.indexOf('IV/A') !== -1 || g === 'IV A' || g === 'IV.A' || g === 'IV') return 'IV/a';

  if (g.indexOf('III/D') !== -1 || g === 'III D' || g === 'III.D') return 'III/d';
  if (g.indexOf('III/C') !== -1 || g === 'III C' || g === 'III.C') return 'III/c';
  if (g.indexOf('III/B') !== -1 || g === 'III B' || g === 'III.B') return 'III/b';
  if (g.indexOf('III/A') !== -1 || g === 'III A' || g === 'III.A' || g === 'III') return 'III/a';

  if (g.indexOf('II/D') !== -1 || g === 'II D' || g === 'II.D') return 'II/d';
  if (g.indexOf('II/C') !== -1 || g === 'II C' || g === 'II.C') return 'II/c';
  if (g.indexOf('II/B') !== -1 || g === 'II B' || g === 'II.B') return 'II/b';
  if (g.indexOf('II/A') !== -1 || g === 'II A' || g === 'II.A' || g === 'II') return 'II/a';

  if (g.indexOf('I/D') !== -1 || g === 'I D' || g === 'I.D') return 'I/d';
  if (g.indexOf('I/C') !== -1 || g === 'I C' || g === 'I.C') return 'I/c';
  if (g.indexOf('I/B') !== -1 || g === 'I B' || g === 'I.B') return 'I/b';
  if (g.indexOf('I/A') !== -1 || g === 'I A' || g === 'I.A' || g === 'I') return 'I/a';

  return isPns ? 'LAINNYA' : 'PPPK';
}

function normalizePppkGol(rawGol) {
  var g = String(rawGol || '').toUpperCase().trim();
  if (!g || g === '-' || g === 'NON GOLONGAN') return 'LAINNYA';

  var romanMatch = g.match(/\b(XVII|XVI|XV|XIV|XIII|XII|XI|X|IX|VIII|VII|VI|V|IV|III|II|I)\b/);
  if (romanMatch) return romanMatch[1];

  var numMatch = g.match(/\b(17|16|15|14|13|12|11|10|9|8|7|6|5|4|3|2|1)\b/);
  if (numMatch) {
    var mapNum = {
      '1': 'I', '2': 'II', '3': 'III', '4': 'IV', '5': 'V',
      '6': 'VI', '7': 'VII', '8': 'VIII', '9': 'IX', '10': 'X',
      '11': 'XI', '12': 'XII', '13': 'XIII', '14': 'XIV', '15': 'XV',
      '16': 'XVI', '17': 'XVII'
    };
    return mapNum[numMatch[1]] || 'LAINNYA';
  }

  return 'LAINNYA';
}

function getRekapGolonganData() {
  try {
    if (typeof ensureDatabaseReady === 'function') ensureDatabaseReady();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheetPNS = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, 'Data_Pegawai_PNS') : ss.getSheetByName('Data_Pegawai_PNS');
    const sheetPPPK = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, 'Data_Pegawai_PPPK') : ss.getSheetByName('Data_Pegawai_PPPK');
    const sheetPPPK_PW = typeof getSheetByNameRobust === 'function' ? getSheetByNameRobust(ss, 'Data_Pegawai_PPPK_Paruh_Waktu') : ss.getSheetByName('Data_Pegawai_PPPK_Paruh_Waktu');

    let rekapMap = {};
    TARGET_ORDER_PD.forEach(function(pd) {
      var norm = normalizeOpdName(pd);
      rekapMap[norm] = {
        nama: norm,
        iv_e: 0, iv_d: 0, iv_c: 0, iv_b: 0, iv_a: 0, sum_iv: 0,
        iii_d: 0, iii_c: 0, iii_b: 0, iii_a: 0, sum_iii: 0,
        ii_d: 0, ii_c: 0, ii_b: 0, ii_a: 0, sum_ii: 0,
        i_d: 0, i_c: 0, i_b: 0, i_a: 0, sum_i: 0,
        sum_pns: 0,
        pppk_i: 0, pppk_iii: 0, pppk_v: 0, pppk_vi: 0, pppk_vii: 0,
        pppk_ix: 0, pppk_x: 0, pppk_xi: 0, pppk_lainnya: 0, sum_pppk: 0
      };
    });

    if (sheetPNS && sheetPNS.getLastRow() > 1) {
      const dataPNS = sheetPNS.getDataRange().getDisplayValues();
      const headers = dataPNS[0];
      const idxGol = findColumnIndex(headers, ["GOL AKHIR NAMA", "GOLONGAN", "GOL", "GOL RUANG", "PANGKAT GOLONGAN"]);
      const idxSatker = findColumnIndex(headers, ["SATUAN KERJA KERJA NAMA", "SATKER", "SATUAN KERJA"]);
      const idxSatkerInduk = findColumnIndex(headers, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK", "INDUK"]);
      const idxUnor = findColumnIndex(headers, ["UNOR NAMA", "UNOR", "UNIT ORGANISASI", "UNIT KERJA"]);

      for (let i = 1; i < dataPNS.length; i++) {
        const row = dataPNS[i];
        if (row.every(function(cell) { return String(cell).trim() === ''; })) continue;

        let sInduk = idxSatkerInduk !== -1 ? row[idxSatkerInduk] : '';
        let sKerja = idxSatker !== -1 ? row[idxSatker] : '';
        let sUnor = idxUnor !== -1 ? row[idxUnor] : '';
        let golVal = idxGol !== -1 ? row[idxGol] : '';

        let targetPD = normalizeOpdName(mapUnorToTarget(sInduk, sKerja, sUnor));
        if (targetPD && rekapMap[targetPD]) {
          let normGol = normalizeGolRuang(golVal, true);
          let obj = rekapMap[targetPD];
          obj.sum_pns++;

          switch (normGol) {
            case 'IV/e': obj.iv_e++; obj.sum_iv++; break;
            case 'IV/d': obj.iv_d++; obj.sum_iv++; break;
            case 'IV/c': obj.iv_c++; obj.sum_iv++; break;
            case 'IV/b': obj.iv_b++; obj.sum_iv++; break;
            case 'IV/a': obj.iv_a++; obj.sum_iv++; break;

            case 'III/d': obj.iii_d++; obj.sum_iii++; break;
            case 'III/c': obj.iii_c++; obj.sum_iii++; break;
            case 'III/b': obj.iii_b++; obj.sum_iii++; break;
            case 'III/a': obj.iii_a++; obj.sum_iii++; break;

            case 'II/d': obj.ii_d++; obj.sum_ii++; break;
            case 'II/c': obj.ii_c++; obj.sum_ii++; break;
            case 'II/b': obj.ii_b++; obj.sum_ii++; break;
            case 'II/a': obj.ii_a++; obj.sum_ii++; break;

            case 'I/d': obj.i_d++; obj.sum_i++; break;
            case 'I/c': obj.i_c++; obj.sum_i++; break;
            case 'I/b': obj.i_b++; obj.sum_i++; break;
            case 'I/a': obj.i_a++; obj.sum_i++; break;
            default:
              if (normGol.indexOf('IV') === 0) obj.sum_iv++;
              else if (normGol.indexOf('III') === 0) obj.sum_iii++;
              else if (normGol.indexOf('II') === 0) obj.sum_ii++;
              else if (normGol.indexOf('I') === 0) obj.sum_i++;
              break;
          }
        }
      }
    }

    const pppkSheets = [sheetPPPK, sheetPPPK_PW];
    for (let ps = 0; ps < pppkSheets.length; ps++) {
      const shPPPK = pppkSheets[ps];
      if (shPPPK && shPPPK.getLastRow() > 1) {
        const dataPPPK = shPPPK.getDataRange().getDisplayValues();
        const headers = dataPPPK[0];
        const idxGol = findColumnIndex(headers, ["GOL AKHIR NAMA", "GOLONGAN", "GOL", "GOL RUANG", "PANGKAT GOLONGAN"]);
        const idxSatker = findColumnIndex(headers, ["SATUAN KERJA KERJA NAMA", "SATKER", "SATUAN KERJA"]);
        const idxSatkerInduk = findColumnIndex(headers, ["SATUAN KERJA INDUK NAMA", "SATKER INDUK", "INDUK"]);
        const idxUnor = findColumnIndex(headers, ["UNOR NAMA", "UNOR", "UNIT ORGANISASI", "UNIT KERJA"]);

        for (let i = 1; i < dataPPPK.length; i++) {
          const row = dataPPPK[i];
          if (row.every(function(cell) { return String(cell).trim() === ''; })) continue;

          let sInduk = idxSatkerInduk !== -1 ? row[idxSatkerInduk] : '';
          let sKerja = idxSatker !== -1 ? row[idxSatker] : '';
          let sUnor = idxUnor !== -1 ? row[idxUnor] : '';
          let golVal = idxGol !== -1 ? row[idxGol] : '';

          let targetPD = normalizeOpdName(mapUnorToTarget(sInduk, sKerja, sUnor));
          if (targetPD && rekapMap[targetPD]) {
            let normPppk = normalizePppkGol(golVal);
            let obj = rekapMap[targetPD];
            obj.sum_pppk++;

            switch (normPppk) {
              case 'I': obj.pppk_i++; break;
              case 'III': obj.pppk_iii++; break;
              case 'V': obj.pppk_v++; break;
              case 'VI': obj.pppk_vi++; break;
              case 'VII': obj.pppk_vii++; break;
              case 'IX': obj.pppk_ix++; break;
              case 'X': obj.pppk_x++; break;
              case 'XI': obj.pppk_xi++; break;
              default: obj.pppk_lainnya++; break;
            }
          }
        }
      }
    }

    let finalData = [];
    let counter = 1;
    let totals = {
      iv_e: 0, iv_d: 0, iv_c: 0, iv_b: 0, iv_a: 0, sum_iv: 0,
      iii_d: 0, iii_c: 0, iii_b: 0, iii_a: 0, sum_iii: 0,
      ii_d: 0, ii_c: 0, ii_b: 0, ii_a: 0, sum_ii: 0,
      i_d: 0, i_c: 0, i_b: 0, i_a: 0, sum_i: 0,
      sum_pns: 0,
      pppk_i: 0, pppk_iii: 0, pppk_v: 0, pppk_vi: 0, pppk_vii: 0,
      pppk_ix: 0, pppk_x: 0, pppk_xi: 0, pppk_lainnya: 0, sum_pppk: 0,
      total_all: 0
    };

    TARGET_ORDER_PD.forEach(function(cat) {
      let cleanCat = normalizeOpdName(cat);
      let r = rekapMap[cleanCat];
      let no = cleanCat === "RSUD Porsea" ? "" : counter++;
      let totalPegawai = r.sum_pns + r.sum_pppk;

      totals.iv_e += r.iv_e; totals.iv_d += r.iv_d; totals.iv_c += r.iv_c; totals.iv_b += r.iv_b; totals.iv_a += r.iv_a; totals.sum_iv += r.sum_iv;
      totals.iii_d += r.iii_d; totals.iii_c += r.iii_c; totals.iii_b += r.iii_b; totals.iii_a += r.iii_a; totals.sum_iii += r.sum_iii;
      totals.ii_d += r.ii_d; totals.ii_c += r.ii_c; totals.ii_b += r.ii_b; totals.ii_a += r.ii_a; totals.sum_ii += r.sum_ii;
      totals.i_d += r.i_d; totals.i_c += r.i_c; totals.i_b += r.i_b; totals.i_a += r.i_a; totals.sum_i += r.sum_i;
      totals.sum_pns += r.sum_pns;

      totals.pppk_i += r.pppk_i; totals.pppk_iii += r.pppk_iii; totals.pppk_v += r.pppk_v; totals.pppk_vi += r.pppk_vi;
      totals.pppk_vii += r.pppk_vii; totals.pppk_ix += r.pppk_ix; totals.pppk_x += r.pppk_x; totals.pppk_xi += r.pppk_xi;
      totals.pppk_lainnya += r.pppk_lainnya; totals.sum_pppk += r.sum_pppk;

      totals.total_all += totalPegawai;

      finalData.push({
        no: no,
        nama: cleanCat,
        iv_e: r.iv_e, iv_d: r.iv_d, iv_c: r.iv_c, iv_b: r.iv_b, iv_a: r.iv_a, sum_iv: r.sum_iv,
        iii_d: r.iii_d, iii_c: r.iii_c, iii_b: r.iii_b, iii_a: r.iii_a, sum_iii: r.sum_iii,
        ii_d: r.ii_d, ii_c: r.ii_c, ii_b: r.ii_b, ii_a: r.ii_a, sum_ii: r.sum_ii,
        i_d: r.i_d, i_c: r.i_c, i_b: r.i_b, i_a: r.i_a, sum_i: r.sum_i,
        sum_pns: r.sum_pns,
        pppk_i: r.pppk_i, pppk_iii: r.pppk_iii, pppk_v: r.pppk_v, pppk_vi: r.pppk_vi,
        pppk_vii: r.pppk_vii, pppk_ix: r.pppk_ix, pppk_x: r.pppk_x, pppk_xi: r.pppk_xi,
        pppk_lainnya: r.pppk_lainnya, sum_pppk: r.sum_pppk,
        jumlah: totalPegawai
      });
    });

    return {
      success: true,
      data: finalData,
      totals: totals
    };
  } catch (e) {
    Logger.log("Error getRekapGolonganData: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function generateRekapGolonganPDF(isDetail) {
  try {
    const res = getRekapGolonganData();
    if (!res || !res.success) {
      throw new Error("Gagal mengambil data rekapitulasi golongan: " + (res ? res.error : ""));
    }

    const t = res.totals || {};
    const rows = res.data || [];
    const tanggalCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd MMMM yyyy');
    const waktuCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'HH:mm:ss');
    const dateFileStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');

    let theadHtml = '';
    let tbodyHtml = '';
    let totalRowHtml = '';

    if (!isDetail) {
      theadHtml = '<tr>' +
        '<th rowspan="2" style="width:25px;">No</th>' +
        '<th rowspan="2" style="width:240px;">Nama Perangkat Daerah / Instansi</th>' +
        '<th colspan="4">PNS (Golongan)</th>' +
        '<th colspan="9">PPPK (Golongan)</th>' +
        '<th rowspan="2" style="width:65px;">Total</th>' +
      '</tr>' +
      '<tr>' +
        '<th style="width:38px;">IV</th><th style="width:38px;">III</th><th style="width:38px;">II</th><th style="width:38px;">I</th>' +
        '<th style="width:26px;">I</th><th style="width:26px;">III</th><th style="width:26px;">V</th><th style="width:26px;">VI</th>' +
        '<th style="width:26px;">VII</th><th style="width:26px;">IX</th><th style="width:26px;">X</th><th style="width:26px;">XI</th>' +
        '<th style="width:45px;background-color:#1e293b;">Jml</th>' +
      '</tr>';

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const displayName = normalizeOpdName(r.nama);
        const isRsud = displayName === 'RSUD Porsea';
        const bgStyle = isRsud ? 'background-color: #f0fdf4;' : (i % 2 === 1 ? 'background-color: #f8fafc;' : 'background-color: #ffffff;');

        tbodyHtml += '<tr style="' + bgStyle + '">' +
          '<td style="text-align:center;border:1px solid #cbd5e1;padding:3px;font-size:7.5pt;">' + (r.no || '-') + '</td>' +
          '<td style="border:1px solid #cbd5e1;padding:3px 6px;font-size:8pt;font-weight:bold;color:#1e293b;">' + displayName + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;">' + (r.sum_iv || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;">' + (r.sum_iii || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;">' + (r.sum_ii || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;">' + (r.sum_i || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;">' + (r.pppk_i || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;">' + (r.pppk_iii || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;">' + (r.pppk_v || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;">' + (r.pppk_vi || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;">' + (r.pppk_vii || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;">' + (r.pppk_ix || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;">' + (r.pppk_x || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;">' + (r.pppk_xi || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px;font-weight:bold;background-color:#f1f5f9;">' + (r.sum_pppk || 0) + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:3px 6px;font-weight:bold;background-color:#e2e8f0;">' + Number(r.jumlah || 0).toLocaleString('id-ID') + '</td>' +
        '</tr>';
      }

      totalRowHtml = '<tr class="total-row">' +
        '<td colspan="2" style="text-align:center;">TOTAL KESELURUHAN PEMKAB TOBA</td>' +
        '<td style="text-align:right;">' + (t.sum_iv || 0) + '</td>' +
        '<td style="text-align:right;">' + (t.sum_iii || 0) + '</td>' +
        '<td style="text-align:right;">' + (t.sum_ii || 0) + '</td>' +
        '<td style="text-align:right;">' + (t.sum_i || 0) + '</td>' +
        '<td style="text-align:right;">' + (t.pppk_i || 0) + '</td>' +
        '<td style="text-align:right;">' + (t.pppk_iii || 0) + '</td>' +
        '<td style="text-align:right;">' + (t.pppk_v || 0) + '</td>' +
        '<td style="text-align:right;">' + (t.pppk_vi || 0) + '</td>' +
        '<td style="text-align:right;">' + (t.pppk_vii || 0) + '</td>' +
        '<td style="text-align:right;">' + (t.pppk_ix || 0) + '</td>' +
        '<td style="text-align:right;">' + (t.pppk_x || 0) + '</td>' +
        '<td style="text-align:right;">' + (t.pppk_xi || 0) + '</td>' +
        '<td style="text-align:right;background-color:#334155;">' + (t.sum_pppk || 0) + '</td>' +
        '<td style="text-align:right;background-color:#00ffaa;color:#0f172a;">' + Number(t.total_all || 0).toLocaleString('id-ID') + '</td>' +
      '</tr>';
    } else {
      theadHtml = '<tr>' +
        '<th rowspan="2" style="width:20px;">No</th>' +
        '<th rowspan="2" style="width:200px;">Perangkat Daerah</th>' +
        '<th colspan="6">Golongan IV</th>' +
        '<th colspan="5">Golongan III</th>' +
        '<th colspan="5">Golongan II</th>' +
        '<th colspan="5">Golongan I</th>' +
        '<th colspan="9">PPPK</th>' +
        '<th rowspan="2" style="width:50px;">Total</th>' +
      '</tr>' +
      '<tr>' +
        '<th>e</th><th>d</th><th>c</th><th>b</th><th>a</th><th>Jml</th>' +
        '<th>d</th><th>c</th><th>b</th><th>a</th><th>Jml</th>' +
        '<th>d</th><th>c</th><th>b</th><th>a</th><th>Jml</th>' +
        '<th>d</th><th>c</th><th>b</th><th>a</th><th>Jml</th>' +
        '<th>I</th><th>III</th><th>V</th><th>VI</th><th>VII</th><th>IX</th><th>X</th><th>XI</th><th>Jml</th>' +
      '</tr>';

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const displayName = normalizeOpdName(r.nama);
        const isRsud = displayName === 'RSUD Porsea';
        const bgStyle = isRsud ? 'background-color: #f0fdf4;' : (i % 2 === 1 ? 'background-color: #f8fafc;' : 'background-color: #ffffff;');

        tbodyHtml += '<tr style="' + bgStyle + '">' +
          '<td style="text-align:center;border:1px solid #cbd5e1;padding:2px;font-size:7pt;">' + (r.no || '-') + '</td>' +
          '<td style="border:1px solid #cbd5e1;padding:2px 4px;font-size:7pt;font-weight:bold;color:#1e293b;">' + displayName + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.iv_e || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.iv_d || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.iv_c || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.iv_b || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.iv_a || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-weight:bold;font-size:6.5pt;background:#f1f5f9;">' + (r.sum_iv || 0) + '</td>' +

          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.iii_d || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.iii_c || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.iii_b || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.iii_a || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-weight:bold;font-size:6.5pt;background:#f1f5f9;">' + (r.sum_iii || 0) + '</td>' +

          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.ii_d || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.ii_c || '-') + '</td>' +
          '<td style="text-align:border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.ii_b || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.ii_a || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-weight:bold;font-size:6.5pt;background:#f1f5f9;">' + (r.sum_ii || 0) + '</td>' +

          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.i_d || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.i_c || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.i_b || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.i_a || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-weight:bold;font-size:6.5pt;background:#f1f5f9;">' + (r.sum_i || 0) + '</td>' +

          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.pppk_i || '-') + '</td>' +
          '<td style="text-align:border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.pppk_iii || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.pppk_v || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.pppk_vi || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.pppk_vii || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.pppk_ix || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.pppk_x || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-size:6.5pt;">' + (r.pppk_xi || '-') + '</td>' +
          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px;font-weight:bold;font-size:6.5pt;background:#f1f5f9;">' + (r.sum_pppk || 0) + '</td>' +

          '<td style="text-align:right;border:1px solid #cbd5e1;padding:2px 4px;font-weight:bold;font-size:7pt;background-color:#e2e8f0;">' + Number(r.jumlah || 0).toLocaleString('id-ID') + '</td>' +
        '</tr>';
      }

      totalRowHtml = '<tr class="total-row">' +
        '<td colspan="2" style="text-align:center;font-size:7pt;">TOTAL KESELURUHAN</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.iv_e || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.iv_d || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.iv_c || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.iv_b || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.iv_a || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;background:#334155;">' + (t.sum_iv || 0) + '</td>' +

        '<td style="text-align:right;font-size:6.5pt;">' + (t.iii_d || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.iii_c || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.iii_b || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.iii_a || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;background:#334155;">' + (t.sum_iii || 0) + '</td>' +

        '<td style="text-align:right;font-size:6.5pt;">' + (t.ii_d || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.ii_c || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.ii_b || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.ii_a || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;background:#334155;">' + (t.sum_ii || 0) + '</td>' +

        '<td style="text-align:right;font-size:6.5pt;">' + (t.i_d || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.i_c || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.i_b || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.i_a || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;background:#334155;">' + (t.sum_i || 0) + '</td>' +

        '<td style="text-align:right;font-size:6.5pt;">' + (t.pppk_i || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.pppk_iii || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.pppk_v || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.pppk_vi || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.pppk_vii || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.pppk_ix || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.pppk_x || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;">' + (t.pppk_xi || 0) + '</td>' +
        '<td style="text-align:right;font-size:6.5pt;background:#334155;">' + (t.sum_pppk || 0) + '</td>' +

        '<td style="text-align:right;font-size:7.5pt;background-color:#00ffaa;color:#0f172a;">' + Number(t.total_all || 0).toLocaleString('id-ID') + '</td>' +
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
          'table.rekap { width: 100%; border-collapse: collapse; margin-bottom: 10px; }' +
          'table.rekap th { background-color: #0f172a; color: #ffffff; border: 1px solid #0f172a; padding: 4px 2px; font-size: 7pt; text-transform: uppercase; text-align: center; }' +
          '.total-row td { background-color: #0f172a !important; color: #ffffff !important; font-weight: bold; padding: 4px 2px; border: 1px solid #0f172a; }' +
          '.ttd-container { width: 100%; margin-top: 8px; page-break-inside: avoid; }' +
          '.ttd-box { float: right; width: 240px; text-align: center; font-size: 8pt; line-height: 1.35; }' +
          '.ttd-space { height: 40px; }' +
          '.footer-info { font-size: 6.5pt; color: #94a3b8; font-style: italic; margin-top: 10px; clear: both; }' +
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
          '<h3>MATRIKS REKAPITULASI SEBARAN GOLONGAN PEGAWAI ASN</h3>' +
          '<p>Organisasi Perangkat Daerah & Kecamatan se-Kabupaten Toba | Format: ' + (isDetail ? 'Rincian Ruang PNS (I/a s.d IV/e) & Golongan PPPK' : 'Ringkasan Golongan PNS & PPPK') + ' | Posisi Per: ' + tanggalCetak + '</p>' +
        '</div>' +
        '<table class="rekap">' +
          '<thead>' + theadHtml + '</thead>' +
          '<tbody>' + tbodyHtml + totalRowHtml + '</tbody>' +
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
      .setName('Rekapitulasi_Golongan_ASN_Toba_' + (isDetail ? 'Detail_' : 'Summary_') + dateFileStr + '.pdf');

    return {
      success: true,
      filename: blob.getName(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    Logger.log("Error generateRekapGolonganPDF: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

function generateRekapGolonganExcel(isDetail) {
  try {
    const res = getRekapGolonganData();
    if (!res || !res.success) {
      throw new Error("Gagal mengambil data rekapitulasi golongan: " + (res ? res.error : ""));
    }

    const t = res.totals || {};
    const rows = res.data || [];
    const tanggalCetak = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    const dateFileStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');

    let tableHeaderXml = '';
    let rowsXml = '';
    let totalRowXml = '';

    if (!isDetail) {
      tableHeaderXml = '<tr class="table-header">' +
        '<th rowspan="2" style="width:40px;">NO</th>' +
        '<th rowspan="2" style="width:340px;">NAMA PERANGKAT DAERAH / INSTANSI</th>' +
        '<th colspan="4">PNS</th>' +
        '<th colspan="9">PPPK</th>' +
        '<th rowspan="2" style="width:110px;">TOTAL</th>' +
      '</tr>' +
      '<tr class="table-header">' +
        '<th style="width:65px;">GOL. IV</th>' +
        '<th style="width:65px;">GOL. III</th>' +
        '<th style="width:65px;">GOL. II</th>' +
        '<th style="width:65px;">GOL. I</th>' +
        '<th style="width:50px;">I</th>' +
        '<th style="width:50px;">III</th>' +
        '<th style="width:50px;">V</th>' +
        '<th style="width:50px;">VI</th>' +
        '<th style="width:50px;">VII</th>' +
        '<th style="width:50px;">IX</th>' +
        '<th style="width:50px;">X</th>' +
        '<th style="width:50px;">XI</th>' +
        '<th style="width:75px;background:#0284c7;">JML PPPK</th>' +
      '</tr>';

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const displayName = normalizeOpdName(r.nama);
        rowsXml += '<tr>' +
          '<td style="text-align:center;border:0.5pt solid #cccccc;">' + (r.no || '') + '</td>' +
          '<td style="border:0.5pt solid #cccccc;font-weight:bold;">' + displayName + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.sum_iv || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.sum_iii || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.sum_ii || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.sum_i || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.pppk_i || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.pppk_iii || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.pppk_v || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.pppk_vi || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.pppk_vii || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.pppk_ix || 0) + '</td>' +
          '<td style="text-align:border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.pppk_x || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;" mso-number-format="#,##0">' + (r.pppk_xi || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background:#e0f2fe;" mso-number-format="#,##0">' + (r.sum_pppk || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background-color:#e6f4ea;" mso-number-format="#,##0">' + (r.jumlah || 0) + '</td>' +
        '</tr>';
      }

      totalRowXml = '<tr>' +
        '<td colspan="2" style="background-color:#0f172a;color:#ffffff;font-weight:bold;text-align:center;border:0.5pt solid #000000;">TOTAL KESELURUHAN PEMKAB TOBA</td>' +
        '<td class="total-cell" mso-number-format="#,##0">' + (t.sum_iv || 0) + '</td>' +
        '<td class="total-cell" mso-number-format="#,##0">' + (t.sum_iii || 0) + '</td>' +
        '<td class="total-cell" mso-number-format="#,##0">' + (t.sum_ii || 0) + '</td>' +
        '<td class="total-cell" mso-number-format="#,##0">' + (t.sum_i || 0) + '</td>' +
        '<td class="total-cell" mso-number-format="#,##0">' + (t.pppk_i || 0) + '</td>' +
        '<td class="total-cell" mso-number-format="#,##0">' + (t.pppk_iii || 0) + '</td>' +
        '<td class="total-cell" mso-number-format="#,##0">' + (t.pppk_v || 0) + '</td>' +
        '<td class="total-cell" mso-number-format="#,##0">' + (t.pppk_vi || 0) + '</td>' +
        '<td class="total-cell" mso-number-format="#,##0">' + (t.pppk_vii || 0) + '</td>' +
        '<td class="total-cell" mso-number-format="#,##0">' + (t.pppk_ix || 0) + '</td>' +
        '<td class="total-cell" mso-number-format="#,##0">' + (t.pppk_x || 0) + '</td>' +
        '<td class="total-cell" mso-number-format="#,##0">' + (t.pppk_xi || 0) + '</td>' +
        '<td class="total-cell" style="background-color:#0284c7;" mso-number-format="#,##0">' + (t.sum_pppk || 0) + '</td>' +
        '<td class="total-cell" style="font-size:12pt;color:#00ffaa;" mso-number-format="#,##0">' + (t.total_all || 0) + '</td>' +
      '</tr>';
    } else {
      tableHeaderXml = '<tr class="table-header">' +
        '<th rowspan="2" style="width:40px;">NO</th>' +
        '<th rowspan="2" style="width:340px;">PERANGKAT DAERAH / INSTANSI</th>' +
        '<th colspan="6">GOLONGAN IV</th>' +
        '<th colspan="5">GOLONGAN III</th>' +
        '<th colspan="5">GOLONGAN II</th>' +
        '<th colspan="5">GOLONGAN I</th>' +
        '<th colspan="9">PPPK</th>' +
        '<th rowspan="2" style="width:110px;">TOTAL</th>' +
      '</tr>' +
      '<tr class="table-header">' +
        '<th>e</th><th>d</th><th>c</th><th>b</th><th>a</th><th>Jml IV</th>' +
        '<th>d</th><th>c</th><th>b</th><th>a</th><th>Jml III</th>' +
        '<th>d</th><th>c</th><th>b</th><th>a</th><th>Jml II</th>' +
        '<th>d</th><th>c</th><th>b</th><th>a</th><th>Jml I</th>' +
        '<th>I</th><th>III</th><th>V</th><th>VI</th><th>VII</th><th>IX</th><th>X</th><th>XI</th><th>Jml PPPK</th>' +
      '</tr>';

      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const displayName = normalizeOpdName(r.nama);
        rowsXml += '<tr>' +
          '<td style="text-align:center;border:0.5pt solid #cccccc;">' + (r.no || '') + '</td>' +
          '<td style="border:0.5pt solid #cccccc;font-weight:bold;">' + displayName + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.iv_e || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.iv_d || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.iv_c || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.iv_b || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.iv_a || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background-color:#f1f5f9;">' + (r.sum_iv || 0) + '</td>' +

          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.iii_d || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.iii_c || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.iii_b || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.iii_a || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background-color:#f1f5f9;">' + (r.sum_iii || 0) + '</td>' +

          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.ii_d || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.ii_c || 0) + '</td>' +
          '<td style="text-align:border:0.5pt solid #cccccc;">' + (r.ii_b || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.ii_a || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background-color:#f1f5f9;">' + (r.sum_ii || 0) + '</td>' +

          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.i_d || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.i_c || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.i_b || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.i_a || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background-color:#f1f5f9;">' + (r.sum_i || 0) + '</td>' +

          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.pppk_i || 0) + '</td>' +
          '<td style="text-align:border:0.5pt solid #cccccc;">' + (r.pppk_iii || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.pppk_v || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.pppk_vi || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.pppk_vii || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.pppk_ix || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.pppk_x || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;">' + (r.pppk_xi || 0) + '</td>' +
          '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background-color:#e0f2fe;">' + (r.sum_pppk || 0) + '</td>' +

          '<td style="text-align:right;border:0.5pt solid #cccccc;font-weight:bold;background-color:#e6f4ea;" mso-number-format="#,##0">' + (r.jumlah || 0) + '</td>' +
        '</tr>';
      }

      totalRowXml = '<tr>' +
        '<td colspan="2" style="background-color:#0f172a;color:#ffffff;font-weight:bold;text-align:center;border:0.5pt solid #000000;">TOTAL KESELURUHAN</td>' +
        '<td class="total-cell">' + (t.iv_e || 0) + '</td>' +
        '<td class="total-cell">' + (t.iv_d || 0) + '</td>' +
        '<td class="total-cell">' + (t.iv_c || 0) + '</td>' +
        '<td class="total-cell">' + (t.iv_b || 0) + '</td>' +
        '<td class="total-cell">' + (t.iv_a || 0) + '</td>' +
        '<td class="total-cell" style="background-color:#0284c7;">' + (t.sum_iv || 0) + '</td>' +

        '<td class="total-cell">' + (t.iii_d || 0) + '</td>' +
        '<td class="total-cell">' + (t.iii_c || 0) + '</td>' +
        '<td class="total-cell">' + (t.iii_b || 0) + '</td>' +
        '<td class="total-cell">' + (t.iii_a || 0) + '</td>' +
        '<td class="total-cell" style="background-color:#0284c7;">' + (t.sum_iii || 0) + '</td>' +

        '<td class="total-cell">' + (t.ii_d || 0) + '</td>' +
        '<td class="total-cell">' + (t.ii_c || 0) + '</td>' +
        '<td class="total-cell">' + (t.ii_b || 0) + '</td>' +
        '<td class="total-cell">' + (t.ii_a || 0) + '</td>' +
        '<td class="total-cell" style="background-color:#0284c7;">' + (t.sum_ii || 0) + '</td>' +

        '<td class="total-cell">' + (t.i_d || 0) + '</td>' +
        '<td class="total-cell">' + (t.i_c || 0) + '</td>' +
        '<td class="total-cell">' + (t.i_b || 0) + '</td>' +
        '<td class="total-cell">' + (t.i_a || 0) + '</td>' +
        '<td class="total-cell" style="background-color:#0284c7;">' + (t.sum_i || 0) + '</td>' +

        '<td class="total-cell">' + (t.pppk_i || 0) + '</td>' +
        '<td class="total-cell">' + (t.pppk_iii || 0) + '</td>' +
        '<td class="total-cell">' + (t.pppk_v || 0) + '</td>' +
        '<td class="total-cell">' + (t.pppk_vi || 0) + '</td>' +
        '<td class="total-cell">' + (t.pppk_vii || 0) + '</td>' +
        '<td class="total-cell">' + (t.pppk_ix || 0) + '</td>' +
        '<td class="total-cell">' + (t.pppk_x || 0) + '</td>' +
        '<td class="total-cell">' + (t.pppk_xi || 0) + '</td>' +
        '<td class="total-cell" style="background-color:#0284c7;">' + (t.sum_pppk || 0) + '</td>' +

        '<td class="total-cell" style="font-size:12pt;color:#00ffaa;">' + (t.total_all || 0) + '</td>' +
      '</tr>';
    }

    const excelHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
      '<head>' +
        '<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">' +
        '<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>' +
        '<x:Name>Rekap Golongan 44 OPD</x:Name>' +
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
          '<tr><td colspan="' + (isDetail ? 29 : 16) + '" class="header-title">PEMERINTAH KABUPATEN TOBA</td></tr>' +
          '<tr><td colspan="' + (isDetail ? 29 : 16) + '" style="font-size:12pt;font-weight:bold;">BADAN KEPEGAWAIAN DAN PENGEMBANGAN SUMBER DAYA MANUSIA</td></tr>' +
          '<tr><td colspan="' + (isDetail ? 29 : 16) + '" class="header-sub">Matriks Rekapitulasi Sebaran ASN (PNS & PPPK) Berdasarkan Golongan di Organisasi Perangkat Daerah / Kecamatan</td></tr>' +
          '<tr><td colspan="' + (isDetail ? 29 : 16) + '" class="header-sub">Format: ' + (isDetail ? 'Rincian Ruang PNS & Golongan PPPK' : 'Ringkasan Golongan PNS & PPPK') + ' | Tanggal Ekspor: ' + tanggalCetak + ' WIB</td></tr>' +
          '<tr></tr>' +
          tableHeaderXml +
          rowsXml +
          totalRowXml +
        '</table>' +
      '</body>' +
      '</html>';

    const blob = Utilities.newBlob(excelHtml, 'application/vnd.ms-excel', 'Rekapitulasi_Golongan_ASN_Toba_' + (isDetail ? 'Detail_' : 'Summary_') + dateFileStr + '.xls');

    return {
      success: true,
      filename: blob.getName(),
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (err) {
    Logger.log("Error generateRekapGolonganExcel: " + err.toString());
    return { success: false, error: err.toString() };
  }
}