/**
 * =========================================================================
 * ACRU EMPLOYEE ANALYTICS SYSTEM - KABUPATEN TOBA
 * File: csvImportService.gs
 * Deskripsi: Parser berkas CSV, analisis otomatis kecocokan kolom (Exact & Alias),
 * penulisan batch data ke sheet target, dan generator ID Log sekuensial.
 * Pasangan Komponen Frontend: subImportCSV.html
 * =========================================================================
 */

function analyzeCSVMapping(csvHeaders, targetAsn) {
  try {
    if (typeof ensureDatabaseReady === 'function') ensureDatabaseReady();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let targetSheetName = 'Data_Pegawai_PNS';
    if (targetAsn === 'PPPK') targetSheetName = 'Data_Pegawai_PPPK';
    else if (targetAsn === 'PPPK_PARUH_WAKTU') targetSheetName = 'Data_Pegawai_PPPK_Paruh_Waktu';
    else if (targetAsn === 'STRUKTUR_JABATAN') targetSheetName = 'Data_Master_Struktur_Jabatan';
    else if (targetAsn === 'KEBUTUHAN_JF') targetSheetName = 'Data_Master_Kebutuhan_JF';
    else if (targetAsn === 'KEBUTUHAN_JP') targetSheetName = 'Data_Master_Kebutuhan_JP';

    const sheetTarget = ss.getSheetByName(targetSheetName);
    if (!sheetTarget) throw new Error("Sheet target " + targetSheetName + " tidak ditemukan.");

    const targetHeaders = sheetTarget.getRange(1, 1, 1, sheetTarget.getLastColumn())
      .getDisplayValues()[0]
      .map(function(h) { return String(h).trim(); });

    const cleanTargetHeaders = targetHeaders.map(function(h) { return h.toUpperCase().replace(/[^A-Z0-9]/g, ''); });
    const cleanCsvHeaders = (csvHeaders || []).map(function(h) { return String(h).toUpperCase().replace(/[^A-Z0-9]/g, ''); });

    const mappings = [];
    let matchedCount = 0;

    for (let tIdx = 0; tIdx < targetHeaders.length; tIdx++) {
      const tName = targetHeaders[tIdx];
      const cleanTName = cleanTargetHeaders[tIdx];
      let matchedCsvIdx = -1;
      let matchType = 'UNMAPPED';

      for (let cIdx = 0; cIdx < cleanCsvHeaders.length; cIdx++) {
        if (cleanCsvHeaders[cIdx] === cleanTName && cleanTName !== '') {
          matchedCsvIdx = cIdx;
          matchType = 'EXACT';
          break;
        }
      }

      if (matchedCsvIdx === -1 && cleanTName.length >= 3) {
        for (let cIdx = 0; cIdx < cleanCsvHeaders.length; cIdx++) {
          const cleanC = cleanCsvHeaders[cIdx];
          if (cleanC.length >= 3 && (cleanTName.indexOf(cleanC) !== -1 || cleanC.indexOf(cleanTName) !== -1)) {
            matchedCsvIdx = cIdx;
            matchType = 'ALIAS';
            break;
          }
        }
      }

      if (matchedCsvIdx !== -1) matchedCount++;

      mappings.push({
        targetIndex: tIdx,
        targetName: tName,
        csvIndex: matchedCsvIdx,
        csvName: matchedCsvIdx !== -1 ? csvHeaders[matchedCsvIdx] : '',
        matchType: matchType
      });
    }

    return {
      success: true,
      targetSheetName: targetSheetName,
      targetHeaders: targetHeaders,
      csvHeaders: csvHeaders,
      mappings: mappings,
      matchedCount: matchedCount,
      totalTargetColumns: targetHeaders.length
    };
  } catch (err) {
    Logger.log("Error analyzeCSVMapping: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

function importCSVDataWithMapping(csvContent, targetAsn, customMappings) {
  try {
    if (typeof ensureDatabaseReady === 'function') ensureDatabaseReady();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let targetSheetName = 'Data_Pegawai_PNS';
    if (targetAsn === 'PPPK') targetSheetName = 'Data_Pegawai_PPPK';
    else if (targetAsn === 'PPPK_PARUH_WAKTU') targetSheetName = 'Data_Pegawai_PPPK_Paruh_Waktu';
    else if (targetAsn === 'STRUKTUR_JABATAN') targetSheetName = 'Data_Master_Struktur_Jabatan';
    else if (targetAsn === 'KEBUTUHAN_JF') targetSheetName = 'Data_Master_Kebutuhan_JF';
    else if (targetAsn === 'KEBUTUHAN_JP') targetSheetName = 'Data_Master_Kebutuhan_JP';

    const sheetTarget = ss.getSheetByName(targetSheetName);
    const sheetLog = ss.getSheetByName('Log_Import');

    if (!sheetTarget || !sheetLog) throw new Error("Sheet " + targetSheetName + " tidak ditemukan.");

    const parsedData = parseCSV(csvContent);
    if (parsedData.length === 0) throw new Error("Format berkas CSV kosong atau tidak valid.");

    const sheetHeaders = sheetTarget.getRange(1, 1, 1, sheetTarget.getLastColumn())
      .getDisplayValues()[0];

    const firstRowClean = parsedData[0].map(function(c) { return String(c).toUpperCase().replace(/[^A-Z0-9]/g, ''); });
    const targetClean = sheetHeaders.map(function(h) { return String(h).toUpperCase().replace(/[^A-Z0-9]/g, ''); });
    let isHeaderDetected = firstRowClean.some(function(c) { return c && targetClean.indexOf(c) !== -1; });

    const startRowIndex = isHeaderDetected ? 1 : 0;
    const newRows = [];

    for (let r = startRowIndex; r < parsedData.length; r++) {
      const csvRow = parsedData[r];
      if (csvRow.length === 0 || (csvRow.length === 1 && csvRow[0] === '')) continue;

      const newRow = new Array(sheetHeaders.length).fill('');
      customMappings.forEach(function(map) {
        const tIdx = parseInt(map.targetIndex, 10);
        const cIdx = parseInt(map.csvIndex, 10);
        if (!isNaN(tIdx) && tIdx >= 0 && tIdx < sheetHeaders.length && !isNaN(cIdx) && cIdx >= 0 && cIdx < csvRow.length) {
          newRow[tIdx] = csvRow[cIdx];
        }
      });
      newRows.push(newRow);
    }

    if (newRows.length > 0) {
      const lastRow = sheetTarget.getLastRow();
      if (lastRow > 1) {
        sheetTarget.getRange(2, 1, lastRow - 1, sheetTarget.getMaxColumns()).clearContent();
      }

      const currentMaxRows = sheetTarget.getMaxRows();
      const requiredRows = newRows.length + 1;

      if (currentMaxRows < requiredRows) {
        sheetTarget.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
      }

      sheetTarget.getRange(2, 1, newRows.length, sheetHeaders.length).setValues(newRows);

      const maxRowsAfterInsert = sheetTarget.getMaxRows();
      if (maxRowsAfterInsert > requiredRows) {
        sheetTarget.deleteRows(requiredRows + 1, maxRowsAfterInsert - requiredRows);
      }
    }

    const logId = generateSequentialLogId();
    const waktuImport = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    const logStatus = "SUCCESS: Berhasil mengimpor " + newRows.length + " data ke " + targetSheetName + " (Visual Mapping Verified)";

    sheetLog.appendRow([logId, waktuImport, logStatus]);
    SpreadsheetApp.flush();

    return { success: true, count: newRows.length, logId: logId };
  } catch (error) {
    Logger.log("Error importCSVDataWithMapping: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

function parseCSV(csvString) {
  if (!csvString) return [];
  const firstLine = csvString.split(/\r?\n/)[0] || '';
  const commaCount = (firstLine.match(/,/g) || []).length;
  const semicolonCount = (firstLine.match(/;/g) || []).length;
  const pipeCount = (firstLine.match(/\|/g) || []).length;

  let delimiter = ',';
  if (semicolonCount > commaCount && semicolonCount > pipeCount) delimiter = ';';
  else if (pipeCount > commaCount && pipeCount > semicolonCount) delimiter = '|';

  const result = [];
  let row = [];
  let cell = '';
  let inQuotes = false;

  for (let i = 0; i < csvString.length; i++) {
    const char = csvString[i];
    const nextChar = csvString[i + 1];

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        cell += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === delimiter && !inQuotes) {
      row.push(cell.trim());
      cell = '';
    } else if ((char === '\r' || char === '\n') && !inQuotes) {
      if (char === '\r' && nextChar === '\n') i++;
      row.push(cell.trim());
      result.push(row);
      row = [];
      cell = '';
    } else {
      cell += char;
    }
  }

  if (cell || row.length > 0) {
    row.push(cell.trim());
    result.push(row);
  }

  return result.filter(function(r) { return r.length > 0 && r.some(function(c) { return c !== ''; }); });
}

function generateSequentialLogId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Log_Import');
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd');
    const prefix = "LOG-" + todayStr + "-";

    if (!sheet || sheet.getLastRow() <= 1) return prefix + "0001";

    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    let maxNum = 0;

    for (let i = 0; i < values.length; i++) {
      const id = String(values[i][0] || '').trim();
      if (id.indexOf(prefix) === 0) {
        const numPart = parseInt(id.replace(prefix, ''), 10);
        if (!isNaN(numPart) && numPart > maxNum) maxNum = numPart;
      }
    }

    return prefix + (maxNum + 1).toString().padStart(4, '0');
  } catch (err) {
    return "LOG-" + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd') + "-0001";
  }
}