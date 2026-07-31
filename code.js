// ============================================================
//  KONFIGURASI
// ============================================================
const SPREADSHEET_ID = '1rPABX-4HbBvDTxNFtGRnigEwdmcICXqKdiuVZBA7tMk';
const SHEET_NAME = 'DATA';

// ============================================================
//  ENTRY POINT
// ============================================================
function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet utama tidak ditemukan');

    // Gabungkan parameter query + body
    let params = { ...e.parameter };
    let action = params.action;
    let payload = null;

    if (e.postData?.contents) {
      payload = JSON.parse(e.postData.contents);
      Object.assign(params, payload);
      action = payload.action || action;
    }

    switch (action) {
      case 'read':     return readRecords(sheet, params);
      case 'create':   return createRecord(sheet, payload);
      case 'update':   return updateRecord(sheet, payload);
      case 'delete':   return deleteRecord(sheet, params);
      case 'dropdown': return getDropdownData(ss, params.field);
      default:         return jsonResponse({ error: 'Aksi tidak valid. Gunakan: read, create, update, delete, dropdown' });
    }
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  READ – dengan pagination
// ============================================================
function readRecords(sheet, params) {
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0] || [];
  const totalRows = data.length - 1;

  if (params.row) {
    const row = Number(params.row);
    if (row < 2 || row > data.length) {
      return jsonResponse({ error: 'Baris tidak valid. Data dimulai dari baris 2.' });
    }
    const obj = {};
    headers.forEach((h, i) => obj[h] = data[row - 1][i]);
    return jsonResponse({ data: obj });
  }

  const page = Math.max(1, Number(params.page) || 1);
  const size = Math.max(1, Number(params.pageSize) || 20);
  const start = (page - 1) * size + 1;
  const end = Math.min(start + size - 1, data.length - 1);
  const records = [];

  for (let i = start; i <= end; i++) {
    const row = { row: i + 1 };
    headers.forEach((h, j) => row[h] = data[i][j]);
    records.push(row);
  }

  const totalPages = totalRows > 0 ? Math.ceil(totalRows / size) : 0;
  return jsonResponse({
    data: records,
    page,
    pageSize: size,
    totalRecords: totalRows,
    totalPages
  });
}

// ============================================================
//  CREATE
// ============================================================
function createRecord(sheet, data) {
  if (!data) return jsonResponse({ error: 'Data kosong' });

  const headers = sheet.getDataRange().getValues()[0];
  if (!headers || headers.length === 0) {
    return jsonResponse({ error: 'Header tidak ditemukan di sheet.' });
  }

  const firstKey = headers[0];
  const value = data[firstKey]?.toString().trim();
  if (!value) {
    return jsonResponse({ error: `Kolom pertama '${firstKey}' wajib diisi dan tidak boleh kosong.` });
  }

  if (isDuplicate(sheet, firstKey, value)) {
    return jsonResponse({ error: `Nilai '${value}' sudah ada di kolom '${firstKey}'. Duplikat tidak diperbolehkan.` });
  }

  const newRow = headers.map(h => data[h] || '');
  sheet.appendRow(newRow);
  const lastRow = sheet.getLastRow();

  return jsonResponse({
    success: true,
    message: 'Data berhasil ditambahkan',
    row: lastRow
  });
}

// ============================================================
//  UPDATE
// ============================================================
function updateRecord(sheet, data) {
  if (!data) return jsonResponse({ error: 'Data kosong' });
  if (!data.row) return jsonResponse({ error: 'Parameter row wajib untuk update.' });

  const row = Number(data.row);
  const lastRow = sheet.getLastRow();
  if (row < 2 || row > lastRow) {
    return jsonResponse({ error: 'Baris tidak valid. Data dimulai dari baris 2.' });
  }

  const headers = sheet.getDataRange().getValues()[0];
  const firstKey = headers[0];
  const value = data[firstKey]?.toString().trim();
  if (!value) {
    return jsonResponse({ error: `Kolom pertama '${firstKey}' wajib diisi dan tidak boleh kosong.` });
  }

  if (isDuplicate(sheet, firstKey, value, row)) {
    return jsonResponse({ error: `Nilai '${value}' sudah ada di baris lain. Duplikat tidak diperbolehkan.` });
  }

  headers.forEach((h, i) => {
    if (data[h] !== undefined && h !== 'row') {
      sheet.getRange(row, i + 1).setValue(data[h]);
    }
  });

  return jsonResponse({
    success: true,
    message: `Baris ${row} berhasil diperbarui`,
    row
  });
}

// ============================================================
//  DELETE
// ============================================================
function deleteRecord(sheet, params) {
  if (!params || !params.row) {
    return jsonResponse({ error: 'Parameter row wajib untuk delete.' });
  }

  const row = Number(params.row);
  const lastRow = sheet.getLastRow();
  if (row < 2 || row > lastRow) {
    return jsonResponse({ error: 'Baris tidak valid. Data dimulai dari baris 2.' });
  }

  sheet.deleteRow(row);
  return jsonResponse({
    success: true,
    message: `Baris ${row} berhasil dihapus`,
    row
  });
}

// ============================================================
//  DROPDOWN – ambil data dari sheet lain
//  Parameter: field (nama sheet)
//  Sheet harus memiliki kolom 'id' dan 'uraian' (case sensitive)
// ============================================================
function getDropdownData(ss, fieldName) {
  if (!fieldName) {
    return jsonResponse({ error: 'Parameter field (nama sheet) wajib.' });
  }

  const dropdownSheet = ss.getSheetByName(fieldName);
  if (!dropdownSheet) {
    return jsonResponse({ error: `Sheet dengan nama '${fieldName}' tidak ditemukan.` });
  }

  const data = dropdownSheet.getDataRange().getValues();
  if (data.length < 2) {
    return jsonResponse({ data: [] });
  }

  const headers = data[0];
  const idIndex = headers.indexOf('id');
  const uraianIndex = headers.indexOf('uraian');

  if (idIndex === -1 || uraianIndex === -1) {
    return jsonResponse({ error: `Sheet '${fieldName}' harus memiliki kolom 'id' dan 'uraian'.` });
  }

  const options = [];
  for (let i = 1; i < data.length; i++) {
    options.push({
      id: data[i][idIndex] !== undefined ? data[i][idIndex].toString() : '',
      uraian: data[i][uraianIndex] !== undefined ? data[i][uraianIndex].toString() : ''
    });
  }

  return jsonResponse({ data: options });
}

// ============================================================
//  HELPER – Cek duplikat
// ============================================================
function isDuplicate(sheet, columnName, value, excludeRow = null) {
  const headers = sheet.getDataRange().getValues()[0];
  const colIndex = headers.indexOf(columnName);
  if (colIndex === -1) return false;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const range = sheet.getRange(2, colIndex + 1, lastRow - 1, 1);
  const data = range.getValues();

  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2;
    if (excludeRow && rowNum === excludeRow) continue;
    if (data[i][0].toString().trim() === value) {
      return true;
    }
  }
  return false;
}
