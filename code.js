// ============================================================
//  KONFIGURASI
// ============================================================
const SPREADSHEET_ID = '1rPABX-4HbBvDTxNFtGRnigEwdmcICXqKdiuVZBA7tMk';
const SHEET_NAME = 'DATA';

// Daftar indeks kolom (0-based) yang menjadi UNIQUE KEY
// Contoh: kolom 1, 4, 7 → indeks [0, 3, 6]
const UNIQUE_COLUMNS = [];   // sesuaikan dengan kebutuhan Anda

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
//  READ – dengan pagination, pencarian, dan row fisik
// ============================================================
function readRecords(sheet, params) {
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0] || [];
  // Buat array objek { fisikRow, cells } untuk semua baris data (mulai baris 2)
  let rows = data.slice(1).map((row, idx) => ({
    fisikRow: idx + 2,   // baris fisik di sheet (baris 1 = header)
    cells: row
  }));

  // --- PENCARIAN (SEARCH) ---
  const search = params.search?.toString().trim().toLowerCase();
  if (search) {
    rows = rows.filter(item => {
      return item.cells.some(cell => cell.toString().toLowerCase().includes(search));
    });
  }

  const totalRows = rows.length;

  // Jika parameter row diberikan, ambil satu baris spesifik (berdasarkan fisikRow)
  if (params.row) {
    const targetRow = Number(params.row);
    const found = rows.find(item => item.fisikRow === targetRow);
    if (!found) {
      return jsonResponse({ error: 'Baris tidak ditemukan.' });
    }
    const obj = { row: found.fisikRow };
    headers.forEach((h, i) => obj[h] = found.cells[i]);
    return jsonResponse({ data: obj });
  }

  // Pagination
  const page = Math.max(1, Number(params.page) || 1);
  const size = Math.max(1, Number(params.pageSize) || 20);
  const start = (page - 1) * size;
  const end = Math.min(start + size, rows.length);
  const pageRows = rows.slice(start, end);

  const records = pageRows.map(item => {
    const obj = { row: item.fisikRow };   // ← nomor baris fisik
    headers.forEach((h, i) => obj[h] = item.cells[i]);
    return obj;
  });

  const totalPages = totalRows > 0 ? Math.ceil(totalRows / size) : 0;
  return jsonResponse({
    data: records,
    page,
    pageSize: size,
    totalRecords: totalRows,
    totalPages,
    search: search || null
  });
}

// ============================================================
//  CREATE – dengan auto-fill timestamp untuk kolom unik
// ============================================================
function createRecord(sheet, data) {
  if (!data) return jsonResponse({ error: 'Data kosong' });

  const headers = sheet.getDataRange().getValues()[0];
  if (!headers || headers.length === 0) {
    return jsonResponse({ error: 'Header tidak ditemukan di sheet.' });
  }

  // Isi otomatis kolom unik jika kosong
  for (let idx of UNIQUE_COLUMNS) {
    const colName = headers[idx];
    const val = data[colName]?.toString().trim();
    if (!val) {
      data[colName] = generateTimestamp();
    }
  }

  // Validasi ulang (seharusnya sudah terisi)
  for (let idx of UNIQUE_COLUMNS) {
    const colName = headers[idx];
    const val = data[colName]?.toString().trim();
    if (!val) {
      return jsonResponse({ error: `Kolom '${colName}' (unik) gagal diisi otomatis.` });
    }
  }

  if (isDuplicate(sheet, data)) {
    const fields = UNIQUE_COLUMNS.map(idx => headers[idx]).join(', ');
    return jsonResponse({ error: `Duplikat terdeteksi pada kombinasi kolom: ${fields}` });
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
//  UPDATE – dengan auto-fill timestamp untuk kolom unik
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

  // Isi otomatis kolom unik jika kosong dan disertakan
  for (let idx of UNIQUE_COLUMNS) {
    const colName = headers[idx];
    if (data.hasOwnProperty(colName)) {
      const val = data[colName]?.toString().trim();
      if (!val) {
        data[colName] = generateTimestamp();
      }
    }
  }

  // Baca data lama untuk kolom yang tidak disertakan
  const oldRow = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const fullData = { ...data };
  for (let idx of UNIQUE_COLUMNS) {
    const colName = headers[idx];
    if (!fullData.hasOwnProperty(colName) || fullData[colName] === undefined) {
      fullData[colName] = oldRow[idx];
    }
  }

  if (isDuplicate(sheet, fullData, row)) {
    const fields = UNIQUE_COLUMNS.map(idx => headers[idx]).join(', ');
    return jsonResponse({ error: `Duplikat terdeteksi pada kombinasi kolom: ${fields}` });
  }

  // Update nilai
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
//  HELPER – generate timestamp format YYYYMMDDHHMMSSMS
// ============================================================
function generateTimestamp() {
  const now = new Date();
  const pad = (n, len = 2) => String(n).padStart(len, '0');
  const yyyy = now.getFullYear();
  const mm   = pad(now.getMonth() + 1);
  const dd   = pad(now.getDate());
  const hh   = pad(now.getHours());
  const min  = pad(now.getMinutes());
  const ss   = pad(now.getSeconds());
  const ms   = pad(now.getMilliseconds(), 3);
  return `${yyyy}${mm}${dd}${hh}${min}${ss}${ms}`;
}

// ============================================================
//  HELPER – Cek duplikat berdasarkan UNIQUE_COLUMNS
// ============================================================
function isDuplicate(sheet, rowData, excludeRow = null) {
  const headers = sheet.getDataRange().getValues()[0];
  if (!headers || headers.length === 0) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  if (UNIQUE_COLUMNS.length === 0) return false;

  const uniqueValues = UNIQUE_COLUMNS.map(idx => {
    const colName = headers[idx];
    return rowData[colName]?.toString().trim() || '';
  });

  const range = sheet.getRange(2, 1, lastRow - 1, headers.length);
  const rows = range.getValues();

  for (let i = 0; i < rows.length; i++) {
    const rowNum = i + 2;
    if (excludeRow && rowNum === excludeRow) continue;

    const rowUniqueValues = UNIQUE_COLUMNS.map(idx => rows[i][idx]?.toString().trim() || '');
    let match = true;
    for (let j = 0; j < uniqueValues.length; j++) {
      if (uniqueValues[j] !== rowUniqueValues[j]) {
        match = false;
        break;
      }
    }
    if (match) return true;
  }
  return false;
}
