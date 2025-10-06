const SHEET_NAME = 'Customer Data';

const COLUMN_KEYS = [
  'timestamp',
  'customerName',
  'licensePlate',
  'policyNumber',
  'registrationDate',
  'actIssuedDate',
  'actExpiryDate',
  'taxRenewalDate',
  'taxExpiryDate',
  'voluntaryIssuedDate',
  'voluntaryExpiryDate',
  'phone',
  'email',
  'notes',
  'status'
];

const DATE_KEYS = new Set([
  'registrationDate',
  'actIssuedDate',
  'actExpiryDate',
  'taxRenewalDate',
  'taxExpiryDate',
  'voluntaryIssuedDate',
  'voluntaryExpiryDate'
]);

function cleanText(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function parseDateInput(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  const text = String(value).trim();
  if (!text) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const parts = text.split('-');
    const year = Number(parts[0]);
    const month = Number(parts[1]) - 1;
    const day = Number(parts[2]);
    const date = new Date(year, month, day);
    return isNaN(date.getTime()) ? '' : date;
  }

  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? '' : parsed;
}

function normaliseRecordForSheet(source) {
  const record = {};
  const input = source || {};

  COLUMN_KEYS.forEach(function(key) {
    if (key === 'timestamp') {
      const coerced = parseDateInput(input.timestamp);
      record.timestamp = coerced || new Date();
      return;
    }

    if (DATE_KEYS.has(key)) {
      record[key] = parseDateInput(input[key]);
      return;
    }

    if (key === 'phone') {
      const phone = cleanText(input.phone);
      record.phone = phone ? (phone.charAt(0) === "'" ? phone : "'" + phone) : '';
      return;
    }

    record[key] = cleanText(input[key]);
  });

  return record;
}

function toRowArray(source) {
  const record = normaliseRecordForSheet(source);
  return COLUMN_KEYS.map(function(key) {
    const value = record[key];
    return value === undefined || value === null ? '' : value;
  });
}

function doGet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ records: [], error: 'Sheet not found: ' + SHEET_NAME }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const rows = sheet.getDataRange().getValues();
  if (!rows || rows.length === 0) {
    return ContentService
      .createTextOutput(JSON.stringify({ records: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const header = rows[0];
  const data = rows.slice(1);
  const records = data.map(function(row, index) {
    const record = {};
    header.forEach(function(key, idx) {
      record[key] = row[idx];
    });
    record.rowNumber = index + 2; // include header row
    return record;
  });

  return ContentService
    .createTextOutput(JSON.stringify({ records: records }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return respond(false, 'Missing POST body');
  }

  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond(false, 'Invalid JSON: ' + err);
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) return respond(false, 'Sheet not found: ' + SHEET_NAME);

  if (payload.action === 'update') {
    const row = Number(payload.rowNumber);
    if (!row || row < 2) return respond(false, 'Invalid row number');

    const record = payload.record || {};
    if (!record.timestamp) {
      const existingTimestamp = sheet.getRange(row, 1).getValue();
      if (existingTimestamp) record.timestamp = existingTimestamp;
    }

    const rowValues = toRowArray(record);
    sheet.getRange(row, 1, 1, COLUMN_KEYS.length).setValues([rowValues]);
    return respond(true);
  }

  if (payload.action === 'delete') {
    const rows = (payload.rows || []).map(Number).filter(function(n) { return n >= 2; });
    rows.sort(function(a, b) { return b - a; }).forEach(function(r) { sheet.deleteRow(r); });
    return respond(true);
  }

  const rowValues = toRowArray(payload || {});
  sheet.appendRow(rowValues);
  return respond(true);
}

function respond(ok, error) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: ok, error: error }))
    .setMimeType(ContentService.MimeType.JSON);
}
