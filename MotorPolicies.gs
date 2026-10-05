const MOTOR_HEADERS = [
  'Motor ID', 'Customer ID', 'Company Name', 'Registration No.', 'Vehicle Registration',
  'Vehicle Type', 'Make / Model', 'Year', 'Engine / Tonnage', 'Coverage Type',
  'Current Insurer', 'Policy No.', 'Policy Start Date', 'Policy Expiry Date',
  'Premium (MVR)', 'Policy Status', 'Quotation Status', 'UW Status', 'Documents Status',
  'Next Action', 'Next Follow-up', 'Remarks', 'Updated At'
];

const MOTOR_OPTIONS = {
  vehicleType: ['Motorcycle', 'Car', 'Pickup', 'Van', 'Truck', 'Bus', 'Heavy Vehicle', 'Other'],
  coverageType: ['Third Party', 'Comprehensive', 'Third Party Fire & Theft', 'Other']
};

function setupMotorSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, MOTOR_SHEET_NAME, MOTOR_HEADERS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, MOTOR_HEADERS.length).setFontWeight('bold').setWrap(true);
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  setDropdown_(sheet, 6, MOTOR_OPTIONS.vehicleType, rows);
  setDropdown_(sheet, 10, MOTOR_OPTIONS.coverageType, rows);
  setDropdown_(sheet, 11, OPTIONS.insurers, rows);
  setDropdown_(sheet, 16, OPTIONS.policyStatus, rows);
  setDropdown_(sheet, 17, OPTIONS.quotationStatus, rows);
  setDropdown_(sheet, 18, OPTIONS.uwStatus, rows);
  setDropdown_(sheet, 19, OPTIONS.documentsStatus, rows);
  setDropdown_(sheet, 20, OPTIONS.actions, rows);
  sheet.getRange(2, 13, rows, 2).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(2, 15, rows, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, 21, rows, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(2, 23, rows, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  [150, 180, 220, 130, 150, 130, 180, 80, 120, 160, 130, 140, 110, 110, 120, 140, 140, 140, 140, 160, 110, 280, 145]
    .forEach(function(width, i) { sheet.setColumnWidth(i + 1, width); });
  return sheet;
}

function getMotorSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MOTOR_SHEET_NAME);
  if (!sheet) sheet = setupMotorSheet_(ss);
  return sheet;
}

function getMotorOptions() {
  return {
    vehicleType: MOTOR_OPTIONS.vehicleType,
    coverageType: MOTOR_OPTIONS.coverageType,
    insurers: OPTIONS.insurers,
    policyStatus: OPTIONS.policyStatus,
    quotationStatus: OPTIONS.quotationStatus,
    uwStatus: OPTIONS.uwStatus,
    documentsStatus: OPTIONS.documentsStatus,
    actions: OPTIONS.actions
  };
}

function getMotorPolicies(customerId) {
  const sheet = getMotorSheet_();
  if (sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, MOTOR_HEADERS.length).getValues();
  const filterId = String(customerId || '').trim();
  const records = [];
  values.forEach(function(row, index) {
    if (!String(row[0] || '').trim()) return;
    if (filterId && String(row[1] || '') !== filterId) return;
    const record = {};
    MOTOR_HEADERS.forEach(function(header, i) {
      const isDate = [12, 13, 20, 22].indexOf(i) !== -1;
      if (isDate && row[i] instanceof Date) {
        record[header] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), i === 22 ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd');
      } else record[header] = row[i] === null || row[i] === undefined ? '' : row[i];
    });
    record._row = index + 2;
    records.push(record);
  });
  records.sort(function(a, b) {
    const da = parseDate_(a['Policy Expiry Date']);
    const db = parseDate_(b['Policy Expiry Date']);
    if (!da && !db) return 0;
    if (!da) return 1;
    if (!db) return -1;
    return da - db;
  });
  return records;
}

function getMotorPolicy(rowNumber) {
  const sheet = getMotorSheet_();
  const row = Number(rowNumber);
  if (!row || row < 2 || row > sheet.getLastRow()) throw new Error('Invalid motor policy record.');
  const values = sheet.getRange(row, 1, 1, MOTOR_HEADERS.length).getValues()[0];
  const record = {};
  MOTOR_HEADERS.forEach(function(header, i) {
    const isDate = [12, 13, 20, 22].indexOf(i) !== -1;
    if (isDate && values[i] instanceof Date) {
      record[header] = Utilities.formatDate(values[i], Session.getScriptTimeZone(), i === 22 ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd');
    } else record[header] = values[i] || '';
  });
  record._row = row;
  return record;
}

function saveMotorPolicy(data) {
  data = data || {};
  const sheet = getMotorSheet_();
  const row = Number(data._row) || sheet.getLastRow() + 1;
  const customer = getCustomerById_(data['Customer ID']);
  if (!customer) throw new Error('Choose a valid CRM customer before saving a motor policy.');

  const existing = row <= sheet.getLastRow() ? sheet.getRange(row, 1, 1, MOTOR_HEADERS.length).getValues()[0] : new Array(MOTOR_HEADERS.length).fill('');
  const values = MOTOR_HEADERS.map(function(header, i) {
    if (header === 'Motor ID') return String(data[header] || existing[i] || Utilities.getUuid());
    if (header === 'Company Name') return customer['Company Name'];
    if (header === 'Registration No.') return customer['Registration No.'];
    if (header === 'Updated At') return new Date();
    let value = Object.prototype.hasOwnProperty.call(data, header) ? data[header] : existing[i];
    if (['Policy Start Date', 'Policy Expiry Date', 'Next Follow-up'].indexOf(header) !== -1) value = parseDate_(value) || '';
    if (header === 'Premium (MVR)' && value !== '') value = Number(value) || 0;
    return value === null || value === undefined ? '' : value;
  });

  sheet.getRange(row, 1, 1, MOTOR_HEADERS.length).setValues([values]);
  sheet.getRange(row, 13, 1, 2).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(row, 21).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(row, 23).setNumberFormat('dd/mm/yyyy hh:mm');
  return getMotorPolicy(row);
}

function getMotorRenewals_(daysAhead) {
  const today = getToday_();
  return getMotorPolicies().filter(function(record) {
    const expiry = parseDate_(record['Policy Expiry Date']);
    if (!expiry) return false;
    const days = differenceInDays_(today, expiry);
    return days <= Number(daysAhead || 30);
  });
}
