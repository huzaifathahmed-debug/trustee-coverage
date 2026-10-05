/**
 * Trustee Coverage — personal insurance sales CRM for Amãna Takaful.
 * Google Apps Script backend.
 *
 * Existing CRM columns A:X are intentionally preserved for compatibility.
 * New system fields are appended in Y:AB.
 */

const APP_NAME = 'Trustee Coverage';
const CRM_SHEET_NAME = 'CRM';
const MOTOR_SHEET_NAME = 'MOTOR POLICIES';
const UW_SHEET_NAME = 'UNDERWRITING';
const CACHE_SECONDS = 30;

const CRM_HEADERS = [
  'Company Name', 'Business Name', 'Registration No.', 'Contact Person',
  'Contact Number', 'Current Insurer', 'Lead Status', 'Products',
  'Last Contact', 'Next Follow-up', 'Action Required', 'Quotation Status',
  'Quotation Date', 'Policy Status', 'Policy Start Date', 'Policy Expiry Date',
  'UW Status', 'Documents Status', 'Priority', 'UW Remarks',
  'Remarks / Notes', 'Days Since Contact', 'Follow-up Status', 'Renewal Status',
  'Customer ID', 'Updated At', 'Amana Policy Check', 'UW Last Updated'
];

const OPTIONS = {
  insurers: ['Amana', 'Allied', 'Sollarelle', 'AIA', 'Other', 'Unknown', 'No Existing Policy'],
  leadStatus: [
    'New', 'Contacted', 'No Answer', 'Interested', 'Meeting Scheduled',
    'Meeting Done', 'Information Requested', 'Information Received',
    'Quotation Pending', 'Quotation Sent', 'Follow-up Required', 'Negotiation',
    'Won / Policy Issued', 'Lost', 'Not Interested', 'Not Required', 'Renewal Later'
  ],
  products: [
    'Fire', 'Burglary', 'Motor', 'Marine Hull', 'Marine Cargo', 'Expat Medical',
    'Group Medical', 'SME Medical', 'CAR', 'Machinery', 'Performance Guarantee',
    'Bid Security', 'Public Liability', 'Property', 'Travel', 'Other'
  ],
  actions: [
    'No Action', 'Call Customer', 'WhatsApp Customer', 'Arrange Meeting',
    'Follow Up', 'Follow Up Quotation', 'Follow Up Documents', 'Request Information',
    'Request Policy Certificate', 'Request Asset Values', 'Send Quotation',
    'Send Revised Quotation', 'Check with UW', 'Awaiting Underwriting',
    'Awaiting Customer', 'Prepare Proposal', 'Request Payment', 'Issue Policy',
    'Check Renewal', 'Other'
  ],
  quotationStatus: [
    'Not Required', 'Not Started', 'Pending Information', 'Pending UW', 'Preparing',
    'Sent', 'Revised', 'Accepted', 'Rejected', 'Expired'
  ],
  policyStatus: [
    'No Policy', 'Existing Policy', 'Renewal Due', 'Quotation Pending',
    'Quotation Sent', 'Accepted', 'Payment Pending', 'Policy Issued', 'Lost'
  ],
  uwStatus: [
    'Not Required', 'Not Sent to UW', 'Pending Information', 'Sent to UW',
    'Under Review', 'UW Questions', 'Terms Received', 'Quote Received',
    'Referred', 'Declined', 'Completed'
  ],
  documentsStatus: [
    'Not Required', 'Not Requested', 'Requested', 'Partially Received',
    'Received', 'Incomplete', 'Sent to UW', 'Completed'
  ],
  priority: ['High', 'Medium', 'Low'],
  amanaPolicyCheck: ['Unknown', 'No Policy With Amana', 'Has Policy With Amana', 'Needs Verification']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Run once after attaching this script to the user's CRM spreadsheet. */
function setupTrusteeCoverage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const crm = getOrCreateSheet_(ss, CRM_SHEET_NAME, CRM_HEADERS);
  setupCrmSheet_(crm);
  setupMotorSheet_(ss);
  ensureCustomerIds_(crm);
  SpreadsheetApp.flush();
  clearCrmCache_();
  return {
    ok: true,
    spreadsheetId: ss.getId(),
    crmSheet: CRM_SHEET_NAME,
    motorSheet: MOTOR_SHEET_NAME,
    message: 'Trustee Coverage CRM setup completed.'
  };
}

/** Compatibility alias for the previous version. */
function setupCRMInterface() {
  return setupTrusteeCoverage();
}

function setupCrmSheet_(sheet) {
  ensureHeaders_(sheet, CRM_HEADERS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, CRM_HEADERS.length)
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('center');

  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);
  setDropdown_(sheet, 6, OPTIONS.insurers, maxRows);
  setDropdown_(sheet, 7, OPTIONS.leadStatus, maxRows);
  setDropdown_(sheet, 8, OPTIONS.products, maxRows);
  setDropdown_(sheet, 11, OPTIONS.actions, maxRows);
  setDropdown_(sheet, 12, OPTIONS.quotationStatus, maxRows);
  setDropdown_(sheet, 14, OPTIONS.policyStatus, maxRows);
  setDropdown_(sheet, 17, OPTIONS.uwStatus, maxRows);
  setDropdown_(sheet, 18, OPTIONS.documentsStatus, maxRows);
  setDropdown_(sheet, 19, OPTIONS.priority, maxRows);
  setDropdown_(sheet, 27, OPTIONS.amanaPolicyCheck, maxRows);

  sheet.getRange(2, 9, maxRows, 2).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(2, 13, maxRows, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(2, 15, maxRows, 2).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(2, 26, maxRows, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  sheet.getRange(2, 28, maxRows, 1).setNumberFormat('dd/mm/yyyy hh:mm');

  setupFormulaRows_(sheet);
  applyCrmConditionalFormatting_(sheet);
  setCrmColumnWidths_(sheet);

  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), CRM_HEADERS.length).createFilter();
  }
}

function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  ensureHeaders_(sheet, headers);
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  const existing = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  const output = headers.map(function(header, i) {
    // Preserve known headers only when they match the expected schema; otherwise repair them.
    return existing[i] === header ? existing[i] : header;
  });
  sheet.getRange(1, 1, 1, headers.length).setValues([output]);
}

function setDropdown_(sheet, column, values, rowCount) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, column, rowCount || 999, 1).setDataValidation(rule);
}

function setupFormulaRows_(sheet) {
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 22, rows, 1)
    .setFormulaR1C1('=IF(RC[-13]="","",TODAY()-RC[-13])');
  sheet.getRange(2, 23, rows, 1)
    .setFormulaR1C1('=IF(OR(RC[-16]="Won / Policy Issued",RC[-16]="Lost",RC[-16]="Not Interested",RC[-16]="Not Required"),"Closed",IF(RC[-13]="","No Date",IF(RC[-13]<TODAY(),"OVERDUE",IF(RC[-13]=TODAY(),"TODAY",IF(RC[-13]<=TODAY()+3,"DUE SOON","UPCOMING")))))');
  sheet.getRange(2, 24, rows, 1)
    .setFormulaR1C1('=IF(RC[-8]="","",IF(RC[-8]<TODAY(),"EXPIRED",IF(RC[-8]<=TODAY()+30,"RENEWAL WITHIN 30 DAYS",IF(RC[-8]<=TODAY()+60,"RENEWAL WITHIN 60 DAYS","NOT DUE"))))');
}

function setupNewRowFormulas_(sheet, row) {
  sheet.getRange(row, 22).setFormula('=IF(I' + row + '="","",TODAY()-I' + row + ')');
  sheet.getRange(row, 23).setFormula('=IF(OR(G' + row + '="Won / Policy Issued",G' + row + '="Lost",G' + row + '="Not Interested",G' + row + '="Not Required"),"Closed",IF(J' + row + '="","No Date",IF(J' + row + '<TODAY(),"OVERDUE",IF(J' + row + '=TODAY(),"TODAY",IF(J' + row + '<=TODAY()+3,"DUE SOON","UPCOMING")))))');
  sheet.getRange(row, 24).setFormula('=IF(P' + row + '="","",IF(P' + row + '<TODAY(),"EXPIRED",IF(P' + row + '<=TODAY()+30,"RENEWAL WITHIN 30 DAYS",IF(P' + row + '<=TODAY()+60,"RENEWAL WITHIN 60 DAYS","NOT DUE"))))');
}

function applyCrmConditionalFormatting_(sheet) {
  const range = sheet.getRange('A2:AB' + sheet.getMaxRows());
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$S2="High"').setBackground('#fee2e2').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$W2="OVERDUE"').setBackground('#fff1f2').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$W2="TODAY"').setBackground('#fef9c3').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$N2="Policy Issued"').setBackground('#ecfdf5').setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}

function setCrmColumnWidths_(sheet) {
  const widths = {
    1: 220, 2: 180, 3: 130, 4: 180, 5: 140, 6: 130, 7: 150, 8: 160,
    9: 105, 10: 105, 11: 170, 12: 135, 13: 105, 14: 135, 15: 105, 16: 105,
    17: 135, 18: 140, 19: 90, 20: 260, 21: 320, 22: 100, 23: 120, 24: 170,
    25: 180, 26: 145, 27: 170, 28: 145
  };
  Object.keys(widths).forEach(function(col) { sheet.setColumnWidth(Number(col), widths[col]); });
}

function getCrmSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CRM_SHEET_NAME);
  if (!sheet) throw new Error('CRM sheet was not found. Run setupTrusteeCoverage() first.');
  return sheet;
}

function getOptions() {
  return OPTIONS;
}

function getRecords_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('crm_records_v2');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const sheet = getCrmSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, CRM_HEADERS.length).getValues();
  const records = [];
  values.forEach(function(row, index) {
    if (!String(row[0] || '').trim()) return;
    const record = {};
    CRM_HEADERS.forEach(function(header, columnIndex) {
      record[header] = formatValueForClient_(row[columnIndex], columnIndex);
    });
    record._row = index + 2;
    records.push(record);
  });

  try { cache.put('crm_records_v2', JSON.stringify(records), CACHE_SECONDS); } catch (e) {}
  return records;
}

function clearCrmCache_() {
  CacheService.getScriptCache().remove('crm_records_v2');
}

function formatValueForClient_(value, columnIndex) {
  if (value === null || value === undefined) return '';
  const dateColumns = [8, 9, 12, 14, 15, 25, 27];
  if (dateColumns.indexOf(columnIndex) !== -1 && value instanceof Date) {
    const pattern = (columnIndex === 25 || columnIndex === 27) ? "yyyy-MM-dd'T'HH:mm:ss" : 'yyyy-MM-dd';
    return Utilities.formatDate(value, Session.getScriptTimeZone(), pattern);
  }
  return value;
}

function getDashboardData() {
  const records = getRecords_();
  const today = getToday_();
  const data = {
    total: records.length, newLeads: 0, interested: 0, quotations: 0,
    policiesIssued: 0, overdue: 0, dueToday: 0, renewals30: 0,
    uwPending: 0, motorRenewals30: 0, followUps: [], renewals: []
  };

  records.forEach(function(record) {
    const lead = record['Lead Status'];
    const quote = record['Quotation Status'];
    const policy = record['Policy Status'];
    const uw = record['UW Status'];
    if (lead === 'New') data.newLeads++;
    if (['Interested', 'Meeting Scheduled', 'Meeting Done', 'Negotiation'].indexOf(lead) !== -1) data.interested++;
    if (['Sent', 'Revised', 'Accepted'].indexOf(quote) !== -1) data.quotations++;
    if (policy === 'Policy Issued') data.policiesIssued++;
    if (['Sent to UW', 'Under Review', 'UW Questions', 'Pending Information'].indexOf(uw) !== -1) data.uwPending++;

    const follow = parseDate_(record['Next Follow-up']);
    if (follow && !isClosed_(record)) {
      if (follow < today) data.overdue++;
      else if (sameDay_(follow, today)) data.dueToday++;
    }

    const expiry = parseDate_(record['Policy Expiry Date']);
    if (expiry) {
      const days = differenceInDays_(today, expiry);
      if (days >= 0 && days <= 30) data.renewals30++;
    }
  });

  data.followUps = getFollowUps_('due').slice(0, 8);
  data.renewals = getRenewals_().slice(0, 8);
  try { data.motorRenewals30 = getMotorRenewals_(30).length; } catch (e) {}
  return data;
}

function searchCustomers(query) {
  const records = getRecords_();
  const searchText = String(query || '').toLowerCase().trim();
  if (!searchText) return records.slice().reverse().slice(0, 150);

  return records.filter(function(record) {
    return [
      record['Company Name'], record['Business Name'], record['Registration No.'],
      record['Contact Person'], record['Contact Number'], record['Products'],
      record['Policy Status'], record['Customer ID']
    ].some(function(value) { return String(value || '').toLowerCase().indexOf(searchText) !== -1; });
  }).slice(0, 200);
}

function getCustomer(rowNumber) {
  const sheet = getCrmSheet_();
  const row = Number(rowNumber);
  if (!row || row < 2 || row > sheet.getLastRow()) throw new Error('Invalid customer record.');
  const values = sheet.getRange(row, 1, 1, CRM_HEADERS.length).getValues()[0];
  const record = {};
  CRM_HEADERS.forEach(function(header, index) {
    record[header] = formatValueForClient_(values[index], index);
  });
  record._row = row;
  return record;
}

function getCustomerById_(customerId) {
  const id = String(customerId || '').trim();
  if (!id) return null;
  const records = getRecords_();
  for (let i = 0; i < records.length; i++) {
    if (records[i]['Customer ID'] === id) return records[i];
  }
  return null;
}

function addLead(data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const sheet = getCrmSheet_();
    const row = findFirstEmptyCustomerRow_(sheet);
    data = data || {};
    data['Customer ID'] = data['Customer ID'] || Utilities.getUuid();
    writeCustomerData_(sheet, row, data, true);
    setupNewRowFormulas_(sheet, row);
    clearCrmCache_();
    return getCustomer(row);
  } finally { lock.releaseLock(); }
}

function saveCustomer(data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const row = Number(data && data._row);
    if (!row || row < 2) throw new Error('Invalid customer record.');
    const sheet = getCrmSheet_();
    writeCustomerData_(sheet, row, data, false);
    setupNewRowFormulas_(sheet, row);
    clearCrmCache_();
    return getCustomer(row);
  } finally { lock.releaseLock(); }
}

function writeCustomerData_(sheet, row, data, isNew) {
  data = data || {};
  const existing = sheet.getRange(row, 1, 1, CRM_HEADERS.length).getValues()[0];
  const values = existing.slice();
  // Editable customer fields A:U.
  for (let index = 0; index < 21; index++) {
    const header = CRM_HEADERS[index];
    if (Object.prototype.hasOwnProperty.call(data, header)) {
      let value = data[header];
      if ([8, 9, 12, 14, 15].indexOf(index) !== -1) value = parseDate_(value) || '';
      values[index] = value === null || value === undefined ? '' : value;
    }
  }

  // System and underwriting fields.
  values[24] = String(data['Customer ID'] || values[24] || Utilities.getUuid());
  values[25] = new Date();
  if (Object.prototype.hasOwnProperty.call(data, 'Amana Policy Check')) values[26] = data['Amana Policy Check'] || '';
  if (Object.prototype.hasOwnProperty.call(data, 'UW Last Updated')) values[27] = parseDateTime_(data['UW Last Updated']) || new Date();

  // Do not overwrite formula columns V:X here.
  sheet.getRange(row, 1, 1, 21).setValues([values.slice(0, 21)]);
  sheet.getRange(row, 25, 1, 4).setValues([[values[24], values[25], values[26], values[27]]]);
  sheet.getRange(row, 9, 1, 2).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(row, 13, 1, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(row, 15, 1, 2).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(row, 26, 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  sheet.getRange(row, 28, 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
}

function findFirstEmptyCustomerRow_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (!String(values[i][0] || '').trim()) return i + 2;
  }
  return lastRow + 1;
}

function ensureCustomerIds_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const company = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  const ids = sheet.getRange(2, 25, lastRow - 1, 1).getDisplayValues();
  const output = ids.map(function(row, i) {
    if (!String(company[i][0] || '').trim()) return [''];
    return [String(row[0] || '').trim() || Utilities.getUuid()];
  });
  sheet.getRange(2, 25, output.length, 1).setValues(output);
}

function getFollowUps(mode) { return getFollowUps_(mode || 'due'); }

function getFollowUps_(mode) {
  const today = getToday_();
  const results = getRecords_().filter(function(record) {
    if (isClosed_(record)) return false;
    const date = parseDate_(record['Next Follow-up']);
    if (!date) return false;
    const days = differenceInDays_(today, date);
    if (mode === 'upcoming') return days > 0 && days <= 14;
    if (mode === 'today') return days === 0;
    if (mode === 'overdue') return days < 0;
    return days <= 0;
  });
  results.sort(function(a, b) {
    return parseDate_(a['Next Follow-up']) - parseDate_(b['Next Follow-up']);
  });
  return results;
}

function getRenewals() { return getRenewals_(); }

function getRenewals_() {
  const today = getToday_();
  const results = getRecords_().filter(function(record) {
    if (record['Lead Status'] === 'Lost' || record['Policy Status'] === 'Lost') return false;
    const expiry = parseDate_(record['Policy Expiry Date']);
    if (!expiry) return false;
    return differenceInDays_(today, expiry) <= 60;
  });
  results.sort(function(a, b) {
    return parseDate_(a['Policy Expiry Date']) - parseDate_(b['Policy Expiry Date']);
  });
  return results;
}

function getToday_() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function parseDate_(value) {
  if (!value) return null;
  if (value instanceof Date) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  if (typeof value === 'string') {
    const iso = value.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    const dmy = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (dmy) return new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]));
  }
  const parsed = new Date(value);
  if (isNaN(parsed.getTime())) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function parseDateTime_(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function sameDay_(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function differenceInDays_(startDate, endDate) {
  const a = Date.UTC(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const b = Date.UTC(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  return Math.round((b - a) / 86400000);
}

function isClosed_(record) {
  return ['Won / Policy Issued', 'Lost', 'Not Interested', 'Not Required'].indexOf(record['Lead Status']) !== -1;
}

/** Import the original 6-column Sheet1 only when CRM is empty. */
function importExistingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName('Sheet1');
  const crm = getCrmSheet_();
  if (!source) throw new Error('Sheet1 was not found.');
  if (getRecords_().length) throw new Error('CRM already contains customer data. Import stopped to prevent duplicates.');
  if (source.getLastRow() < 2) return 'No data found in Sheet1.';

  const sourceData = source.getRange(2, 1, source.getLastRow() - 1, 6).getValues();
  const output = [];
  sourceData.forEach(function(row) {
    if (!row[0]) return;
    const values = new Array(CRM_HEADERS.length).fill('');
    values[0] = row[0]; values[1] = row[1]; values[2] = row[2]; values[4] = row[3];
    values[6] = 'New'; values[18] = 'Medium'; values[19] = row[4]; values[20] = row[5];
    values[24] = Utilities.getUuid(); values[25] = new Date(); values[26] = 'Unknown';
    output.push(values);
  });
  if (!output.length) return 'No usable rows found in Sheet1.';
  crm.getRange(2, 1, output.length, CRM_HEADERS.length).setValues(output);
  setupFormulaRows_(crm);
  clearCrmCache_();
  return output.length + ' customers imported successfully.';
}

function testCRM() {
  return {
    app: APP_NAME,
    sheet: getCrmSheet_().getName(),
    records: getRecords_().length,
    motorRecords: getMotorPolicies().length,
    uwConfigured: Boolean(PropertiesService.getScriptProperties().getProperty('UW_WORKBOOK_ID'))
  };
}
