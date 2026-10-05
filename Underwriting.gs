const UW_HEADERS = [
  'UW ID', 'Customer ID', 'Company Name', 'Product', 'Sales Request',
  'Submitted Date', 'UW Status', 'Amana Policy Check', 'UW Response',
  'UW Reviewer', 'Next Action', 'Due Date', 'Last Updated', 'Sales Notes'
];

/**
 * Creates a separate Google Spreadsheet for underwriting collaboration.
 * This is deliberate: Google Sheets sharing is workbook-level, not tab-level.
 */
function setupUnderwritingWorkspace() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty('UW_WORKBOOK_ID');
  if (existingId) {
    try {
      const existing = SpreadsheetApp.openById(existingId);
      setupUnderwritingSheet_(existing);
      return underwritingWorkspaceInfo_();
    } catch (e) {
      props.deleteProperty('UW_WORKBOOK_ID');
    }
  }

  const ss = SpreadsheetApp.create(APP_NAME + ' - Underwriting Workbench');
  const first = ss.getSheets()[0];
  first.setName(UW_SHEET_NAME);
  setupUnderwritingSheet_(ss);
  props.setProperty('UW_WORKBOOK_ID', ss.getId());
  return underwritingWorkspaceInfo_();
}

function setupUnderwritingSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, UW_SHEET_NAME, UW_HEADERS);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, UW_HEADERS.length).setFontWeight('bold').setWrap(true);
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  setDropdown_(sheet, 7, OPTIONS.uwStatus, rows);
  setDropdown_(sheet, 8, OPTIONS.amanaPolicyCheck, rows);
  setDropdown_(sheet, 11, OPTIONS.actions, rows);
  sheet.getRange(2, 6, rows, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  sheet.getRange(2, 12, rows, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(2, 13, rows, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  [150, 180, 220, 160, 300, 150, 140, 180, 320, 150, 170, 110, 150, 280]
    .forEach(function(width, i) { sheet.setColumnWidth(i + 1, width); });
  return sheet;
}

function underwritingWorkspaceInfo_() {
  const id = PropertiesService.getScriptProperties().getProperty('UW_WORKBOOK_ID');
  if (!id) return { configured: false };
  const ss = SpreadsheetApp.openById(id);
  return { configured: true, id: id, name: ss.getName(), url: ss.getUrl() };
}

function getUnderwritingWorkspaceInfo() {
  try { return underwritingWorkspaceInfo_(); }
  catch (e) { return { configured: false, error: e.message }; }
}

function getUwSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('UW_WORKBOOK_ID');
  if (!id) throw new Error('Underwriting workspace is not configured. Click “Create UW Workspace” first.');
  const ss = SpreadsheetApp.openById(id);
  return setupUnderwritingSheet_(ss);
}

function pushCustomerToUnderwriting(customerRow, product, requestText, dueDate, salesNotes) {
  const customer = getCustomer(customerRow);
  const sheet = getUwSheet_();
  const row = sheet.getLastRow() + 1;
  const values = [
    Utilities.getUuid(), customer['Customer ID'], customer['Company Name'],
    product || customer['Products'] || '', requestText || '', new Date(),
    'Sent to UW', customer['Amana Policy Check'] || 'Unknown', '', '',
    'Awaiting Underwriting', parseDate_(dueDate) || '', new Date(), salesNotes || ''
  ];
  sheet.getRange(row, 1, 1, UW_HEADERS.length).setValues([values]);
  sheet.getRange(row, 6).setNumberFormat('dd/mm/yyyy hh:mm');
  sheet.getRange(row, 12).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(row, 13).setNumberFormat('dd/mm/yyyy hh:mm');

  const crm = getCrmSheet_();
  crm.getRange(Number(customerRow), 17).setValue('Sent to UW');
  crm.getRange(Number(customerRow), 28).setValue(new Date()).setNumberFormat('dd/mm/yyyy hh:mm');
  crm.getRange(Number(customerRow), 11).setValue('Awaiting Underwriting');
  clearCrmCache_();
  return getUnderwritingItems();
}

function getUnderwritingItems() {
  const info = getUnderwritingWorkspaceInfo();
  if (!info.configured) return { workspace: info, items: [] };
  const sheet = getUwSheet_();
  if (sheet.getLastRow() < 2) return { workspace: info, items: [] };
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, UW_HEADERS.length).getValues();
  const items = [];
  values.forEach(function(row, index) {
    if (!String(row[0] || '').trim()) return;
    const item = {};
    UW_HEADERS.forEach(function(header, i) {
      if ([5, 11, 12].indexOf(i) !== -1 && row[i] instanceof Date) {
        item[header] = Utilities.formatDate(row[i], Session.getScriptTimeZone(), i === 11 ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ss");
      } else item[header] = row[i] || '';
    });
    item._row = index + 2;
    items.push(item);
  });
  items.sort(function(a, b) { return Number(b._row) - Number(a._row); });
  return { workspace: info, items: items };
}

/** Pulls underwriter changes back into matching CRM customer rows. */
function syncUnderwritingToCrm() {
  const payload = getUnderwritingItems();
  if (!payload.workspace.configured) throw new Error('Underwriting workspace is not configured.');
  const sheet = getCrmSheet_();
  const records = getRecords_();
  const byId = {};
  records.forEach(function(record) { byId[record['Customer ID']] = record; });

  const latest = {};
  payload.items.forEach(function(item) {
    const id = item['Customer ID'];
    if (!id) return;
    if (!latest[id] || Number(item._row) > Number(latest[id]._row)) latest[id] = item;
  });

  let updated = 0;
  Object.keys(latest).forEach(function(customerId) {
    const record = byId[customerId];
    if (!record) return;
    const item = latest[customerId];
    const row = record._row;
    const uwRemarks = [item['UW Response'], item['UW Reviewer'] ? 'UW: ' + item['UW Reviewer'] : '']
      .filter(Boolean).join('\n');
    sheet.getRange(row, 17).setValue(item['UW Status'] || '');
    sheet.getRange(row, 20).setValue(uwRemarks);
    sheet.getRange(row, 27).setValue(item['Amana Policy Check'] || 'Unknown');
    sheet.getRange(row, 28).setValue(new Date()).setNumberFormat('dd/mm/yyyy hh:mm');
    if (item['Next Action']) sheet.getRange(row, 11).setValue(item['Next Action']);
    updated++;
  });
  clearCrmCache_();
  return { updated: updated, payload: getUnderwritingItems() };
}
