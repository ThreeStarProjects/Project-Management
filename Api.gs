/**
 * ============================================================================
 * 3SP PROJECT MANAGEMENT — Api (Raise Ticket web app)
 * doGet serves Index.html publicly. api_* functions feed the form dropdowns
 * and accept the submission. Deploy as:
 *   Execute as: Me
 *   Access: Anyone (even anonymous)
 * so workers can open the URL without a Google sign-in.
 * ============================================================================
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('3SP Ticket')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Projects.Label is col I (9) after Owner/Contract were removed. */
function api_getProjectLabels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const proj = ss.getSheetByName(SHEETS.PROJECTS);
  if (!proj) return [];
  const last = proj.getLastRow();
  if (last < 2) return [];
  return proj.getRange(2, 9, last - 1, 1).getValues().map(r => r[0]).filter(Boolean);
}

function api_getEmployees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const p = ss.getSheetByName(SHEETS.PEOPLE);
  if (!p) return [];
  const last = p.getLastRow();
  if (last < 2) return [];
  return p.getRange(2, 1, last - 1, 3).getValues()
    .filter(r => r[0])
    .map(r => ({ name: String(r[0]).trim(), department: String(r[1] || '').trim(), role: String(r[2] || '').trim() }));
}

function api_getVendors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const v = ss.getSheetByName(SHEETS.VENDORS);
  if (!v) return [];
  const last = v.getLastRow();
  if (last < 2) return [];
  return v.getRange(2, 1, last - 1, 5).getValues()
    .filter(r => r[0])
    .map(r => ({ name: String(r[0]).trim(), contact: String(r[1] || '').trim(), category: String(r[2] || '').trim(),
                 phone: String(r[3] || '').trim(), email: String(r[4] || '').trim() }));
}

function api_getTicketConfig() {
  return { categories: TICKET_CATEGORIES, priorities: TICKET_PRIORITIES, impacts: TICKET_IMPACTS };
}

function api_submitTicket(formData) {
  if (!formData) return { error: 'MISSING_DATA', message: 'No form data received.' };
  const raisedBy   = String(formData.raisedBy || '').trim();
  const category   = String(formData.category || '').trim();
  const subject    = String(formData.subject || '').trim();
  const assignedTo = String(formData.assignedTo || '').trim();
  if (!raisedBy)   return { error: 'MISSING_NAME',     message: 'Your name is required.' };
  if (!category)   return { error: 'MISSING_CATEGORY', message: 'Category is required.' };
  if (!subject)    return { error: 'MISSING_SUBJECT',  message: 'Subject is required.' };
  if (!assignedTo) return { error: 'MISSING_ASSIGNEE', message: 'Please pick who to assign this to.' };
  const cfg = TICKET_CATEGORIES[category];
  if (!cfg) return { error: 'INVALID_CATEGORY', message: 'Unknown category: ' + category };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.TICKETS);
  if (!sh) return { error: 'SHEET_MISSING', message: 'Tickets sheet not found.' };

  const nextId = 'TKT-' + String(getMaxIdNumber_(sh, 'TKT') + 1).padStart(3, '0');
  const newRow = firstEmptyRow_(sh, 1);
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let dueDate;
  if (formData.dueDate) {
    const m = String(formData.dueDate).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) dueDate = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  }
  if (!dueDate) { dueDate = new Date(today); dueDate.setDate(dueDate.getDate() + cfg.dueDays); }
  const priority = String(formData.priority || cfg.priority).trim();
  const impact = String(formData.impact || priority).trim();
  const amountRaw = formData.amount;
  const amount = (amountRaw === '' || amountRaw === null || amountRaw === undefined)
    ? '' : parseFloat(String(amountRaw).replace(/[₹,\s]/g, '')) || '';

  // 23-col row layout — matches createTicketsSheet_
  const row = [
    nextId, today, raisedBy, category, subject,
    String(formData.description || '').trim(),
    priority, assignedTo, 'Open', dueDate,
    String(formData.relatedProject || '').trim(),
    String(formData.machineId || '').trim(),
    String(formData.batchCode || '').trim(),
    String(formData.vendorName || '').trim(),
    String(formData.drawingRev || '').trim(),
    amount, impact, '', '', '',
    String(formData.remarks || '').trim()
    // V and W (Overdue Flag, TAT) are ARRAYFORMULAs — do NOT write.
  ];
  for (let i = 0; i < row.length; i++) {
    sh.getRange(newRow, i + 1).setValue(row[i]);
  }
  return { success: true, ticketId: nextId, assignedTo: assignedTo, priority: priority,
           dueDate: Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') };
}
