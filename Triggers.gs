/**
 * ============================================================================
 * 3SP PROJECT MANAGEMENT — Triggers
 * Installable onEdit handler: auto-ID generation + Hardware List → Purchases push.
 * Activity logging was removed when the Log sheet was dropped.
 * ============================================================================
 */

function onEditHandler(e) {
  try {
    if (!e || !e.range) return;
    handleAutoIdGeneration_(e);
    handleHardwareListEdit_(e);
  } catch (err) {
    Logger.log('onEdit error: ' + err.message + '\n' + err.stack);
  }
}

/**
 * Fills col A with an auto-ID when a new row gets content elsewhere.
 * Sheet → prefix mapping lives in ID_PREFIX_MAP.
 */
function handleAutoIdGeneration_(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const prefix = ID_PREFIX_MAP[sheet.getName()];
  if (!prefix) return;
  const startRow = range.getRow(), numRows = range.getNumRows();
  const startCol = range.getColumn(), numCols = range.getNumColumns();
  if (startRow < 2) return;
  if (startCol === 1 && numCols === 1) return; // user typing ID manually

  const colARange = sheet.getRange(startRow, 1, numRows, 1);
  const colAValues = colARange.getValues();
  const editedValues = range.getValues();
  const rowsNeedingId = [];
  for (let i = 0; i < numRows; i++) {
    if (colAValues[i][0] !== '' && colAValues[i][0] !== null) continue;
    if (editedValues[i].some(v => v !== '' && v !== null)) rowsNeedingId.push(i);
  }
  if (!rowsNeedingId.length) return;
  let maxNum = getMaxIdNumber_(sheet, prefix);
  rowsNeedingId.forEach(i => { maxNum++; colAValues[i][0] = prefix + '-' + String(maxNum).padStart(3, '0'); });
  colARange.setValues(colAValues);
}

/**
 * When a Hardware List row's Status flips to "Out of Stock", push it into Purchases.
 * One-way. No re-push if already pushed (col I holds the PUR-code).
 * No sync of subsequent edits to Purchases (avoids clobbering procurement work).
 */
function handleHardwareListEdit_(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== SHEETS.HARDWARE) return;
  if (range.getRow() < 2) return;

  const startCol = range.getColumn();
  const numCols  = range.getNumColumns();
  if (startCol > 7 || startCol + numCols - 1 < 7) return;  // col G = Status

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const purSheet = ss.getSheetByName(SHEETS.PURCHASES);
  if (!purSheet) return;

  const startRow = range.getRow();
  const numRows  = range.getNumRows();
  for (let r = startRow; r < startRow + numRows; r++) {
    pushHardwareItemToPurchases_(sheet, r, purSheet);
  }
}

/**
 * Push one Hardware List row into Purchases if Status=Out of Stock and not already pushed.
 * Returns the new PUR-code, or null if skipped.
 */
function pushHardwareItemToPurchases_(hwSheet, hwRow, purSheet) {
  const row = hwSheet.getRange(hwRow, 1, 1, 9).getValues()[0];
  const [hwId, project, material, spec, qty, unit, status, /*remarks*/, pushed] = row;
  if (status !== 'Out of Stock') return null;
  if (pushed) return null;
  if (!hwId || !material) return null;

  const purNum = getMaxIdNumber_(purSheet, 'PUR') + 1;
  const purCode = 'PUR-' + String(purNum).padStart(3, '0');
  const newRow = firstEmptyRow_(purSheet, 1);

  // Purchases cols: A ID, B Project, C Material, D Spec, E Required, F Ordered, G Unit,
  //                 H Vendor, I PO Issued (Y/N), J Expected, K Actual, L Variance (formula),
  //                 M Status, N Remarks
  purSheet.getRange(newRow, 1).setValue(purCode);
  purSheet.getRange(newRow, 2).setValue(project);
  purSheet.getRange(newRow, 3).setValue(material);
  purSheet.getRange(newRow, 4).setValue(spec || '');
  purSheet.getRange(newRow, 5).setValue(qty || '');
  purSheet.getRange(newRow, 6).setValue(qty || '');  // Ordered defaults to Required
  purSheet.getRange(newRow, 7).setValue(unit || '');
  purSheet.getRange(newRow, 9).setValue('No');       // PO Issued default
  purSheet.getRange(newRow, 13).setValue('Not Ordered');
  purSheet.getRange(newRow, 14).setValue('From ' + hwId);

  hwSheet.getRange(hwRow, 9).setValue(purCode);
  return purCode;
}
