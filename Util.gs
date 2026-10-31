/**
 * ============================================================================
 * 3SP PROJECT MANAGEMENT — Util
 * Row-level helpers used across triggers, menu actions, and the ticket API.
 * ============================================================================
 */

function getMaxIdNumber_(sheet, prefix) {
  if (!sheet) return 0;
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const vals = sheet.getRange(2, 1, last - 1, 1).getValues();
  let max = 0;
  for (let i = 0; i < vals.length; i++) {
    const v = String(vals[i][0] || '');
    if (v.startsWith(prefix + '-')) {
      const n = parseInt(v.split('-')[1], 10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return max;
}

function firstEmptyRow_(sheet, col) {
  const last = sheet.getLastRow();
  if (last < 1) return 2;
  const vals = sheet.getRange(2, col, Math.max(last - 1, 1), 1).getValues();
  for (let i = 0; i < vals.length; i++) if (!vals[i][0]) return i + 2;
  return last + 1;
}

function findRowByValue_(sheet, col, value) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const vals = sheet.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === value) return i + 2;
  return null;
}

function deleteAllRowsByValue_(sheet, col, value) {
  if (!sheet) return 0;
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const vals = sheet.getRange(2, col, last - 1, 1).getValues();
  let deleted = 0;
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] === value) { sheet.deleteRow(i + 2); deleted++; }
  }
  return deleted;
}
