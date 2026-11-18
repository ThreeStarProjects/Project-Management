/**
 * ============================================================================
 * 3SP PROJECT MANAGEMENT — Util
 * Row-level helpers used across triggers, menu actions, and the ticket API.
 * ============================================================================
 */

// Rows above the table: Hardware List / Purchases / Production carry TOP_ROWS control rows (see Config.gs). Every other sheet: 0.
function topRows_(name) { return FILTER_SHEETS.indexOf(name) >= 0 ? TOP_ROWS : 0; }
function hdrRow_(sheet) { return 1 + topRows_(sheet.getName()); }     // the header row
function firstRow_(sheet) { return 2 + topRows_(sheet.getName()); }   // the first data row
/** A1 text written for the OLD layout (rows from 2) -> the same cells on a sheet with control rows, clamped to the sheet's size. */
function shiftA1_(sheet, a1) {
  const top = topRows_(sheet.getName());
  if (!top) return a1;
  const max = sheet.getMaxRows();
  return a1.replace(/([A-Z]+)(\d+)/g, function (m, c, r) { return c + Math.min(Number(r) + top, max); });
}

function getMaxIdNumber_(sheet, prefix) {
  if (!sheet) return 0;
  const last = sheet.getLastRow(), f = firstRow_(sheet);
  if (last < f) return 0;
  const vals = sheet.getRange(f, 1, last - f + 1, 1).getValues();
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
  const last = sheet.getLastRow(), f = firstRow_(sheet);
  if (last < f) return f;
  const vals = sheet.getRange(f, col, Math.max(last - f + 1, 1), 1).getValues();
  for (let i = 0; i < vals.length; i++) if (!vals[i][0]) return i + f;
  return last + 1;
}

function findRowByValue_(sheet, col, value) {
  const last = sheet.getLastRow(), f = firstRow_(sheet);
  if (last < f) return null;
  const vals = sheet.getRange(f, col, last - f + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === value) return i + f;
  return null;
}

