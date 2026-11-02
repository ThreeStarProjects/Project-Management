/**
 * _Interlink_Reserve.gs  (v0.4 Phase 4, PM side)
 * -----------------------------------------------------------------------------
 * "Reserve for project": select Hardware List rows → queue reservation requests
 * in IMS ('PM Reservation Requests', processed by IMS every minute via its own
 * addReservation()). Hardware List col K shows the outcome.
 *
 *   K  IMS Reservation:  "PMR-xxxx | pending | 5"  →  "RSV-0012 | reserved 5"
 *                        or "⚠ <reason>" (a ⚠ row can be re-requested)
 *
 * Reserves min(Required Qty, In Stock). Any shortfall stays on the row for the
 * Purchases flow (Phase 5). Qty is passed as-is, so the PM Unit and the IMS UOM
 * are assumed to match.
 *
 * Needs EDIT access to the IMS sheet for whoever clicks the button (it appends
 * to the IMS queue sheet). IMS must have run IM_installReservationQueue().
 *
 * One-time: run IM_installReserve() from the editor (adds the K header).
 */

var IM_HW_RSV_COL = 11;  // K
var IM_RSV_QUEUE  = 'PM Reservation Requests';  // same name as IMS side

function IM_installReserve() {
  var hw = SpreadsheetApp.getActive().getSheetByName(SHEETS.HARDWARE);
  if (!hw) { IM_sayR_('Hardware List sheet not found.'); return; }
  hw.getRange(1, IM_HW_RSV_COL).setValue('IMS Reservation')
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  hw.setColumnWidth(IM_HW_RSV_COL, 200);
  IM_sayR_('Hardware List col K ready. Use 🔗 Interlink → Reserve selected rows in IMS.');
}

function IM_sayR_(msg) {
  try { SpreadsheetApp.getUi().alert('IMS Reservation', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(msg); }
}

/** Selected Hardware List rows → IMS reservation requests. */
function IM_reserveSelected() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var hw = ss.getActiveSheet();
  if (hw.getName() !== SHEETS.HARDWARE) { ui.alert('Select rows on the Hardware List sheet first.'); return; }

  var sel = hw.getActiveRange();
  var first = Math.max(sel.getRow(), 2), last = sel.getRow() + sel.getNumRows() - 1;
  if (last < first) { ui.alert('Select at least one data row.'); return; }

  var labels = {};
  ss.getSheetByName(IM_ITEMS_SHEET).getRange('B2:B').getValues().forEach(function (r) { if (r[0]) labels[r[0]] = true; });

  var me = Session.getActiveUser().getEmail() || 'unknown';
  var now = new Date();
  var queue = [], marks = [], skipped = [];
  for (var row = first; row <= last; row++) {
    var v = hw.getRange(row, 1, 1, IM_HW_RSV_COL).getValues()[0];
    var hwId = v[0], project = v[1], material = v[2], qty = Number(v[4]), stock = Number(v[5]), prior = String(v[10] || '');
    if (!hwId) continue;
    if (prior && prior.charAt(0) !== '⚠') { skipped.push(hwId + ': already requested'); continue; }
    if (!labels[material])   { skipped.push(hwId + ': Material is not an IMS item'); continue; }
    if (!project)            { skipped.push(hwId + ': no Project Label'); continue; }
    if (!(qty > 0))          { skipped.push(hwId + ': no Required Qty'); continue; }
    var n = Math.min(qty, isNaN(stock) ? 0 : stock);
    if (!(n > 0))            { skipped.push(hwId + ': nothing in stock'); continue; }
    var reqId = 'PMR-' + Utilities.getUuid().slice(0, 8);
    queue.push([reqId, now, me, project, material, n, hwId, 'Pending', '', '', '']);
    marks.push([row, reqId + ' | pending | ' + n]);
  }
  if (!queue.length) { ui.alert('Nothing to reserve.\n\n' + skipped.join('\n')); return; }

  var ok = ui.alert('Reserve in IMS?',
    queue.length + ' row(s) will be reserved in IMS (qty = min of Required and In Stock).' +
    (skipped.length ? '\n\nSkipped:\n' + skipped.join('\n') : ''), ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  var q = SpreadsheetApp.openById(IM_IMS_SHEET_ID).getSheetByName(IM_RSV_QUEUE);
  if (!q) { ui.alert('IMS has no "' + IM_RSV_QUEUE + '" sheet. Run IM_installReservationQueue() in IMS first.'); return; }
  var start = q.getLastRow() + 1;
  q.getRange(start, 1, queue.length, queue[0].length).setValues(queue);
  marks.forEach(function (m) { hw.getRange(m[0], IM_HW_RSV_COL).setValue(m[1]); });
  ss.toast('Queued ' + queue.length + ' reservation(s). IMS processes them within a minute; use Refresh reservation status.', '🔗 Interlink', 8);
}

/** Pull Done / Error results from the IMS queue back onto Hardware List col K. */
function IM_refreshReservationStatus() {
  var ss = SpreadsheetApp.getActive();
  var hw = ss.getSheetByName(SHEETS.HARDWARE);
  if (!hw || hw.getLastRow() < 2) return;

  var k = hw.getRange(2, IM_HW_RSV_COL, hw.getLastRow() - 1, 1).getValues();
  var q = SpreadsheetApp.openById(IM_IMS_SHEET_ID).getSheetByName(IM_RSV_QUEUE);
  if (!q || q.getLastRow() < 2) return;
  var byId = {};
  q.getRange(2, 1, q.getLastRow() - 1, 11).getValues().forEach(function (r) { byId[r[0]] = r; });

  var changed = 0;
  k.forEach(function (cell, i) {
    var m = /^(PMR-\w+) \| pending \| ([\d.]+)/.exec(String(cell[0]));
    var r = m && byId[m[1]];
    if (!r || r[7] === 'Pending') return;
    hw.getRange(i + 2, IM_HW_RSV_COL).setValue(
      r[7] === 'Done' ? r[8] + ' | reserved ' + m[2] : '⚠ ' + (r[10] || 'IMS error'));
    changed++;
  });
  ss.toast(changed ? 'Updated ' + changed + ' reservation(s).' : 'No new results yet.', '🔗 Interlink', 5);
}
