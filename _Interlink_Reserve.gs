/**
 * _Interlink_Reserve.gs  (v0.4 Phase 4, PM side)
 * -----------------------------------------------------------------------------
 * PM → IMS reservations via a request queue. IMS processes 'PM Reservation
 * Requests' (its own trigger + addReservation()). Hardware List col K shows the
 * outcome:
 *
 *   K  IMS Reservation:  "PMR-xxxx | pending | 5"  →  "RSV-0012 | reserved 5"
 *                        (then follows IMS: "released 5" / "consumed 5")
 *                        or "⚠ <reason>" (a ⚠ row is re-evaluated on its next edit)
 *
 * Reservations are requested by IM_autoEvaluateHardwareRow_ (_Interlink_HardwareFlow.gs),
 * which runs automatically on edit and from the 🔗 Interlink menu. This file holds the
 * queue plumbing and the status refresh.
 *
 * No duplicates: (1) PM never queues a row whose K is already filled (unless ⚠),
 * (2) IM_queueReservations_ skips a request still Pending for the same row + item,
 * (3) IMS skips any request that already has an Active reservation for the same
 * item + reference (see IMS _Interlink_Reservations.gs).
 *
 * Needs EDIT access to the IMS sheet for whoever triggers it (it appends to the IMS
 * queue sheet). Qty is passed as-is: PM Unit and IMS UOM are assumed to match.
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
  IM_sayR_('Hardware List col K ready.');
}

function IM_sayR_(msg) {
  try { SpreadsheetApp.getUi().alert('IMS Reservation', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(msg); }
}

/**
 * Append reservation requests to the IMS queue. items: [{ project, material, qty, hwId }].
 * A request already Pending in the queue for the same row + item is reused, not repeated.
 * Returns [{ reqId, qty }] in the same order. Throws if the IMS queue sheet is missing.
 */
function IM_queueReservations_(items) {
  var q = SpreadsheetApp.openById(IM_IMS_SHEET_ID).getSheetByName(IM_RSV_QUEUE);
  if (!q) throw new Error('IMS has no "' + IM_RSV_QUEUE + '" sheet. Run IM_installReservationQueue() in IMS first.');

  var pending = {};  // "HW-001|project|item" → existing queue row
  if (q.getLastRow() >= 2) {
    q.getRange(2, 1, q.getLastRow() - 1, 11).getValues().forEach(function (r) {
      if (r[7] === 'Pending') pending[r[6] + '|' + r[3] + '|' + r[4]] = r;
    });
  }
  var me = Session.getActiveUser().getEmail() || 'unknown';
  var now = new Date(), rows = [], out = [];
  items.forEach(function (it) {
    var dup = pending[it.hwId + '|' + it.project + '|' + it.material];
    if (dup) { out.push({ reqId: dup[0], qty: dup[5] }); return; }
    var reqId = 'PMR-' + Utilities.getUuid().slice(0, 8);
    rows.push([reqId, now, me, it.project, it.material, it.qty, it.hwId, 'Pending', '', '', '']);
    out.push({ reqId: reqId, qty: it.qty });
  });
  if (rows.length) q.getRange(q.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return out;
}

/** Menu: run the row evaluation (Status, reservation, shortfall purchase) on the selected Hardware rows. */
function IM_reserveSelected() {
  var ui = SpreadsheetApp.getUi();
  var hw = SpreadsheetApp.getActiveSheet();
  if (hw.getName() !== SHEETS.HARDWARE) { ui.alert('Select rows on the Hardware List sheet first.'); return; }

  var sel = hw.getActiveRange();
  var first = Math.max(sel.getRow(), 2), last = sel.getRow() + sel.getNumRows() - 1;
  var lines = [];
  for (var row = first; row <= last; row++) {
    var id = hw.getRange(row, 1).getValue();
    if (id) lines.push(id + ': ' + IM_autoEvaluateHardwareRow_(hw, row));
  }
  ui.alert('Hardware rows', lines.length ? lines.join('\n') : 'Select at least one row that has a Hardware ID.', ui.ButtonSet.OK);
}

/**
 * Bring Hardware List col K up to date from IMS (read-only):
 *   "PMR-x | pending | n"  →  "RSV-0012 | reserved n"   (queue request processed), or "⚠ reason"
 *   "RSV-0012 | reserved n" follows the Reservation Register: reserved / released / consumed.
 * Runs on sheet open and from the 🔗 Interlink menu.
 */
function IM_refreshReservationStatus(quiet) {  // quiet = from onOpen: no toast unless something changed
  var ss = SpreadsheetApp.getActive();
  var hw = ss.getSheetByName(SHEETS.HARDWARE);
  if (!hw || hw.getLastRow() < 2) return;

  var ims = SpreadsheetApp.openById(IM_IMS_SHEET_ID);
  var q = ims.getSheetByName(IM_RSV_QUEUE), reg = ims.getSheetByName('Reservation Register');
  var byId = {}, state = {};  // request id → queue row; reservation no → Active / Released / Consumed
  if (q && q.getLastRow() >= 2) q.getRange(2, 1, q.getLastRow() - 1, 11).getValues().forEach(function (r) { byId[r[0]] = r; });
  if (reg && reg.getLastRow() >= 2) reg.getRange(2, 1, reg.getLastRow() - 1, 9).getValues().forEach(function (r) { state[r[1]] = r[8]; });
  var WORD = { Active: 'reserved', Released: 'released', Consumed: 'consumed' };

  var k = hw.getRange(2, IM_HW_RSV_COL, hw.getLastRow() - 1, 1).getValues();
  var changed = 0;
  k.forEach(function (cell, i) {
    var s = String(cell[0]), next = null, m;
    if ((m = /^(PMR-\w+) \| pending \| ([\d.]+)/.exec(s))) {
      var r = byId[m[1]];
      if (!r || r[7] === 'Pending') return;
      next = r[7] === 'Done' ? r[8] + ' | ' + (WORD[state[r[8]]] || 'reserved') + ' ' + m[2] : '⚠ ' + (r[10] || 'IMS error');
    } else if ((m = /^(RSV-\d+) \| (reserved|released|consumed) ([\d.]+)/.exec(s))) {
      var w = WORD[state[m[1]]];
      if (w) next = m[1] + ' | ' + w + ' ' + m[3];
    }
    if (next && next !== s) { hw.getRange(i + 2, IM_HW_RSV_COL).setValue(next); changed++; }
  });
  if (changed || !quiet) ss.toast(changed ? 'Updated ' + changed + ' reservation(s).' : 'No new results yet.', '🔗 Interlink', 5);
}
