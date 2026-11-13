/**
 * _Interlink_Reserve.gs  (v0.4 Phase 4, PM side)
 * -----------------------------------------------------------------------------
 * PM → IMS reservations through IMS's own 'Reservation Register' (one sheet for all reservations).
 * PM appends a row with Status "Requested" and a Request ID; IMS's 1-minute trigger turns it into
 * an Active reservation (Reservation No assigned) or an "Error" row. Hardware List col M shows the
 * outcome:
 *
 *   M  IMS Reservation:  "PMR-xxxx | pending | 5"  →  "RSV-0012 | reserved 5"
 *                        (then follows IMS: "partially consumed 5" / "consumed 5" / "released 5")
 *                        or "⚠ <reason>" (a ⚠ row is re-evaluated on its next edit)
 *
 * Reservations are requested by IM_autoEvaluateHardwareRow_ (_Interlink_HardwareFlow.gs),
 * which runs when the row's "Reserve?" checkbox (col H) is ticked. This file holds the
 * queue plumbing and the status refresh.
 *
 * No duplicates: (1) PM never queues a row whose M is already filled (unless ⚠),
 * (2) IM_queueReservations_ skips a request still Requested for the same row + item,
 * (3) IMS drops any request that already has an open reservation for the same
 * item + reference (see IMS _Interlink_Reservations.gs).
 *
 * Needs EDIT access to the IMS sheet for whoever triggers it (it appends to the IMS
 * Register). Qty is passed as-is: PM Unit and IMS UOM are assumed to match.
 *
 * One-time: run IM_installReserve() from the editor (adds the M header).
 */

var IM_RSV_REGISTER = 'Reservation Register';  // IMS sheet; A:P, see IMS _Interlink_Reservations.gs

function IM_installReserve() {
  var hw = SpreadsheetApp.getActive().getSheetByName(SHEETS.HARDWARE);
  if (!hw) { IM_sayR_('Hardware List sheet not found.'); return; }
  hw.getRange(1, IM_HW_RSV_COL).setValue('IMS Reservation')
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  hw.setColumnWidth(IM_HW_RSV_COL, 200);
  IM_sayR_('Hardware List col M ready.');
}

function IM_sayR_(msg) {
  try { SpreadsheetApp.getUi().alert('IMS Reservation', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(msg); }
}

/**
 * Append reservation requests to the IMS Reservation Register as Status "Requested".
 * items: [{ project, material, qty, hwId }]. A request already waiting (Requested) for the same
 * row + item is reused, not repeated. Returns [{ reqId, qty }] in the same order.
 * Throws if the IMS Register has no Request ID column yet (run IM_installReservationRegister() in IMS).
 */
function IM_queueReservations_(items) {
  var reg = SpreadsheetApp.openById(IM_IMS_SHEET_ID).getSheetByName(IM_RSV_REGISTER);
  if (!reg || reg.getRange(1, 16).getValue() !== 'Request ID') {
    throw new Error('IMS Reservation Register is not set up for PM requests. Run IM_installReservationRegister() in IMS first.');
  }
  var last = reg.getLastRow();
  var vals = last >= 2 ? reg.getRange(2, 1, last - 1, 16).getValues() : [];
  var waiting = {}, lastSr = 1;  // "forRef|item" → request row; lastSr = last used row number (col A), 1 = header only
  vals.forEach(function (r, i) {
    if (r[0] !== '') lastSr = i + 2;
    if (r[8] === 'Requested') waiting[r[7] + '|' + r[3]] = r;
  });
  var me = Session.getActiveUser().getEmail() || 'unknown';
  var now = new Date(), rows = [], out = [];
  items.forEach(function (it) {
    var forRef = it.project + ' / ' + it.hwId;
    var dup = waiting[forRef + '|' + it.material];
    if (dup) { out.push({ reqId: dup[15], qty: dup[5] }); return; }
    var reqId = 'PMR-' + Utilities.getUuid().slice(0, 8);
    var row = lastSr + 1 + rows.length;
    // A Sr | B Res No (IMS fills) | C Date | D Item | E Desc | F Qty | G UOM | H For | I Status | J Notes | K By | L At | M | N | O Consumed | P Request ID
    rows.push([row - 1, '', now, it.material, '', it.qty, '', forRef, 'Requested', '', me, now, '', '', '', reqId]);
    out.push({ reqId: reqId, qty: it.qty });
  });
  if (rows.length) reg.getRange(lastSr + 1, 1, rows.length, 16).setValues(rows);
  return out;
}

/**
 * Bring Hardware List col M up to date from IMS (read-only):
 *   "PMR-x | pending | n"  →  "RSV-0012 | reserved n"   (request processed), or "⚠ reason"
 *   "RSV-0012 | reserved n" follows the Reservation Register: reserved / partially consumed / released / consumed.
 * Runs on sheet open and from the 🔗 Interlink menu.
 */
function IM_refreshReservationStatus(quiet) {  // quiet = from onOpen: no toast unless something changed
  var ss = SpreadsheetApp.getActive();
  var hw = ss.getSheetByName(SHEETS.HARDWARE);
  if (!hw || hw.getLastRow() < 2) return;

  var reg = SpreadsheetApp.openById(IM_IMS_SHEET_ID).getSheetByName(IM_RSV_REGISTER);
  var byId = {}, state = {};  // request id → Register row; reservation no → Register row
  if (reg && reg.getLastRow() >= 2) reg.getRange(2, 1, reg.getLastRow() - 1, 16).getValues().forEach(function (r) {
    if (r[15]) byId[r[15]] = r;
    if (r[1]) state[r[1]] = r;
  });
  // Register row → word used in col M. A released reservation that was partly consumed first reads "partially consumed".
  var word = function (r) {
    var s = r[8], used = Number(r[14]) || 0;
    if (s === 'Active') return 'reserved';
    if (s === 'Partially Consumed' || (s === 'Released' && used > 0)) return 'partially consumed';
    return s === 'Released' ? 'released' : s === 'Consumed' ? 'consumed' : null;
  };

  var k = hw.getRange(2, IM_HW_RSV_COL, hw.getLastRow() - 1, 1).getValues();
  var changed = 0;
  k.forEach(function (cell, i) {
    var s = String(cell[0]), next = null, m;
    if ((m = /^(PMR-\w+) \| pending \| ([\d.]+)/.exec(s))) {
      var r = byId[m[1]];
      if (!r || r[8] === 'Requested') return;
      next = r[8] === 'Error' ? '⚠ ' + (r[9] || 'IMS error') : r[1] + ' | ' + (word(r) || 'reserved') + ' ' + m[2];
    } else if ((m = /^(RSV-\d+) \| (reserved|partially consumed|released|consumed) ([\d.]+)/.exec(s))) {
      var w = state[m[1]] && word(state[m[1]]);
      if (w) next = m[1] + ' | ' + w + ' ' + m[3];
    }
    if (next && next !== s) {
      hw.getRange(i + 2, IM_HW_RSV_COL).setValue(next); changed++;
      var cell = hw.getRange(i + 2, IM_HW_ITEMSTATUS_COL);   // hover the Item Status cell to see why, in plain words
      if (next.charAt(0) === '⚠') cell.setNote('IMS refused this reservation. ' + IM_friendlyText_(next.replace(/^⚠\s*/, '')));
      else if (s.charAt(0) === '⚠') cell.clearNote();
    }
  });
  if (changed || !quiet) ss.toast(changed ? 'Updated ' + changed + ' reservation(s).' : 'No new results yet.', '🔗 Interlink', 5);
}
