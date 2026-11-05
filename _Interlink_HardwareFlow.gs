/**
 * _Interlink_HardwareFlow.gs  (v0.4 Phase 5)
 * -----------------------------------------------------------------------------
 * Hardware List row → Stock Status + IMS reservation + Purchases shortfall, automatically.
 *
 * Runs when the row's "Reserve?" checkbox (H) is ticked (IM_onReserveTicked_ in _Interlink_Items.gs).
 *
 *   stock ≥ required  → Status "In Stock",  reserve the full qty
 *   stock <  required → Status "Out of Stock", reserve what exists (if any) and push the
 *                       shortfall (required − reserved) to Purchases
 *   no stock at all   → Status "Out of Stock", push the full qty
 *
 * Once a row has a reservation request (col M) or a Purchases link (col L) it is left
 * alone, so ticking again can never reserve or order the same thing twice. To redo a row,
 * clear M (and L) on purpose. A ⚠ in M (IMS error) is retried on the next tick.
 * A manual Status flip still goes through handleHardwareListEdit_ (Triggers.gs).
 *
 * Column map (Hardware List): A ID | B Project | C Material | D Spec | E Req Qty |
 *   F In Stock (IMS) | G Unit | H Reserve? | I Item Status | J Stock Status | K Remarks |
 *   L Pushed to Purchases | M IMS Reservation   (column constants: _Interlink_Items.gs)
 */

/** Pure decision: how much to reserve from stock and how much to buy. */
function IM_decideHardwareRow_(req, stock) {
  var reserve = Math.max(0, Math.min(stock, req));
  var buy = Math.round((req - reserve) * 1000) / 1000;
  return { status: buy > 0 ? 'Out of Stock' : 'In Stock', reserveQty: reserve, purchaseQty: buy };
}

/** Evaluate one Hardware row. Returns a short result line ("skipped: …" when nothing was done). */
function IM_autoEvaluateHardwareRow_(sh, row) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) return 'skipped: sheet busy, try again';
  try {
    SpreadsheetApp.flush();  // let the In Stock formula settle before reading it
    var v = sh.getRange(row, 1, 1, IM_HW_RSV_COL).getValues()[0];
    var hwId = v[0], project = v[1], material = v[2], req = Number(v[4]), stock = v[5],
        pushed = v[IM_HW_PUSHED_COL - 1], k = String(v[IM_HW_RSV_COL - 1] || '');

    if (!hwId || !project || !material || !(req > 0)) return 'skipped: needs Project, Material and Required Qty';
    if (typeof stock !== 'number') return 'skipped: not an IMS item';
    if (pushed) return 'skipped: already sent to Purchases (' + pushed + ')';
    if (k && k.charAt(0) !== '⚠') return 'skipped: reservation already requested (' + k.split(' | ')[0] + ')';

    var d = IM_decideHardwareRow_(req, stock);
    if (d.reserveQty > 0) {  // queue first: if IMS is unreachable we stop before touching Status
      var q = IM_queueReservations_([{ project: project, material: material, qty: d.reserveQty, hwId: hwId }])[0];
      sh.getRange(row, IM_HW_RSV_COL).setValue(q.reqId + ' | pending | ' + q.qty);
    }
    sh.getRange(row, IM_HW_STATUS_COL).setValue(d.status);

    var msg = d.status + (d.reserveQty > 0 ? ', reserving ' + d.reserveQty : '');
    if (d.purchaseQty > 0) {
      SpreadsheetApp.flush();
      var pur = pushHardwareItemToPurchases_(sh, row, SpreadsheetApp.getActive().getSheetByName(SHEETS.PURCHASES));
      msg += pur ? ', ' + pur + ' created for ' + d.purchaseQty : ', Purchases row NOT created';
    }
    return msg;
  } catch (e) {
    return 'error: ' + e.message;
  } finally {
    lock.releaseLock();
  }
}

// -----------------------------------------------------------------------------
// Live check of the Hardware → Purchases push. Run from the editor. Uses the first
// empty Hardware row as scratch (HW-TEST / PRJ-TEST), checks the Purchases row cell by
// cell, then clears every cell it wrote (formulas in F / M are restored or untouched).
// -----------------------------------------------------------------------------
function IM_verifyPurchasePush() {
  var ss = SpreadsheetApp.getActive();
  var hw = ss.getSheetByName(SHEETS.HARDWARE), pur = ss.getSheetByName(SHEETS.PURCHASES);
  if (!hw || !pur) { IM_say_('Hardware List / Purchases sheet missing.'); return; }

  var out = [];
  var check = function (label, ok) { out.push((ok ? 'PASS  ' : 'FAIL  ') + label); };
  var cases = [
    { name: 'free-text, qty 7',             mat: 'Selftest widget',   qty: 7,  stock: null, rsv: '',                         expect: 7 },
    { name: 'IMS 30 in stock, 30 reserved', mat: 'Selftest IMS item', qty: 50, stock: 30,   rsv: 'RSV-TEST | reserved 30',   expect: 20 },
    { name: 'IMS stock, nothing reserved',  mat: 'Selftest IMS item', qty: 50, stock: 30,   rsv: '',                         expect: null },
    { name: 'IMS zero stock',               mat: 'Selftest IMS item', qty: 50, stock: 0,    rsv: '',                         expect: 50 },
    { name: 'IMS fully reserved (pending)', mat: 'Selftest IMS item', qty: 10, stock: 30,   rsv: 'PMR-test | pending | 10',  expect: null }
  ];
  var PUR_COLS = [1, 2, 3, 4, 5, 7, 8, 10, 14, 15];  // every Purchases cell the push writes
  var HW_COLS  = [1, 2, 3, 4, 5, 7, IM_HW_STATUS_COL, IM_HW_PUSHED_COL, IM_HW_RSV_COL];  // everything but the F and I formulas
  var lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  var r = firstEmptyRow_(hw, 1);
  var reset = function () {
    HW_COLS.forEach(function (c) { hw.getRange(r, c).clearContent(); });
    hw.getRange(r, IM_STOCK_COL).setFormula(IM_stockFormula_(r));
  };
  try {
    cases.forEach(function (c) {
      hw.getRange(r, 1, 1, 5).setValues([['HW-TEST', 'PRJ-TEST', c.mat, 'selftest spec', c.qty]]);
      hw.getRange(r, 7).setValue('Pcs');
      hw.getRange(r, IM_HW_STATUS_COL).setValue('Out of Stock');
      hw.getRange(r, IM_HW_PUSHED_COL, 1, 2).setValues([['', c.rsv]]);  // L Pushed, M IMS Reservation
      if (c.stock !== null) hw.getRange(r, IM_STOCK_COL).setValue(c.stock);  // simulate an IMS item
      SpreadsheetApp.flush();

      var prow = firstEmptyRow_(pur, 1);
      var code = pushHardwareItemToPurchases_(hw, r, pur);
      if (c.expect === null) {
        check(c.name + ': nothing pushed', code === null && pur.getRange(prow, 1).getValue() === '');
      } else {
        var p = pur.getRange(prow, 1, 1, 15).getValues()[0];
        check(c.name + ': Purchases Required = Ordered = ' + c.expect, !!code && p[4] === c.expect && p[6] === c.expect);
        check(c.name + ': row mapped (project, material, unit, PO No, Not Ordered)',
              p[1] === 'PRJ-TEST' && p[2] === c.mat && p[7] === 'Pcs' && p[9] === 'No' && p[13] === 'Not Ordered');
        check(c.name + ': remark names HW-TEST' + (c.expect !== c.qty ? ' and the shortfall' : ''),
              String(p[14]).indexOf('From HW-TEST') === 0 && (c.expect === c.qty || String(p[14]).indexOf('short ' + c.expect + ' of ' + c.qty) >= 0));
        check(c.name + ': Hardware L (Pushed) = ' + code, hw.getRange(r, IM_HW_PUSHED_COL).getValue() === code);
        PUR_COLS.forEach(function (col) { pur.getRange(prow, col).clearContent(); });
      }
      reset();
    });
  } finally {
    reset();
    lock.releaseLock();
  }
  Logger.log(out.join('\n'));
  IM_say_(out.join('\n'));
  return out;
}
