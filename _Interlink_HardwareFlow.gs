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

// -----------------------------------------------------------------------------
// Stock Status (col J) fills itself for rows that are not yet ticked / reserved / ordered:
//   IMS item with Project, Material and Required Qty filled → "In Stock" if the free stock (col F =
//   IMS Available, i.e. already net of every reservation) covers the Required Qty, else "Out of Stock".
// It only ever writes the cell; nothing is reserved or ordered until Reserve? is ticked (script writes do
// not fire the Out-of-Stock → Purchases rule). Runs when B / C / E of a row is edited, when PM opens and
// from 🔗 Interlink → Refresh stock status. A row with a tick, a reservation (M) or a Purchases link (L)
// is never touched. A value the user chose by hand is respected: the script remembers what it wrote
// (hidden sheet _IMS_Items, T = Hardware ID, U = value) and only overwrites a cell that is blank or
// still holds its own value.
// -----------------------------------------------------------------------------
var IM_AUTO_ID_COL  = 20;  // T on _IMS_Items
var IM_AUTO_VAL_COL = 21;  // U on _IMS_Items

/** Pure rule. '' = no opinion (free-text material, or Required Qty / stock not usable). */
function IM_stockStatusFor_(req, stock) {
  if (!(req > 0) || typeof stock !== 'number') return '';
  return stock >= req ? 'In Stock' : 'Out of Stock';
}

/** only = null (every row) or [row numbers]. Returns how many Stock Status cells were changed. */
function IM_refreshStockStatus_(only) {
  var ss = SpreadsheetApp.getActive();
  var hw = ss.getSheetByName(SHEETS.HARDWARE), it = ss.getSheetByName(IM_ITEMS_SHEET);
  if (!hw || !it || hw.getLastRow() < 2) return 0;
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) return 0;
  try {
    if (it.getMaxColumns() < IM_AUTO_VAL_COL) it.insertColumnsAfter(it.getMaxColumns(), IM_AUTO_VAL_COL - it.getMaxColumns());
    SpreadsheetApp.flush();  // let the In Stock formula settle before reading it
    var n = hw.getLastRow() - 1;
    var v = hw.getRange(2, 1, n, IM_HW_RSV_COL).getValues();
    var ml = it.getLastRow(), mine = {};  // Hardware ID → value this script last wrote
    if (ml >= 2) it.getRange(2, IM_AUTO_ID_COL, ml - 1, 2).getValues().forEach(function (r) { if (r[0] !== '') mine[r[0]] = r[1]; });
    var want = null;
    if (only) { want = {}; only.forEach(function (r) { want[r] = true; }); }

    var changed = 0, keep = [];
    v.forEach(function (r, i) {
      var id = r[0], cur = r[IM_HW_STATUS_COL - 1];
      if (!id) return;
      if (r[IM_HW_RESERVE_COL - 1] === true || r[IM_HW_PUSHED_COL - 1] || r[IM_HW_RSV_COL - 1]) return;  // decided: hands off
      var owned = cur === '' || (mine[id] !== undefined && cur === mine[id]);  // blank, or still the script's own value
      if (!owned) return;                                                       // chosen by hand: respect it
      var target = (r[1] && r[2]) ? IM_stockStatusFor_(Number(r[4]), r[5]) : '';
      if ((!want || want[i + 2]) && target && cur !== target) { r[IM_HW_STATUS_COL - 1] = target; cur = target; changed++; }
      if (cur !== '') keep.push([id, cur]);
    });
    if (changed) hw.getRange(2, IM_HW_STATUS_COL, n, 1).setValues(v.map(function (r) { return [r[IM_HW_STATUS_COL - 1]]; }));
    if (ml >= 2) it.getRange(2, IM_AUTO_ID_COL, ml - 1, 2).clearContent();
    if (keep.length) it.getRange(2, IM_AUTO_ID_COL, keep.length, 2).setValues(keep);
    return changed;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menu: bring the whole Hardware List up to date in one click: IMS Reservation (col M) from the IMS Register,
 * and Stock Status (col J) of the rows that are not ticked yet. Both also run by themselves when PM opens.
 */
function IM_refreshHardwareNow() {
  IM_refreshReservationStatus(true);
  var n = IM_refreshStockStatus_(null);
  SpreadsheetApp.getActive().toast('Reservations refreshed from IMS; ' + (n ? 'Stock Status updated on ' + n + ' row(s).' : 'Stock Status already up to date.'), '🔗 Interlink', 5);
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
