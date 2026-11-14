/**
 * _Interlink_Items.gs  (v0.4 Phase 3)
 * -----------------------------------------------------------------------------
 * IMS items on Hardware List and Purchases.
 *
 *   _IMS_Items (hidden)  IMPORTRANGE of IMS Item Master + SFG Master + stock
 *   C Material           dropdown of IMS items (warn-only: free text still allowed)
 *   F In Stock (IMS)     Available − Reserved, one plain formula per row
 *
 * Both sheets get the same new column F (right after Required Qty); everything
 * from the old F onward moves one column right. Hardware List also gets
 * H "Reserve?" (a checkbox: tick it to send the row's reservation to IMS) and
 * I "Item Status", a per-row formula over M (IMS Reservation) and L (Pushed to Purchases):
 *   Hardware List: A ID | B Project | C Material | D Spec | E Req Qty | F In Stock |
 *                  G Unit | H Reserve? | I Item Status | J Stock Status | K Remarks |
 *                  L Pushed to Purchases | M IMS Reservation
 *   Purchases:     A ID | B Project | C Material | D Spec | E Req Qty | F In Stock |
 *                  G Ordered | H Unit | I Vendor | J PO Issued | K Expected | L Actual |
 *                  M Variance | N Status | O Remarks
 * Picking an item in C fills Unit (G on Hardware, H on Purchases) if blank.
 *
 * Hardware List col J is "Stock Status" (In Stock / Out of Stock dropdown); col I is "Item Status".
 *
 * Item Status (I) is a live per-row formula; no script has to run for it to change:
 *   Reservation: Not Reserved (IMS item, nothing reserved) / Reserving… / Reserved / Released /
 *     Consumed / ⚠ Reservation error. It follows IMS directly: _IMS_Items imports IMS's
 *     'Reservation Register' (K:M by Request ID, O:R by Reservation No), so a release or consume
 *     in IMS shows up here by itself. Col M's text is only the fallback.
 *   Order (from the linked Purchases row's Status): To order / Ordered / Delivered.
 *   Joined as e.g. "Reserved + Ordered". Free-text (non-IMS) materials never show "Not Reserved".
 *
 * Rollout (push the code, then straight away, from the editor):
 *   1. IM_installItems()      idempotent; inserts col F in the live sheets (native
 *                             insert shifts existing formulas / validations / colours),
 *                             then re-seeds ARRAYFORMULAs from the updated code.
 *   2. IM_installItemStatus() (re)writes the IMS imports and the Item Status formulas (idempotent).
 *                             On a sheet still in the old layout (no "Reserve?" column) run
 *                             IM_installReserveCheckbox() instead: it inserts H and re-seeds I.
 *   3. Open _IMS_Items (View → Hidden sheets), click A2, press "Allow access".
 *   4. IM_verifyWiring() + IM_verifyItemStatus()   PASS/FAIL lists.
 * Re-run IM_installItems() and IM_installItemStatus() after "Create/Rebuild Hardware List Sheet".
 */

var IM_IMS_SHEET_ID = '1CglvPxUWZnIHNeToGEDz8-Xj3eNfvaF351j_ogpRTgk';
var IM_ITEMS_SHEET  = '_IMS_Items';
var IM_MAT_COL      = 3;   // C
var IM_STOCK_COL    = 6;   // F
var IM_STOCK_HDR    = 'In Stock (IMS)';
var IM_HW_RESERVE_COL    = 8;    // H  checkbox: tick = send this row's reservation to IMS
var IM_HW_ITEMSTATUS_COL = 9;    // I
var IM_HW_STATUS_COL     = 10;   // J  Stock Status (In Stock / Out of Stock)
var IM_HW_PUSHED_COL     = 12;   // L  Pushed to Purchases (PUR code)
var IM_HW_RSV_COL        = 13;   // M  IMS Reservation text
var IM_RESERVE_HDR       = 'Reserve?';
var IM_ITEMSTATUS_HDR    = 'Item Status';
var IM_STOCKSTATUS_HDR   = 'Stock Status';   // Hardware List col J (was "Status")

/** Sheets that carry the Material dropdown + In Stock column, and where their Unit lives. */
function IM_itemSheets_() {
  return [
    { name: SHEETS.HARDWARE,  unitCol: 7 },
    { name: SHEETS.PURCHASES, unitCol: 8 }
  ];
}

function IM_installItems() {
  var ss = SpreadsheetApp.getActive();
  var targets = IM_itemSheets_().filter(function (t) { return ss.getSheetByName(t.name); });
  if (!targets.length) { IM_say_('Hardware List / Purchases sheets not found.'); return; }

  // --- helper sheet ---------------------------------------------------------
  var it = ss.getSheetByName(IM_ITEMS_SHEET) || ss.insertSheet(IM_ITEMS_SHEET);
  it.clear();
  var id = IM_IMS_SHEET_ID;
  it.getRange('A1:E1').setValues([['Code', 'Label', 'UOM', 'Item Name', 'Net Stock']]);
  it.getRange('H1:I1').setValues([['Stock Code', 'Stock Net']]);
  // Items + SFGs. Search Label = Item Master T / SFG Master N.
  // ponytail: no Active? filter (errored on IMS data); inactive items show. Add back once column type is known.
  it.getRange('A2').setFormula(
    '={QUERY(IMPORTRANGE("' + id + '","Item Master!A2:T"),"select Col1,Col20,Col7,Col4 where Col1 is not null",0);' +
    'QUERY(IMPORTRANGE("' + id + '","SFG Master!A2:N"),"select Col1,Col14,Col7,Col4 where Col1 is not null",0)}');
  // Stock: RM = Snapshot Available − Reserved; SFG = SFG Master Available (already net).
  it.getRange('H2').setFormula(
    '={QUERY(IMPORTRANGE("' + id + '","Stock Snapshot!A2:J"),"select Col1,Col9-Col10 where Col1 is not null",0);' +
    'QUERY(IMPORTRANGE("' + id + '","SFG Master!A2:M"),"select Col1,Col13 where Col1 is not null",0)}');
  it.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(VLOOKUP(A2:A,H:I,2,FALSE),0)))');
  IM_writeStatusImports_(it);
  it.hideSheet();

  // --- per-sheet column migration -----------------------------------------
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(it.getRange('B2:B5000'), true)
    .setAllowInvalid(true).build();  // warn-only so free-text materials still work

  targets.forEach(function (t) {
    var sh = ss.getSheetByName(t.name);
    if (sh.getRange(1, IM_STOCK_COL).getValue() !== IM_STOCK_HDR) {
      // earlier Phase 3 attempt put Item Code / In Stock in J:K of Hardware List
      if (t.name === SHEETS.HARDWARE && sh.getRange(1, 10).getValue() === 'Item Code (IMS)') sh.deleteColumns(10, 2);
      sh.insertColumnAfter(IM_STOCK_COL - 1);  // new F; native insert shifts formulas/validations/CF
    }
    sh.getRange(1, IM_STOCK_COL).setValue(IM_STOCK_HDR)
      .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG)
      .setFontWeight('bold').setFontSize(11)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sh.setColumnWidth(IM_STOCK_COL, 110);

    // F: one plain formula per row (survives row deletes and the Purchases import)
    var f = [];
    for (var r = 2; r <= MAX_ROWS; r++) f.push([IM_stockFormula_(r)]);
    sh.getRange(2, IM_STOCK_COL, MAX_ROWS - 1, 1).setFormulas(f).setHorizontalAlignment('center');
    sh.getRange(2, IM_MAT_COL, MAX_ROWS - 1, 1).setDataValidation(rule);
  });

  // Re-seed the ARRAYFORMULA cells (Purchases M2 etc.) from the updated code so they
  // are canonical, not just natively shifted. Project View / Dashboard cells are
  // checked by IM_verifyWiring().
  reseedArrayFormulas_(ss);

  // --- autofill trigger (separate from the existing onEditHandler) ---------
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'IM_onHardwareItemEdit_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('IM_onHardwareItemEdit_').forSpreadsheet(ss).onEdit().create();

  IM_say_('Columns inserted. Now open the hidden "' + IM_ITEMS_SHEET + '" sheet (View → Hidden sheets), ' +
          'click A2 and press "Allow access". Then run IM_verifyWiring().');
}

/** In Stock (IMS) formula for one row: item stock for IMS items, "n/a" for free text, blank if no material. */
function IM_stockFormula_(r) {
  return '=IF($C' + r + '="","",IFERROR(VLOOKUP($C' + r + ',' + IM_ITEMS_SHEET + '!$B:$E,4,FALSE),"n/a"))';
}

/**
 * IMS reservation data for Item Status, imported into the hidden helper sheet (live, no script):
 *   K:M  Reservation Register (rows with a Request ID) → Request ID, Reservation No, Status
 *   O:R  Reservation Register → Reservation No, Status, Qty Reserved, Consumed Qty
 *        Status: Requested / Active / Partially Consumed / Consumed / Released / Error
 */
function IM_writeStatusImports_(it) {
  var id = IM_IMS_SHEET_ID;
  it.getRange('K1:M1').setValues([['Request ID', 'Res No', 'Queue Status']]);
  it.getRange('K2').setFormula(
    '=IFERROR(QUERY(IMPORTRANGE("' + id + '","Reservation Register!A2:P"),"select Col16,Col2,Col9 where Col16 is not null",0),"")');
  it.getRange('O1:R1').setValues([['Res No', 'Register Status', 'Qty', 'Consumed Qty']]);
  it.getRange('O2').setFormula(
    '=IFERROR(QUERY(IMPORTRANGE("' + id + '","Reservation Register!A2:P"),"select Col2,Col9,Col6,Col15 where Col2 is not null",0),"")');
}

/**
 * Hardware List "Item Status" (col I) for one row, from M (IMS Reservation text), L (PUR code) and F (stock).
 * Reservation part, live from IMS via the helper sheet (falls back to M's text if IMS has no row for it):
 *   Consumed / Partially Consumed / Released / Reserved (Register status) · ⚠ Reservation error ·
 *   Reserving… (request still pending) · Not Reserved (IMS item with no reservation yet) · blank for
 *   free-text materials. A reservation released after a part-consumption reads "Partially Consumed
 *   (rest released)".
 * Order part (linked Purchases Status): To order (Not Ordered) / Ordered (PO Issued, In Transit) /
 *   Delivered / otherwise that status text.  Both present → "Reserved + Ordered".
 */
function IM_itemStatusFormula_(r) {
  var I = IM_ITEMS_SHEET, P = SHEETS.PURCHASES;
  return '=LET(resvRaw,$M' + r + ',purch,$L' + r + ',isIms,ISNUMBER($F' + r + '),' +
    'reqId,IFERROR(REGEXEXTRACT(resvRaw,"^PMR-\\w+"),""),' +
    'rsvNo,IFERROR(REGEXEXTRACT(resvRaw,"^RSV-\\d+"),IF(reqId="","",IFERROR(VLOOKUP(reqId,' + I + '!$K:$M,2,FALSE),""))),' +
    'qStat,IF(reqId="","",IFERROR(VLOOKUP(reqId,' + I + '!$K:$M,3,FALSE),"")),' +
    'regStat,IF(rsvNo="","",IFERROR(VLOOKUP(rsvNo,' + I + '!$O:$R,2,FALSE),"")),' +
    'used,IF(rsvNo="",0,IFERROR(VLOOKUP(rsvNo,' + I + '!$O:$R,4,FALSE)+0,0)),' +
    'liveState,IF(regStat="Active","Reserved",IF(regStat="Partially Consumed","Partially Consumed",' +
      'IF(regStat="Released",IF(used>0,"Partially Consumed (rest released)","Released"),IF(regStat="Consumed","Consumed","")))),' +
    'textState,IF(REGEXMATCH(resvRaw,"partially consumed"),"Partially Consumed",IF(REGEXMATCH(resvRaw,"consumed"),"Consumed",' +
      'IF(REGEXMATCH(resvRaw,"released"),"Released",IF(REGEXMATCH(resvRaw,"reserved"),"Reserved","")))),' +
    'state,IF(liveState<>"",liveState,textState),' +
    'resvText,IF(state<>"",state,IF(OR(qStat="Error",LEFT(resvRaw,1)="⚠"),"⚠ Reservation error",' +
      'IF(reqId<>"","Reserving…",IF(isIms,"Not Reserved","")))),' +
    'pstat,IF(purch="","",IFERROR(VLOOKUP(purch,' + P + '!$A:$N,14,FALSE),"")),' +
    'ordText,IF(purch="","",IF(OR(pstat="PO Issued",pstat="In Transit"),"Ordered",IF(pstat="Delivered","Delivered",' +
      'IF(OR(pstat="Not Ordered",pstat=""),"To order",pstat)))),' +
    'IF(AND(resvText<>"",ordText<>""),resvText&" + "&ordText,resvText&ordText))';
}

/**
 * Hardware List: (re)write the IMS imports and the per-row Item Status formulas (col I). Idempotent.
 * Needs the current layout (G Unit | H Reserve? | I Item Status). A sheet from before the checkbox
 * column goes through IM_installReserveCheckbox() first.
 */
function IM_installItemStatus() {
  var ss = SpreadsheetApp.getActive();
  var hw = ss.getSheetByName(SHEETS.HARDWARE), it = ss.getSheetByName(IM_ITEMS_SHEET);
  if (!hw) { IM_say_('Hardware List sheet not found.'); return; }
  if (!it) { IM_say_('Run IM_installItems() first (the hidden ' + IM_ITEMS_SHEET + ' sheet is missing).'); return; }
  if (hw.getRange(1, IM_HW_RESERVE_COL).getValue() !== IM_RESERVE_HDR) {
    IM_say_('Hardware List has no "' + IM_RESERVE_HDR + '" column in H yet. Run IM_installReserveCheckbox() instead; nothing changed.');
    return;
  }
  IM_writeStatusImports_(it);
  IM_styleHwHeader_(hw, IM_HW_ITEMSTATUS_COL, IM_ITEMSTATUS_HDR, 150);
  hw.getRange(2, IM_HW_ITEMSTATUS_COL, MAX_ROWS - 1, 1).clearDataValidations();
  hw.getRange(1, IM_HW_STATUS_COL).setValue(IM_STOCKSTATUS_HDR);

  var f = [];
  for (var r = 2; r <= MAX_ROWS; r++) f.push([IM_itemStatusFormula_(r)]);
  hw.getRange(2, IM_HW_ITEMSTATUS_COL, MAX_ROWS - 1, 1).setFormulas(f).setHorizontalAlignment('center');
  IM_say_('Hardware List: I "Item Status" ready. Open the hidden ' + IM_ITEMS_SHEET +
          ' sheet and click K2 and O2 to press "Allow access" if they show an error. Then run IM_verifyWiring() and IM_verifyItemStatus().');
}

function IM_styleHwHeader_(hw, col, text, width) {
  hw.getRange(1, col).setValue(text)
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  hw.setColumnWidth(col, width);
}

/**
 * Hardware List: insert col H "Reserve?" (checkboxes) right after Unit. Item Status moves to I,
 * Stock Status to J, Remarks K, Pushed to Purchases L, IMS Reservation M (native insert shifts every
 * live formula, validation and colour rule). Rows that already have a reservation or a Purchases link
 * come out ticked. Then re-seeds the Item Status formulas for the new letters. Idempotent.
 */
function IM_installReserveCheckbox() {
  var ss = SpreadsheetApp.getActive();
  var hw = ss.getSheetByName(SHEETS.HARDWARE), it = ss.getSheetByName(IM_ITEMS_SHEET);
  if (!hw || !it) { IM_say_('Hardware List or the hidden ' + IM_ITEMS_SHEET + ' sheet is missing.'); return; }
  if (hw.getRange(1, IM_HW_RESERVE_COL).getValue() !== IM_RESERVE_HDR) {
    if (hw.getRange(1, 7).getValue() !== 'Unit' || hw.getRange(1, 8).getValue() !== IM_ITEMSTATUS_HDR) {
      IM_say_('Expected "Unit" in G and "Item Status" in H of Hardware List; nothing changed.');
      return;
    }
    hw.insertColumnAfter(7);  // new H
  }
  hw.getRange(2, IM_HW_RESERVE_COL, MAX_ROWS - 1, 1).clearDataValidations().insertCheckboxes().setHorizontalAlignment('center');
  IM_styleHwHeader_(hw, IM_HW_RESERVE_COL, IM_RESERVE_HDR, 80);

  // rows that were already reserved / ordered under the old flow start ticked (a ⚠ row stays unticked)
  var last = hw.getLastRow();
  if (last >= 2) {
    var cur = hw.getRange(2, IM_HW_RESERVE_COL, last - 1, 1).getValues();
    var lm = hw.getRange(2, IM_HW_PUSHED_COL, last - 1, 2).getValues();  // L, M
    var ticks = lm.map(function (x, i) { return [cur[i][0] === true || !!x[0] || (!!x[1] && String(x[1]).charAt(0) !== '⚠')]; });
    hw.getRange(2, IM_HW_RESERVE_COL, last - 1, 1).setValues(ticks);
  }
  IM_installItemStatus();
}

var IM_QUIET_ = false;  // set by the self-test: log instead of popping up dialogs
/**
 * Keep exactly ONE trigger for each of PM's edit / open handlers. A second copy makes every edit run the
 * same code twice (guards stop double reservations, but it is wasted work and doubled toasts).
 * Runs on open and in the self-test. Returns how many extra triggers were removed.
 */
function IM_removeDuplicateTriggers_() {
  var KEEP = { onEditHandler: 1, IM_onHardwareItemEdit_: 1, IM_onOpen_: 1 }, seen = {}, removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var h = t.getHandlerFunction();
    if (!KEEP[h]) return;
    if (seen[h]) { ScriptApp.deleteTrigger(t); removed++; } else seen[h] = true;
  });
  return removed;
}

function IM_say_(msg) {
  if (IM_QUIET_) { Logger.log(msg); return; }
  try { SpreadsheetApp.getUi().alert('Items', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(msg); }  // run from editor: no UI context
}

/**
 * Edit trigger for Hardware List / Purchases.
 *  - Material (C) picked  → fill Unit if blank (Hardware G, Purchases H).
 *  - Hardware List "Reserve?" (H) ticked → IM_autoEvaluateHardwareRow_ for that row
 *    (Stock Status set, reservation queued, shortfall pushed to Purchases).
 *  - Hardware List Project (B), Material (C) or Required Qty (E) edited → IM_refreshStockStatus_ for that row
 *    (Stock Status fills itself for unticked rows; see _Interlink_HardwareFlow.gs).
 */
function IM_onHardwareItemEdit_(e) {
  try {
    if (!e || !e.range) return;
    var r = e.range, sh = r.getSheet();
    var t = IM_itemSheets_().filter(function (x) { return x.name === sh.getName(); })[0];
    if (!t || r.getRow() < 2 || r.getNumRows() !== 1 || r.getNumColumns() !== 1) return;
    var col = r.getColumn();
    if (col === IM_MAT_COL) IM_fillUnit_(e, sh, t, r);
    if (t.name === SHEETS.HARDWARE && col === IM_HW_RESERVE_COL && r.getValue() === true) IM_onReserveTicked_(e, sh, r.getRow());
    if (t.name === SHEETS.HARDWARE && (col === 2 || col === IM_MAT_COL || col === 5)) IM_refreshStockStatus_([r.getRow()]);  // Stock Status fills itself
  } catch (err) {
    Logger.log('IM_onHardwareItemEdit_ error: ' + err.message);
  }
}

/**
 * Reserve? ticked. The tick stays only while the row is (or just became) sent; a row that cannot
 * be sent (incomplete, free-text material, IMS unreachable) is unticked again with the reason.
 */
function IM_onReserveTicked_(e, sh, row) {
  var res = IM_autoEvaluateHardwareRow_(sh, row);
  var failed = res.indexOf('skipped') === 0 || res.indexOf('error') === 0;
  if (failed && !/already (requested|sent)/.test(res)) sh.getRange(row, IM_HW_RESERVE_COL).setValue(false);
  var msg = failed ? IM_friendlyText_(res) : res + (/reserving/.test(res) ? '. IMS confirms within about a minute.' : '');
  e.source.toast(msg, 'Hardware List', 10);
}

/** Fill Unit from the picked IMS item's UOM, if the Unit cell is blank. */
function IM_fillUnit_(e, sh, t, r) {
  try {
    var label = String(r.getValue());
    if (!label) return;

    var rows = e.source.getSheetByName(IM_ITEMS_SHEET).getRange('A2:D').getValues();
    var m = rows.filter(function (x) { return x[1] === label; })[0];
    if (!m) return;  // free-text material

    // IMS UOM → PM Unit dropdown value; unmapped UOMs (e.g. Liters) are left for the user.
    // IMS also lists PM's own names exactly (see IMS _Interlink_Units.gs), so those map to themselves.
    var uomMap = { 'no.s': 'Pcs', 'nos': 'Pcs', 'pcs': 'Pcs', 'units': 'Pcs', 'unit': 'Pcs',
                   'pieces': 'Pcs', 'piece': 'Pcs', 'kilograms': 'Kg', 'kg': 'Kg',
                   'meters': 'Meter', 'meter': 'Meter', 'sets': 'Set', 'set': 'Set',
                   'boxes': 'Box', 'box': 'Box', 'rolls': 'Roll', 'roll': 'Roll',
                   'square feet': 'Sq.ft', 'sq.ft': 'Sq.ft', 'square meters': 'Sq.m', 'sq.m': 'Sq.m',
                   'other': 'Other' };
    var u = uomMap[String(m[2]).toLowerCase()];
    var cell = sh.getRange(r.getRow(), t.unitCol);
    if (u && !cell.getValue()) cell.setValue(u);
  } catch (err) {
    Logger.log('IM_fillUnit_ error: ' + err.message);
  }
}

// -----------------------------------------------------------------------------
// Wiring check — run after IM_installItems(). Logs one PASS/FAIL line per check.
// -----------------------------------------------------------------------------
function IM_verifyWiring() {
  var ss = SpreadsheetApp.getActive();
  var out = [];
  var check = function (label, ok) { out.push((ok ? 'PASS  ' : 'FAIL  ') + label); };
  var hdr = function (sheet, col) {
    var sh = ss.getSheetByName(sheet);
    return sh ? String(sh.getRange(1, col).getValue()) : '(missing sheet)';
  };
  var dvHas = function (sheet, a1, value) {
    var sh = ss.getSheetByName(sheet);
    var dv = sh && sh.getRange(a1).getDataValidation();
    var vals = dv ? dv.getCriteriaValues()[0] : null;
    return Array.isArray(vals) && vals.indexOf(value) >= 0;
  };

  var H = SHEETS.HARDWARE, P = SHEETS.PURCHASES;
  [[H, 5, 'Required Qty'], [H, 6, IM_STOCK_HDR], [H, 7, 'Unit'], [H, 8, IM_RESERVE_HDR], [H, 9, IM_ITEMSTATUS_HDR], [H, 10, IM_STOCKSTATUS_HDR],
   [H, 12, 'Pushed to Purchases'], [H, 13, 'IMS Reservation'],
   [P, 5, 'Required Qty'], [P, 6, IM_STOCK_HDR], [P, 7, 'Ordered Qty'], [P, 8, 'Unit'], [P, 9, 'Vendor'],
   [P, 10, 'PO Issued'], [P, 11, 'Expected Delivery'], [P, 12, 'Actual Delivery'], [P, 13, 'Variance (Days)'],
   [P, 14, 'Status'], [P, 15, 'Remarks']
  ].forEach(function (c) { check(c[0] + ' header col ' + c[1] + ' = ' + c[2], hdr(c[0], c[1]) === c[2]); });

  var pur = ss.getSheetByName(P);
  var mf = pur ? pur.getRange('M2').getFormula() : '';
  check('Purchases M2 variance = L − K', mf.indexOf('L2:L-K2:K') >= 0 && mf.indexOf('K2:K=""') >= 0);

  var hwSheet = ss.getSheetByName(H);
  check('Hardware I2 Item Status formula reads L and M', !!hwSheet && /\$M2/.test(hwSheet.getRange('I2').getFormula()) && /\$L2/.test(hwSheet.getRange('I2').getFormula()));
  var cbDv = hwSheet && hwSheet.getRange('H2').getDataValidation();
  check('Hardware H (Reserve?) is a checkbox', !!cbDv && cbDv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX);
  check('Hardware J (Stock Status) dropdown', dvHas(H, 'J2', 'Out of Stock'));
  check('Hardware G (Unit) dropdown',   dvHas(H, 'G2', 'Pcs'));
  check('Hardware I (Item Status) has no dropdown', !!hwSheet && !hwSheet.getRange('I2').getDataValidation());
  check('Purchases N (Status) dropdown', dvHas(P, 'N2', 'Delivered'));
  check('Purchases J (PO Issued) dropdown', dvHas(P, 'J2', 'Yes'));
  check('Purchases H (Unit) dropdown',   dvHas(P, 'H2', 'Pcs'));

  var pv = ss.getSheetByName(SHEETS.VIEW);
  if (pv) {
    var fs = pv.getRange(1, 1, 70, 5).getFormulas();
    var flat = [].concat.apply([], fs).filter(String);
    check('Project View: no formula still reads Purchases!M:M',
          !flat.some(function (f) { return /Purchases'?!M:M/.test(f); }));
    check('Project View: Delivered count reads Purchases!N:N',
          flat.some(function (f) { return /Purchases'?!N:N/.test(f) && f.indexOf('Delivered') >= 0; }));
    var a59 = pv.getRange('A59').getFormula();
    check('Project View A59 query uses G2:H, I2:I, N2:N',
          a59.indexOf('G2:H') >= 0 && a59.indexOf('I2:I') >= 0 && a59.indexOf('N2:N') >= 0);
  }

  var handlers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  check('trigger onEditHandler installed',          handlers.indexOf('onEditHandler') >= 0);
  check('trigger IM_onHardwareItemEdit_ installed', handlers.indexOf('IM_onHardwareItemEdit_') >= 0);

  Logger.log(out.join('\n'));
  IM_say_(out.join('\n'));
  return out;
}

// -----------------------------------------------------------------------------
// Live check of the Item Status formula (Hardware List col I). Uses the first empty
// Hardware row and first empty Purchases row as scratch, then clears what it wrote.
// -----------------------------------------------------------------------------
function IM_verifyItemStatus() {
  var ss = SpreadsheetApp.getActive();
  var hw = ss.getSheetByName(SHEETS.HARDWARE), pur = ss.getSheetByName(SHEETS.PURCHASES);
  if (!hw || !pur) { IM_say_('Hardware List / Purchases sheet missing.'); return; }

  // [F stock (number = IMS item, "n/a" = free text), M IMS Reservation, L PUR code, linked Purchases status, expected]
  // These use made-up RSV/PMR numbers, so they exercise the text fallback. The live IMS lookup
  // (Register / queue imports) is checked by the first real reservation: reserve, then release in IMS.
  var cases = [
    [30,    '',                     '',         '',            'Not Reserved'],
    [30,    'PMR-t | pending | 5',  '',         '',            'Reserving…'],
    [30,    'RSV-1 | reserved 5',   '',         '',            'Reserved'],
    [30,    'RSV-1 | released 5',   '',         '',            'Released'],
    [30,    'RSV-1 | consumed 5',   '',         '',            'Consumed'],
    [30,    'RSV-1 | partially consumed 5', '', '',                'Partially Consumed'],
    [30,    '⚠ IMS error',          '',         '',            '⚠ Reservation error'],
    ['n/a', '',                     '',         '',            ''],
    ['n/a', '',                     'PUR-TEST', 'Not Ordered', 'To order'],
    ['n/a', '',                     'PUR-TEST', 'In Transit',  'Ordered'],
    ['n/a', '',                     'PUR-TEST', 'Delivered',   'Delivered'],
    [30,    '',                     'PUR-TEST', 'Not Ordered', 'Not Reserved + To order'],
    [0,     '',                     'PUR-TEST', 'PO Issued',   'Not Reserved + Ordered'],
    [30,    'RSV-1 | reserved 5',   'PUR-TEST', 'PO Issued',   'Reserved + Ordered'],
    [30,    'RSV-1 | consumed 5',   'PUR-TEST', 'Delivered',   'Consumed + Delivered'],
    [30,    'RSV-1 | partially consumed 5', 'PUR-TEST', 'PO Issued', 'Partially Consumed + Ordered']
  ];
  var out = [];
  var lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  var hr = firstEmptyRow_(hw, 1), pr = firstEmptyRow_(pur, 1);
  var reset = function () {
    hw.getRange(hr, IM_HW_PUSHED_COL, 1, 2).clearContent();   // L, M
    hw.getRange(hr, IM_STOCK_COL).setFormula(IM_stockFormula_(hr));
    pur.getRange(pr, 1).clearContent(); pur.getRange(pr, 14).clearContent();  // A ID, N Status
  };
  try {
    cases.forEach(function (c) {
      hw.getRange(hr, IM_STOCK_COL).setValue(c[0]);           // simulate an IMS item / a free-text one
      hw.getRange(hr, IM_HW_PUSHED_COL, 1, 2).setValues([[c[2], c[1]]]);
      pur.getRange(pr, 1).setValue(c[2] ? 'PUR-TEST' : '');
      pur.getRange(pr, 14).setValue(c[3]);
      SpreadsheetApp.flush();
      var got = String(hw.getRange(hr, IM_HW_ITEMSTATUS_COL).getValue());
      out.push((got === c[4] ? 'PASS  ' : 'FAIL  ') + 'F=' + c[0] + ' M="' + c[1] + '" L="' + c[2] + '" purchase="' + c[3] + '" → "' + got + '"' +
               (got === c[4] ? '' : ' (expected "' + c[4] + '")'));
    });
  } finally {
    reset();
    lock.releaseLock();
  }
  Logger.log(out.join('\n'));
  IM_say_(out.join('\n'));
  return out;
}

// =============================================================================
// SYSTEM SELF-TEST  (🔗 Interlink → Run system self-test)
// -----------------------------------------------------------------------------
// Walks the whole chain on LIVE data with a throw-away project "ZZ-SELFTEST …" and deletes it at the end
// (the real cross-system project delete). Runs as steps advanced by a 1-minute trigger, so it can wait for
// IMS (1-min trigger) and Finance (10-min trigger) without hitting the 6-minute limit; ~15 minutes in all.
//   preflight → project + deadlines → Masters → pick IMS items → Hardware rows (stock status) → reserve
//   (+ Purchases shortfall, double-tick guard) → IMS reservation → statuses → requisition #1 (partial)
//   → requisitions #2 + direct issue → Finance import → cleanup (delete everywhere) → report.
// Everything is logged to the visible sheet "Self-test report" as PASS / FAIL / WARN / INFO with the exact
// error text. Safety: nothing is posted or deleted unless the project name contains ZZ-SELFTEST; IMS only
// accepts self-test requisitions for such projects (qty 1-5); the cleanup stops if the delete preview shows
// more than a handful of rows. Limits: typing in cells (onEdit) is not simulated; only the trigger's existence
// is checked. It burns a few document numbers (REQ / RSV / COST) for good.
// =============================================================================
var ST_PROP = 'IM_SELFTEST';
var ST_REPORT = 'Self-test report';
var ST_FN = 'IM_selfTestTick';
var ST_TAG = 'ZZ-SELFTEST';
var ST_CMD = '_SelfTest_Cmd';

function ST_load_() {
  var raw = PropertiesService.getScriptProperties().getProperty(ST_PROP);
  return raw ? JSON.parse(raw) : null;
}
function ST_save_(st) { PropertiesService.getScriptProperties().setProperty(ST_PROP, JSON.stringify(st)); }

function ST_reportSheet_(fresh) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(ST_REPORT);
  if (!sh) sh = ss.insertSheet(ST_REPORT);
  if (fresh) {
    sh.clear();
    sh.getRange(1, 1, 1, 5).setValues([['Self-test', '', '', '', '']]).setFontWeight('bold').setFontSize(13);
    sh.getRange(3, 1, 1, 5).setValues([['Time', 'Area', 'Check', 'Result', 'Detail']]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
    sh.setFrozenRows(3);
    sh.setColumnWidth(1, 70); sh.setColumnWidth(2, 130); sh.setColumnWidth(3, 380); sh.setColumnWidth(4, 70); sh.setColumnWidth(5, 620);
  }
  return sh;
}

/** One report line. result: PASS | FAIL | WARN | INFO | SKIP. Returns true for PASS. */
function ST_log_(area, check, result, detail) {
  var sh = ST_reportSheet_(false), row = Math.max(sh.getLastRow() + 1, 4);
  sh.getRange(row, 1, 1, 5).setValues([[Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss'), area, check, result, String(detail === undefined ? '' : detail)]]);
  var color = { PASS: '#C6EFCE', FAIL: '#FFC7CE', WARN: '#FFEB9C', INFO: '#DDEBF7', SKIP: '#EDEDED' }[result];
  if (color) sh.getRange(row, 4).setBackground(color);
  return result === 'PASS';
}
function ST_chk_(area, check, ok, detail) { return ST_log_(area, check, ok ? 'PASS' : 'FAIL', detail); }

function ST_ims_() { return SpreadsheetApp.openById(IM_IMS_SHEET_ID); }
function ST_finBook_() {
  var id = Masters.getSystemSheetIds().finance;
  if (!id) throw new Error('Finance has not registered its spreadsheet with Masters yet (open Finance once)');
  return SpreadsheetApp.openById(id);
}
function ST_pad_(n) { return (n < 10 ? '00' : n < 100 ? '0' : '') + n; }
function ST_date_(offsetDays) { var d = new Date(); d.setHours(0, 0, 0, 0); d.setDate(d.getDate() + offsetDays); return d; }
function ST_dstr_(v) { return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v).slice(0, 10); }
function ST_close_(a, b, tol) { return Math.abs(Number(a) - Number(b)) <= (tol === undefined ? 0.001 : tol); }

function ST_runChecks_(area, lines) {
  (lines || []).forEach(function (l) {
    var s = String(l), pass = /^PASS/.test(s);
    ST_log_(area, s.replace(/^(PASS|FAIL)\s+/, ''), pass ? 'PASS' : 'FAIL', '');
  });
}

// ---------------------------------------------------------------- steps (true = finished, false = keep waiting)

function ST_preflight_(st) {
  IM_healthChecks_().forEach(function (r) { ST_log_(r.area, r.check, r.result, r.detail + (r.fix ? '  → ' + r.fix : '')); });
  ST_runChecks_('PM wiring', IM_verifyWiring());
  ST_runChecks_('Item Status formula', IM_verifyItemStatus());
  ST_runChecks_('Hardware → Purchases', IM_verifyPurchasePush());
  return true;
}

function ST_project_(st) {
  var A = 'Project', ss = SpreadsheetApp.getActive(), proj = ss.getSheetByName(SHEETS.PROJECTS);
  st.name = ST_TAG + ' ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMdd-HHmm');
  st.id = addProject_(st.name);
  st.label = st.id + ' — ' + st.name;
  st.pmRow = findRowByValue_(proj, 1, st.id);
  var mile = ss.getSheetByName(SHEETS.MILESTONES);
  if (mile && !findRowByValue_(mile, 1, st.label)) mile.getRange(firstEmptyRow_(mile, 1), 1).setValue(st.label);
  ST_save_(st);
  ST_chk_(A, 'Add project: next PRJ-### id and name written', /^PRJ-\d+$/.test(st.id) && !!st.pmRow, st.label);

  var deadline = function (offset, expect) {
    proj.getRange(st.pmRow, 5).setValue(ST_date_(-1)); proj.getRange(st.pmRow, 6).setValue(ST_date_(offset));
    SpreadsheetApp.flush();
    var v = proj.getRange(st.pmRow, 9, 1, 3).getValues()[0];  // I label, J delay status, K days left
    ST_chk_(A, 'Deadline ' + (offset >= 0 ? '+' : '') + offset + ' days → Delay Status "' + expect + '"', String(v[1]) === expect, 'Delay Status = "' + v[1] + '", days left = ' + v[2]);
    return v;
  };
  var v0 = deadline(-2, 'Overdue');
  ST_chk_(A, 'Label formula builds "PRJ-### — name"', String(v0[0]) === st.label, String(v0[0]));
  deadline(3, 'Near Deadline');
  var v2 = deadline(30, 'On Track');
  ST_chk_(A, 'Days left counts down from the handover date', ST_close_(v2[2], 30, 1), 'days left = ' + v2[2]);

  var sync = IM_syncProjectsToMasters_();
  ST_chk_(A, 'Project sent to Masters (sync)', sync.pushed >= 1 && sync.errors.length === 0, 'pushed ' + sync.pushed + (sync.errors.length ? '; ' + sync.errors.join('; ') : ''));
  var m = (Masters.getProjects() || []).filter(function (p) { return String(p.Project_ID).toUpperCase() === st.id; })[0];
  ST_chk_(A, 'Masters holds the project: same id and name, Active', !!m && m.Name === st.name && String(m['Active?']) === 'Yes', m ? 'Masters: ' + m.Project_ID + ' — ' + m.Name : 'not found in Masters');
  ST_chk_(A, 'Masters holds the handover date', !!m && ST_dstr_(m.Target_Handover) === Utilities.formatDate(ST_date_(30), Session.getScriptTimeZone(), 'yyyy-MM-dd'), m ? String(m.Target_Handover) : '');
  return true;
}

function ST_items_(st) {
  var A = 'IMS items', it = SpreadsheetApp.getActive().getSheetByName(IM_ITEMS_SHEET);
  var rates = {}, im = ST_ims_().getSheetByName('Item Master');
  var head = im.getRange(1, 1, 1, im.getLastColumn()).getValues()[0], cRate = head.indexOf('Avg Rate (Rs.)');
  if (cRate >= 0 && im.getLastRow() >= 2) im.getRange(2, 1, im.getLastRow() - 1, im.getLastColumn()).getValues().forEach(function (r) { if (r[0]) rates[r[0]] = Number(r[cRate]) || 0; });
  var rows = it.getRange(2, 1, Math.max(it.getLastRow() - 1, 1), 5).getValues()
    .filter(function (r) { return r[0] && r[1] && typeof r[4] === 'number'; })
    .map(function (r) { return { code: String(r[0]), label: String(r[1]), uom: String(r[2]), name: String(r[3]), stock: r[4], rate: rates[r[0]] || 0 }; });
  var isTest = function (x) { return /test/i.test(x.name + ' ' + x.label); };
  var pick = function (list) { return list.filter(isTest)[0] || list[0]; };
  var xs = rows.filter(function (x) { return x.stock >= 3 && x.rate > 0; }).sort(function (a, b) { return a.stock - b.stock; });
  st.x = pick(xs);
  var zs = rows.filter(function (x) { return !st.x || x.code !== st.x.code; });
  st.z = pick(zs);
  if (st.z) st.z.qty = Math.ceil(st.z.stock) + 5;
  ST_save_(st);
  ST_chk_(A, 'An IMS item with stock >= 3 and an Avg Rate exists for the in-stock / requisition path', !!st.x, st.x ? st.x.code + ' (' + st.x.name + '): stock ' + st.x.stock + ', Avg Rate ' + st.x.rate : 'none found: in-stock, reservation, requisition and Finance steps are skipped');
  if (st.z) ST_log_(A, 'Out-of-stock item', 'INFO', st.z.code + ' (' + st.z.name + '): stock ' + st.z.stock + ', will ask for ' + st.z.qty);
  else ST_log_(A, 'A second IMS item exists for the out-of-stock path', 'SKIP', 'only one item in IMS');
  return true;
}

function ST_hardware_(st) {
  var A = 'Hardware List', ss = SpreadsheetApp.getActive(), hw = ss.getSheetByName(SHEETS.HARDWARE);
  var add = function (material, qty) {
    var r = firstEmptyRow_(hw, 1), id = 'HW-' + ST_pad_(getMaxIdNumber_(hw, 'HW') + 1);
    hw.getRange(r, 1).setValue(id); hw.getRange(r, 2).setValue(st.label); hw.getRange(r, 3).setValue(material);
    hw.getRange(r, 5).setValue(qty); hw.getRange(r, 7).setValue('Pcs'); hw.getRange(r, IM_HW_RESERVE_COL).setValue(false);
    return { row: r, id: id };
  };
  if (st.x) st.a = add(st.x.label, 2);
  if (st.z) st.b = add(st.z.label, st.z.qty);
  st.c = add('ZZ selftest free-text widget', 1);
  ST_save_(st);
  SpreadsheetApp.flush();
  var rows = [st.a, st.b, st.c].filter(Boolean).map(function (x) { return x.row; });
  IM_refreshStockStatus_(rows);
  SpreadsheetApp.flush();
  var f = function (o) { return hw.getRange(o.row, IM_STOCK_COL).getValue(); };
  var j = function (o) { return hw.getRange(o.row, IM_HW_STATUS_COL).getValue(); };
  if (st.a) {
    ST_chk_(A, 'In Stock (IMS) column shows a number for an IMS item', typeof f(st.a) === 'number', 'F = ' + f(st.a));
    if (typeof f(st.a) === 'number' && !ST_close_(f(st.a), st.x.stock, 0.001)) ST_log_(A, 'In Stock figure equals the IMS stock read earlier', 'WARN', 'F = ' + f(st.a) + ', earlier ' + st.x.stock + ' (stock moved in between?)');
    ST_chk_(A, 'Stock Status fills itself "In Stock" (needs 2, IMS has ' + st.x.stock + ')', j(st.a) === 'In Stock', 'J = "' + j(st.a) + '"');
  }
  if (st.b) ST_chk_(A, 'Stock Status fills itself "Out of Stock" (needs ' + st.z.qty + ', IMS has ' + st.z.stock + ')', j(st.b) === 'Out of Stock', 'J = "' + j(st.b) + '"');
  ST_chk_(A, 'Free-text material: In Stock shows n/a and Stock Status stays blank', f(st.c) === 'n/a' && j(st.c) === '', 'F = "' + f(st.c) + '", J = "' + j(st.c) + '"');
  ST_chk_(A, 'Nothing reserved or ordered before the tick', rows.every(function (r) { return hw.getRange(r, IM_HW_RSV_COL).getValue() === '' && hw.getRange(r, IM_HW_PUSHED_COL).getValue() === ''; }), '');
  return true;
}

function ST_reserve_(st) {
  var A = 'Reserve', ss = SpreadsheetApp.getActive(), hw = ss.getSheetByName(SHEETS.HARDWARE), pur = ss.getSheetByName(SHEETS.PURCHASES);
  var tick = function (o) { hw.getRange(o.row, IM_HW_RESERVE_COL).setValue(true); SpreadsheetApp.flush(); return IM_autoEvaluateHardwareRow_(hw, o.row); };
  if (st.a) {
    var ra = tick(st.a), ma = String(hw.getRange(st.a.row, IM_HW_RSV_COL).getValue());
    ST_chk_(A, 'Tick on an in-stock row reserves the full quantity', /^In Stock, reserving 2$/.test(ra) && /^PMR-\w+ \| pending \| 2$/.test(ma), ra + ' → ' + ma);
    ST_chk_(A, 'Ticking again does not reserve twice', /already requested/.test(tick(st.a)), '');
    ST_chk_(A, 'In-stock row sends nothing to Purchases', hw.getRange(st.a.row, IM_HW_PUSHED_COL).getValue() === '', '');
  }
  if (st.b) {
    var rb = tick(st.b), pc = String(hw.getRange(st.b.row, IM_HW_PUSHED_COL).getValue());
    var wantRes = st.z.stock > 0 ? ', reserving ' + st.z.stock : '';
    ST_chk_(A, 'Tick on an out-of-stock row reserves what exists and orders the shortfall (5)', rb.indexOf('Out of Stock' + wantRes) === 0 && /^PUR-\d+$/.test(pc), rb);
    if (/^PUR-\d+$/.test(pc)) {
      st.b.pur = pc; ST_save_(st);
      var pr = findRowByValue_(pur, 1, pc), p = pr ? pur.getRange(pr, 1, 1, 15).getValues()[0] : null;
      ST_chk_(A, 'Purchases row: project, material, Required = Ordered = shortfall 5, "Not Ordered"', !!p && p[1] === st.label && p[2] === st.z.label && p[4] === 5 && p[6] === 5 && p[13] === 'Not Ordered', p ? JSON.stringify(p.slice(0, 8)) + ' status ' + p[13] : 'no Purchases row ' + pc);
    }
  }
  var rc = tick(st.c);
  ST_chk_(A, 'Tick on a free-text row is refused ("not an IMS item")', /not an IMS item/.test(rc), rc);
  hw.getRange(st.c.row, IM_HW_RESERVE_COL).setValue(false);
  return true;
}

/** Register rows of this project, as arrays. */
function ST_regRows_(st) {
  var reg = ST_ims_().getSheetByName('Reservation Register');
  if (reg.getLastRow() < 2) return [];
  return reg.getRange(2, 1, reg.getLastRow() - 1, 16).getValues().filter(function (r) { return String(r[7]).indexOf(st.label + ' / ') === 0; });
}

function ST_waitReservations_(st) {
  var A = 'IMS reservation', rows = ST_regRows_(st);
  var want = (st.a ? 1 : 0) + (st.b && st.z.stock > 0 ? 1 : 0);
  if (rows.some(function (r) { return r[8] === 'Error'; })) {
    ST_log_(A, 'IMS turns the PM request into a reservation', 'FAIL', rows.filter(function (r) { return r[8] === 'Error'; }).map(function (r) { return r[9]; }).join('; '));
    return true;
  }
  var active = rows.filter(function (r) { return r[8] === 'Active'; });
  if (active.length < want) { st.waitNote = active.length + ' of ' + want + ' reservation(s) active so far (IMS trigger runs every minute)'; return false; }
  ST_log_(A, 'IMS turns the PM requests into Active reservations within the trigger interval', 'PASS', active.length + ' reservation(s)');
  var ra = st.a && active.filter(function (r) { return r[7] === st.label + ' / ' + st.a.id; });
  if (st.a) {
    var x = ra[0];
    ST_chk_(A, 'One reservation only for the in-stock row (no duplicate), qty 2, right item and UOM', ra.length === 1 && x[3] === st.x.code && x[5] === 2 && String(x[6]) !== '' && String(x[4]) !== '', x ? x[1] + ' ' + x[3] + ' × ' + x[5] + ' ' + x[6] : 'missing');
    st.rsvA = x ? x[1] : ''; ST_save_(st);
  }
  if (st.b && st.z.stock > 0) {
    var rb = active.filter(function (r) { return r[7] === st.label + ' / ' + st.b.id; });
    ST_chk_(A, 'Out-of-stock row reserved only what exists (' + st.z.stock + ')', rb.length === 1 && ST_close_(rb[0][5], st.z.stock), rb[0] ? rb[0][1] + ' × ' + rb[0][5] : 'missing');
  }
  IM_refreshReservationStatus(true);
  var hw = SpreadsheetApp.getActive().getSheetByName(SHEETS.HARDWARE);
  if (st.a) { var m = String(hw.getRange(st.a.row, IM_HW_RSV_COL).getValue()); ST_chk_(A, 'PM column "IMS Reservation" follows IMS: "RSV-… | reserved 2"', /^RSV-\d+ \| reserved 2$/.test(m), m); }
  return true;
}

function ST_waitItemStatus_(st) {
  if (st.skipDownstream) return true;
  var A = 'Item Status', hw = SpreadsheetApp.getActive().getSheetByName(SHEETS.HARDWARE);
  var got = function (o) { return String(hw.getRange(o.row, IM_HW_ITEMSTATUS_COL).getValue()); };
  var okA = !st.a || got(st.a) === 'Reserved';
  var okB = !st.b || /^(Reserved|Not Reserved) \+ To order$/.test(got(st.b));
  if (!(okA && okB)) { st.waitNote = 'A = "' + (st.a ? got(st.a) : '-') + '", B = "' + (st.b ? got(st.b) : '-') + '" (the IMS import formula can lag a few minutes)'; return false; }
  ST_log_(A, 'Item Status: "Reserved" for the in-stock row, "… + To order" for the shortfall row', 'PASS', (st.a ? got(st.a) : '') + ' / ' + (st.b ? got(st.b) : ''));
  return true;
}
function ST_itemStatusTimeout_(st) {
  var hw = SpreadsheetApp.getActive().getSheetByName(SHEETS.HARDWARE);
  ST_log_('Item Status', 'Item Status formulas caught up with IMS', 'WARN', 'still "' + (st.a ? hw.getRange(st.a.row, IM_HW_ITEMSTATUS_COL).getValue() : '-') + '" / "' + (st.b ? hw.getRange(st.b.row, IM_HW_ITEMSTATUS_COL).getValue() : '-') + '" after the wait: _IMS_Items may need "Allow access", or IMPORTRANGE is slow');
}

/** Write requisition commands for IMS (its 1-minute trigger runs them with the real postRequisitionRows). */
function ST_sendCmds_(st, qtys) {
  var ims = ST_ims_(), sh = ims.getSheetByName(ST_CMD);
  if (!sh) { sh = ims.insertSheet(ST_CMD); sh.getRange(1, 1, 1, 5).setValues([['Id', 'Action', 'Params', 'Status', 'Result']]); sh.hideSheet(); }
  var start = Math.max(sh.getLastRow(), 1) + 1, rows = qtys.map(function (q, i) {
    return ['ST-' + (start + i), 'requisition', JSON.stringify({ project: st.label, itemCode: st.x.code, qty: q }), 'Pending', ''];
  });
  sh.getRange(start, 1, rows.length, 5).setValues(rows);
  st.cmdFrom = start; st.cmdTo = start + rows.length - 1; ST_save_(st);
}
function ST_cmdResults_(st) {
  var sh = ST_ims_().getSheetByName(ST_CMD);
  return sh.getRange(st.cmdFrom, 1, st.cmdTo - st.cmdFrom + 1, 5).getValues();
}
/** true when every command has finished; logs errors. */
function ST_cmdsDone_(st, area) {
  var res = ST_cmdResults_(st);
  if (res.some(function (r) { return r[3] === 'Pending'; })) { st.waitNote = 'IMS has not run the requisition command yet (trigger runs every minute)'; return false; }
  var bad = res.filter(function (r) { return r[3] !== 'Done'; });
  ST_chk_(area, 'IMS posted the self-test requisition(s) through the real requisition code', bad.length === 0, bad.length ? bad.map(function (r) { return r[4]; }).join('; ') : res.map(function (r) { return r[4]; }).join(', '));
  return true;
}

function ST_req1_(st) {
  if (st.skipDownstream) { ST_log_('Requisition', 'Requisitions, reservation consumption and the Finance import', 'SKIP', 'IMS did not turn the PM reservation requests into reservations, so nothing after that can be tested'); st.skipReq = true; return true; }
  if (!st.x) { ST_log_('Requisition', 'Requisition and reservation consumption', 'SKIP', 'no usable IMS item'); st.skipReq = true; return true; }
  if (!st.sent1) { ST_sendCmds_(st, [1]); st.sent1 = true; ST_save_(st); st.waitNote = 'command sent to IMS'; return false; }
  if (!ST_cmdsDone_(st, 'Requisition')) return false;
  var A = 'Requisition', r = ST_regRows_(st).filter(function (x) { return x[1] === st.rsvA; })[0];
  ST_chk_(A, 'Issuing 1 of 2 reserved: reservation becomes "Partially Consumed", Consumed Qty 1', !!r && r[8] === 'Partially Consumed' && ST_close_(r[14], 1), r ? r[8] + ', consumed ' + r[14] : 'reservation not found');
  IM_refreshReservationStatus(true);
  var m = String(SpreadsheetApp.getActive().getSheetByName(SHEETS.HARDWARE).getRange(st.a.row, IM_HW_RSV_COL).getValue());
  ST_chk_(A, 'PM column "IMS Reservation" now says "partially consumed"', /partially consumed/.test(m), m);
  return true;
}

function ST_req2_(st) {
  if (st.skipReq) return true;
  if (!st.sent2) { ST_sendCmds_(st, [1, 1]); st.sent2 = true; ST_save_(st); st.waitNote = 'commands sent to IMS'; return false; }
  if (!ST_cmdsDone_(st, 'Requisition')) return false;
  var A = 'Requisition', ims = ST_ims_(), rq = ims.getSheetByName('Material Requisition Register');
  var lines = rq.getLastRow() >= 2 ? rq.getRange(2, 1, rq.getLastRow() - 1, 7).getValues().filter(function (r) { return r[3] === st.label; }) : [];
  ST_chk_(A, 'Three requisition lines exist for the project, exact label, right item, qty 1 each', lines.length === 3 && lines.every(function (r) { return r[4] === st.x.code && r[6] === 1; }), lines.length + ' line(s)');
  var r = ST_regRows_(st).filter(function (x) { return x[1] === st.rsvA; })[0];
  ST_chk_(A, 'After the second issue the reservation is fully "Consumed" (2 of 2); the third, un-reserved issue changed nothing else', !!r && r[8] === 'Consumed' && ST_close_(r[14], 2), r ? r[8] + ', consumed ' + r[14] : 'not found');
  IM_refreshReservationStatus(true);
  var m = String(SpreadsheetApp.getActive().getSheetByName(SHEETS.HARDWARE).getRange(st.a.row, IM_HW_RSV_COL).getValue());
  ST_chk_(A, 'PM column "IMS Reservation" now says "consumed"', /consumed/.test(m) && !/partially/.test(m), m);
  return true;
}

function ST_finance_(st) {
  if (st.skipReq) return true;
  var A = 'Finance import', fin = ST_finBook_(), pc = fin.getSheetByName('Project_Costs');
  var key = 'IMS:' + st.id + ':' + st.x.code + '#1';
  var rows = pc.getLastRow() >= 2 ? pc.getRange(2, 1, pc.getLastRow() - 1, 19).getValues().filter(function (r) { return String(r[12]).indexOf('IMS:' + st.id + ':') === 0; }) : [];
  var g = rows.filter(function (r) { return r[12] === key; })[0];
  if (!g || !ST_close_(g[9], 3)) { st.waitNote = g ? 'cost row has qty ' + g[9] + ' of 3 so far' : 'no cost row yet (Finance import runs every 10 minutes)'; return false; }
  ST_log_(A, 'Finance imported the requisitions within its trigger interval', 'PASS', '');
  ST_chk_(A, 'Three requisitions of one item became ONE grouped cost row (qty 3)', rows.length === 1, rows.length + ' cost row(s) for the project');
  ST_chk_(A, 'Cost row project is the PM/Masters label, type Material, Pending, no GST', g[2] === st.label && g[3] === 'Material' && g[13] === 'Pending' && g[17] === '', 'project "' + g[2] + '", ' + g[3] + ', ' + g[13]);
  var fresh = 0, im = ST_ims_().getSheetByName('Item Master'), head = im.getRange(1, 1, 1, im.getLastColumn()).getValues()[0], c = head.indexOf('Avg Rate (Rs.)');
  if (c >= 0) { var rr = im.getRange(2, 1, im.getLastRow() - 1, im.getLastColumn()).getValues().filter(function (x) { return x[0] === st.x.code; })[0]; fresh = rr ? Number(rr[c]) : 0; }
  ST_chk_(A, 'Cost amount = 3 × the IMS Avg Rate (frozen at import, GST 0)', fresh > 0 && ST_close_(g[7], 3 * fresh, Math.max(0.5, 0.02 * 3 * fresh)), 'amount ' + g[7] + ', 3 × ' + fresh + ' = ' + (3 * fresh));
  var mp = fin.getSheetByName('M_Project'), mrows = mp && mp.getLastRow() >= 2 ? mp.getRange(2, 1, mp.getLastRow() - 1, 3).getValues().filter(function (r) { return String(r[0]).toUpperCase() === st.id; }) : [];
  ST_chk_(A, 'Project is in Finance\'s project list (code + PM name), so invoices / customers can be attached', mrows.length === 1 && mrows[0][1] === st.name, mrows.length ? mrows[0].join(' | ') : 'not in M_Project');
  var lg = fin.getSheetByName('IMS_Cost_Log'), lrows = lg && lg.getLastRow() >= 2 ? lg.getRange(2, 1, lg.getLastRow() - 1, 11).getValues().filter(function (r) { return r[2] === st.id; }) : [];
  ST_chk_(A, 'Import log holds the 3 lines with frozen price and group key', lrows.length === 3 && lrows.every(function (r) { return r[10] === 'Imported' && r[9] === key; }), lrows.length + ' line(s)');
  return true;
}

function ST_cleanup_(st) {
  var A = 'Cleanup';
  if (!st.id || String(st.name).indexOf(ST_TAG) < 0) { ST_log_(A, 'Delete the test project', 'SKIP', 'no test project was created'); return true; }
  var ss = SpreadsheetApp.getActive();
  var pv = Masters.previewProjectDelete(st.id, ss.getId(), st.label);
  if (pv.problems.length) { ST_log_(A, 'Delete preview', 'FAIL', pv.problems.join(' / ') + '. NOT deleted: remove ' + st.label + ' with Delete Project.'); st.noClean = true; return true; }
  var total = function (list) { return list.reduce(function (n, x) { return n + (x.n || 0); }, 0); };
  var n = { pm: total(pv.systems.pm), ims: total(pv.systems.ims), fin: total(pv.systems.fin) };
  var summary = 'PM ' + n.pm + ', IMS ' + n.ims + ', Finance ' + n.fin + ' row(s)';
  if (n.pm > 12 || n.ims > 12 || n.fin > 12) {
    ST_log_(A, 'Delete preview is small enough to be only the test data', 'FAIL', summary + ': far more than a self-test creates. NOT deleted. Check ' + st.label + ' by hand.'); st.noClean = true; return true;
  }
  ST_log_(A, 'Delete preview', 'INFO', summary);
  var res = Masters.executeProjectDelete(st.id, ss.getId(), st.label, st.id);
  reseedArrayFormulas_(ss);
  ST_chk_(A, 'Project deleted everywhere (Finance, IMS, PM, Masters)', res.complete && !res.problems.length, res.problems.join('; ') || 'request ' + res.requestId);
  if (!res.complete) st.noClean = true;
  return true;
}

function ST_verifyCleanup_(st) {
  if (st.noClean || !st.id) return true;
  var A = 'Cleanup', left = [];
  var ss = SpreadsheetApp.getActive();
  if (findRowByValue_(ss.getSheetByName(SHEETS.PROJECTS), 1, st.id)) left.push('PM Projects');
  var hw = ss.getSheetByName(SHEETS.HARDWARE);
  if (hw.getLastRow() >= 2 && hw.getRange(2, 2, hw.getLastRow() - 1, 1).getValues().some(function (r) { return r[0] === st.label; })) left.push('PM Hardware List');
  if ((Masters.getProjects() || []).some(function (p) { return String(p.Project_ID).toUpperCase() === st.id; })) left.push('Masters Projects');
  var ims = ST_ims_();
  if (ST_regRows_(st).length) left.push('IMS Reservation Register');
  var rq = ims.getSheetByName('Material Requisition Register');
  if (rq.getLastRow() >= 2 && rq.getRange(2, 4, rq.getLastRow() - 1, 1).getValues().some(function (r) { return r[0] === st.label; })) left.push('IMS Material Requisition Register');
  try {
    var fin = ST_finBook_(), pc = fin.getSheetByName('Project_Costs');
    if (pc.getLastRow() >= 2 && pc.getRange(2, 13, pc.getLastRow() - 1, 1).getValues().some(function (r) { return String(r[0]).indexOf('IMS:' + st.id + ':') === 0; })) left.push('Finance Project_Costs');
    var mp = fin.getSheetByName('M_Project');
    if (mp && mp.getLastRow() >= 2 && mp.getRange(2, 1, mp.getLastRow() - 1, 1).getValues().some(function (r) { return String(r[0]).toUpperCase() === st.id; })) left.push('Finance project list');
  } catch (e) { left.push('Finance (' + e.message + ')'); }
  ST_chk_(A, 'Nothing of the test project is left in any system', left.length === 0, left.length ? 'still present in: ' + left.join(', ') : '');
  return true;
}

function ST_verifyStock_(st) {
  if (st.noClean || !st.x) return true;
  var it = SpreadsheetApp.getActive().getSheetByName(IM_ITEMS_SHEET);
  var r = it.getRange(2, 1, Math.max(it.getLastRow() - 1, 1), 5).getValues().filter(function (x) { return x[0] === st.x.code; })[0];
  if (!r || !ST_close_(r[4], st.x.stock, 0.001)) { st.waitNote = 'IMS stock of ' + st.x.code + ' is ' + (r ? r[4] : '?') + ', was ' + st.x.stock + ' (snapshots refresh within a minute, the import formula a few more)'; return false; }
  ST_log_('Cleanup', 'IMS stock of ' + st.x.code + ' is back to ' + st.x.stock + ' (requisitions and reservations are gone)', 'PASS', '');
  return true;
}
function ST_stockTimeout_(st) { ST_log_('Cleanup', 'IMS stock of ' + st.x.code + ' back to ' + st.x.stock, 'WARN', st.waitNote + '. Open IMS and check Stock Snapshot for that item.'); }

function ST_final_(st) {
  try { var ims = ST_ims_(), c = ims.getSheetByName(ST_CMD); if (c) ims.deleteSheet(c); } catch (e) {}
  st.done = true;
  return true;
}

var ST_STEPS_ = [
  { n: 'preflight', f: ST_preflight_, min: 2 }, { n: 'project', f: ST_project_, min: 2 }, { n: 'items', f: ST_items_, min: 2 },
  { n: 'hardware', f: ST_hardware_, min: 2 }, { n: 'reserve', f: ST_reserve_, min: 2 },
  { n: 'reservation in IMS', f: ST_waitReservations_, min: 5, fail: function (st) { st.skipDownstream = true; st.skipReq = true; } },
  { n: 'item status', f: ST_waitItemStatus_, min: 6, soft: ST_itemStatusTimeout_ },
  { n: 'requisition 1', f: ST_req1_, min: 5, fail: function (st) { st.skipReq = true; } }, { n: 'requisition 2', f: ST_req2_, min: 5, fail: function (st) { st.skipReq = true; } },
  { n: 'finance import', f: ST_finance_, min: 16 },
  { n: 'cleanup', f: ST_cleanup_, min: 3 }, { n: 'verify cleanup', f: ST_verifyCleanup_, min: 3 },
  { n: 'verify stock', f: ST_verifyStock_, min: 5, soft: ST_stockTimeout_ }, { n: 'finish', f: ST_final_, min: 2 }
];

/** Trigger target (every minute while a test runs). */
function IM_selfTestTick() {
  IM_QUIET_ = true;   // the verify routines used by preflight must not pop up dialogs
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var st = ST_load_();
    if (!st || st.done) { ST_stopTrigger_(); return; }
    var t0 = Date.now();
    while (!st.done && Date.now() - t0 < 240000) {
      var def = ST_STEPS_[st.step], res;
      if (!def) { st.done = true; break; }
      try { res = def.f(st); }
      catch (e) { ST_log_('Error', 'Step "' + def.n + '"', 'FAIL', e.message + (e.stack ? ' @ ' + String(e.stack).split('\n')[1] : '')); st.failed = (st.failed || []).concat(def.n); res = true; }
      if (res === true) { st.step++; st.stepAt = Date.now(); st.waitNote = ''; ST_save_(st); continue; }
      if (Date.now() - st.stepAt > def.min * 60000) {
        if (def.soft) def.soft(st);
        else { ST_log_('Timeout', 'Step "' + def.n + '"', 'FAIL', 'gave up after ' + def.min + ' min: ' + (st.waitNote || 'no progress')); if (def.fail) def.fail(st); }
        st.step++; st.stepAt = Date.now(); st.waitNote = ''; ST_save_(st); continue;
      }
      ST_save_(st); break;
    }
    if (st.done) { ST_save_(st); ST_finish_(st); }
  } finally { lock.releaseLock(); }
}

function ST_stopTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === ST_FN) ScriptApp.deleteTrigger(t); });
}

function ST_finish_(st) {
  ST_stopTrigger_();
  var sh = ST_reportSheet_(false), last = sh.getLastRow();
  var vals = last >= 4 ? sh.getRange(4, 4, last - 3, 1).getValues().map(function (r) { return r[0]; }) : [];
  var c = function (k) { return vals.filter(function (v) { return v === k; }).length; };
  var verdict = c('FAIL') ? 'PROBLEMS FOUND' : (c('WARN') ? 'PASSED with warnings' : 'ALL GOOD');
  var line = verdict + ' · ' + c('PASS') + ' passed, ' + c('FAIL') + ' failed, ' + c('WARN') + ' warning(s) · finished ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm') + (st.noClean ? ' · TEST PROJECT NOT DELETED: ' + st.label : '');
  sh.getRange(1, 1).setValue('Self-test: ' + line);
  sh.getRange(1, 1).setBackground(c('FAIL') ? '#FFC7CE' : c('WARN') ? '#FFEB9C' : '#C6EFCE');
  PropertiesService.getScriptProperties().setProperty('IM_SELFTEST_MSG', line);
  PropertiesService.getScriptProperties().deleteProperty(ST_PROP);
}

/** Menu: start. */
function IM_selfTestStart() {
  var ui = SpreadsheetApp.getUi();
  var cur = ST_load_();
  if (cur && !cur.done) { ui.alert('Self-test', 'A self-test is already running (step ' + (cur.step + 1) + ' of ' + ST_STEPS_.length + '). Watch the "' + ST_REPORT + '" sheet, or use "Stop self-test and clean up".', ui.ButtonSet.OK); return; }
  var ok = ui.alert('Run the system self-test?',
    'It creates a test project "' + ST_TAG + ' …" with a few Hardware List rows, reserves and issues about 3 units of one IMS item, waits for the Finance import, then DELETES the test project everywhere.\n\n' +
    '• Takes about 15 minutes; you can keep working, but do not edit the test rows.\n• Burns a few REQ / RSV / COST numbers.\n• Results appear in the "' + ST_REPORT + '" sheet.\n\nStart now?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  ST_reportSheet_(true);
  var st = { v: 1, step: 0, stepAt: Date.now(), startedAt: Date.now(), waitNote: '' };
  ST_save_(st);
  ST_stopTrigger_();
  ScriptApp.newTrigger(ST_FN).timeBased().everyMinutes(1).create();
  SpreadsheetApp.getActive().setActiveSheet(ST_reportSheet_(false));
  SpreadsheetApp.getActive().toast('Self-test started. It runs by itself; results appear in the report sheet.', '🧪 Self-test', 8);
  SpreadsheetApp.getActive().toast('The first checks run within a minute.', '🧪 Self-test', 5);
}

/** Menu: abort the waiting and clean up now. */
function IM_selfTestStop() {
  var st = ST_load_();
  if (!st) { SpreadsheetApp.getUi().alert('Self-test', 'No self-test is running.', SpreadsheetApp.getUi().ButtonSet.OK); return; }
  ST_log_('Stop', 'Self-test stopped by the user', 'WARN', 'jumping to cleanup');
  var i = 0; while (ST_STEPS_[i] && ST_STEPS_[i].n !== 'cleanup') i++;
  st.step = i; st.stepAt = Date.now(); ST_save_(st);
  IM_selfTestTick();
}

/** Menu: open the last report. */
function IM_selfTestShowReport() {
  var sh = SpreadsheetApp.getActive().getSheetByName(ST_REPORT);
  if (!sh) { SpreadsheetApp.getUi().alert('Self-test', 'No report yet. Run the self-test first.', SpreadsheetApp.getUi().ButtonSet.OK); return; }
  SpreadsheetApp.getActive().setActiveSheet(sh);
}

/** On open: tell the user how the last self-test ended (once). */
function IM_selfTestAnnounce_() {
  var props = PropertiesService.getScriptProperties(), msg = props.getProperty('IM_SELFTEST_MSG');
  if (!msg) return;
  SpreadsheetApp.getActive().toast(msg + ' (see the "' + ST_REPORT + '" sheet)', '🧪 Self-test', 20);
  props.deleteProperty('IM_SELFTEST_MSG');
}

// =============================================================================
// ➕ ADD ITEM TO IMS  (🔗 Interlink → Add item to IMS)
// Creates a new Item Master row in IMS from PM, without opening IMS. PM cannot run IMS code, so the dialog leaves
// a command in IMS's hidden sheet _PM_Commands and IMS's 1-minute trigger runs the REAL addItem() there (same
// checks, duplicate detection and item-code numbering as the IMS form). The dialog waits for the answer and
// shows the new code. The item shows up in the Material dropdown after the hidden _IMS_Items list refreshes
// (a few minutes). Opening stock is not set here: it arrives with a GRN or an opening adjustment in IMS.
// =============================================================================
var IM_CMD_SHEET = '_PM_Commands';

/** The three lists IMS validates against (Dropdowns A = category, C = material type, E = UOM), read from an open IMS book. */
function IM_addItemListsFrom_(ims) {
  var dd = ims.getSheetByName('Dropdowns');
  if (!dd) throw new Error('IMS has no Dropdowns sheet');
  var col = function (c) {
    if (dd.getLastRow() < 2) return [];
    return dd.getRange(2, c, dd.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }).filter(String);
  };
  return { categories: col(1), materialTypes: col(3), uoms: col(5) };
}
function IM_addItemLists() { return IM_addItemListsFrom_(SpreadsheetApp.openById(IM_IMS_SHEET_ID)); }

/** Validate and queue the item. d = { itemName, size, grade, uom, category, materialType, avgRate, forceCreate }. Returns { row }. */
function IM_addItemSubmit(d) {
  var ims = SpreadsheetApp.openById(IM_IMS_SHEET_ID), lists = IM_addItemListsFrom_(ims);   // IMS is opened once
  var name = String(d.itemName || '').trim();
  if (!name) throw new Error('Item name is required.');
  if (name.length > 120) throw new Error('Item name is too long (max 120 characters).');
  if (lists.uoms.indexOf(d.uom) < 0) throw new Error('Unit "' + d.uom + '" is not in the IMS list.');
  if (lists.categories.indexOf(d.category) < 0) throw new Error('Category "' + d.category + '" is not in the IMS list.');
  if (lists.materialTypes.indexOf(d.materialType) < 0) throw new Error('Material type "' + d.materialType + '" is not in the IMS list.');
  var rate = d.avgRate === '' || d.avgRate == null ? 0 : Number(d.avgRate);
  if (isNaN(rate) || rate < 0) throw new Error('Avg rate must be a number, 0 or more.');
  var params = { itemName: name, size: String(d.size || '').trim(), grade: String(d.grade || '').trim(), uom: d.uom,
                 category: d.category, materialType: d.materialType, avgRate: rate, forceCreate: d.forceCreate === true };
  var sh = ims.getSheetByName(IM_CMD_SHEET);
  if (!sh) {
    sh = ims.insertSheet(IM_CMD_SHEET);
    sh.getRange(1, 1, 1, 7).setValues([['Id', 'Action', 'Params', 'Status', 'Result', 'By', 'At']]);
    sh.hideSheet();
  }
  var row = Math.max(sh.getLastRow(), 1) + 1;
  sh.getRange(row, 1, 1, 7).setValues([['PMC-' + row, 'addItem', JSON.stringify(params), 'Pending', '',
    Session.getActiveUser().getEmail() || 'unknown', new Date()]]);
  return { row: row };
}

/**
 * What happened to a queued item: { status: Pending | Done | Error, result: {success, message, code, duplicate, existingCode, existingDesc} }.
 * A read that Google answers with a timeout / "server error" is NOT a failure of the item: it is reported as Pending
 * so the dialog simply asks again. Anything else is a real error and is raised.
 */
function IM_addItemPoll(row) {
  try {
    var sh = SpreadsheetApp.openById(IM_IMS_SHEET_ID).getSheetByName(IM_CMD_SHEET);
    if (!sh) throw new Error('IMS command sheet not found');
    var r = sh.getRange(Number(row), 1, 1, 5).getValues()[0], result = null;
    if (r[3] !== 'Pending' && r[4]) { try { result = JSON.parse(r[4]); } catch (e) { result = { success: false, message: String(r[4]) }; } }
    return { status: String(r[3]), result: result };
  } catch (e) {
    if (/DEADLINE|timed out|server error|try again|temporar|Service/i.test(String(e && e.message))) return { status: 'Pending', result: null, slow: true };
    throw e;
  }
}

function IM_addItemDialog() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;font-size:13px;padding:14px">' +
    '<div id="form"><p style="margin-top:0"><b>Add a new item to IMS</b><br><span style="color:#555">It is created in IMS Item Master and shows in the Material dropdown in a few minutes.</span></p>' +
    '<label>Item name *</label><input id="name" style="width:100%;padding:6px;margin:3px 0 8px" maxlength="120">' +
    '<table style="width:100%"><tr><td style="width:50%;padding-right:6px"><label>Size</label><input id="size" style="width:100%;padding:6px;margin:3px 0 8px"></td>' +
    '<td><label>Grade / colour</label><input id="grade" style="width:100%;padding:6px;margin:3px 0 8px"></td></tr>' +
    '<tr><td style="padding-right:6px"><label>Unit *</label><select id="uom" style="width:100%;padding:6px;margin:3px 0 8px"></select></td>' +
    '<td><label>Avg rate (optional)</label><input id="rate" type="number" min="0" step="any" style="width:100%;padding:6px;margin:3px 0 8px"></td></tr></table>' +
    '<label>Category *</label><select id="cat" style="width:100%;padding:6px;margin:3px 0 8px"></select>' +
    '<label>Material type *</label><select id="mtype" style="width:100%;padding:6px;margin:3px 0 8px"></select>' +
    '<p><button id="go" onclick="send(false)" style="padding:7px 18px;background:#1F4E79;color:#fff;border:0">Add to IMS</button> ' +
    '<button onclick="google.script.host.close()" style="padding:7px 18px">Cancel</button></p></div>' +
    '<div id="wait" style="display:none"><p id="waittxt"></p></div>' +
    '<div id="done" style="display:none"><p id="donetxt" style="white-space:pre-wrap"></p><p id="dupbtn" style="display:none">' +
    '<button onclick="send(true)" style="padding:7px 18px">Create anyway</button> </p>' +
    '<p><button onclick="google.script.host.close()" style="padding:7px 18px">Close</button></p></div>' +
    '<p id="err" style="color:#b00020"></p></div>' +
    '<script>' +
    'var fails=0;' +
    'function el(i){return document.getElementById(i);}' +
    'function fail(e){el("err").textContent=(e&&e.message)?e.message:String(e);el("go").disabled=false;}' +
    'function transient(e){return /DEADLINE|server error|timed out|try again|temporar|storage|Service/i.test((e&&e.message)?e.message:String(e));}' +
    'function fill(id,list){var s=el(id);s.innerHTML="";var o=document.createElement("option");o.value="";o.textContent="-- select --";s.appendChild(o);' +
    'list.forEach(function(x){var q=document.createElement("option");q.value=x;q.textContent=x;s.appendChild(q);});}' +
    'google.script.run.withSuccessHandler(function(l){fill("uom",l.uoms);fill("cat",l.categories);fill("mtype",l.materialTypes);}).withFailureHandler(fail).IM_addItemLists();' +
    'function send(force){el("err").textContent="";fails=0;var d={itemName:el("name").value,size:el("size").value,grade:el("grade").value,uom:el("uom").value,category:el("cat").value,materialType:el("mtype").value,avgRate:el("rate").value,forceCreate:force};' +
    'if(!d.itemName.trim()||!d.uom||!d.category||!d.materialType){fail("Fill the item name, unit, category and material type.");return;}' +
    'el("go").disabled=true;google.script.run.withSuccessHandler(function(r){wait(r.row,0);}).withFailureHandler(fail).IM_addItemSubmit(d);}' +
    'function wait(row,n){el("form").style.display="none";el("done").style.display="none";el("wait").style.display="block";' +
    'el("waittxt").textContent="Sent to IMS, waiting for it to be created ("+(n*6)+"s). IMS picks it up within a minute."+(fails?" Google is a little slow, still trying.":"");' +
    'if(n>25){finish("IMS has not answered yet. Your request was sent: look in IMS Item Master before sending it again. If it is not there, check that IMS is running its 1-minute job (IMS: Interlink > Process PM reservation requests now).",false);return;}' +
    'setTimeout(function(){google.script.run.withSuccessHandler(function(p){if(p.status==="Pending"){wait(row,n+1);return;}fails=0;' +
    'var r=p.result||{};if(r.success){finish("Added "+r.code+"\\n"+r.message+"\\n\\nIt appears in the Material dropdown within a few minutes.",false);}' +
    'else if(r.duplicate){finish(r.message+"\\n\\nIf this really is a different item, create it anyway.",true);}else{finish("Not added: "+(r.message||"unknown error"),false);}})' +
    '.withFailureHandler(function(e){if(transient(e)&&fails<10){fails++;wait(row,n+1);return;}' +
    'finish("Your request was sent to IMS, but this window could not read the answer ("+((e&&e.message)?e.message:e)+"). Look in IMS Item Master before sending it again: the item may already be there.",false);}).IM_addItemPoll(row);},6000);}' +
    'function finish(t,dup){el("wait").style.display="none";el("done").style.display="block";el("donetxt").textContent=t;el("dupbtn").style.display=dup?"block":"none";}' +
    '</script>'
  ).setWidth(470).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, '➕ Add item to IMS');
}

// =============================================================================
// FRIENDLY HELP · HEALTH CHECK · NEEDS ATTENTION  (for people who are not technical)
//   IM_friendlyError_        plain-language version of an error, with what to do next
//   IM_healthChecks_         the fast checks (connections, background jobs) shared by the 🩺 Health check button,
//                            📋 Needs attention and the long self-test
//   IM_needsAttentionNow     the "Needs attention" sheet: what is stuck or waiting for a person, and what to do
//   IM_installHelpNotes_     explanations on the Hardware List column headers (hover the header)
//   IM_healFormulaColumns_   puts back the automatic formulas (In Stock, Item Status) if someone typed over them
//   IM_showStartHere         the "Start here" tab
// =============================================================================

/** Plain-language form of an error / skip message: { text, hint }. Unknown messages come back unchanged with no hint. */
function IM_friendlyError_(msg) {
  var m = String(msg == null ? '' : msg).replace(/^(skipped|error):\s*/i, '').replace(/^Exception:\s*/i, '').trim(), x;
  var rules = [
    [/needs Project, Material and Required Qty/i, function () { return ['Fill in the Project, Material and Required Qty first.', 'Then tick Reserve? again.']; }],
    [/not an IMS item/i, function () { return ['This material is not in IMS, so it cannot be reserved.', 'Add it with 🔗 Interlink → ➕ Add item to IMS and pick it from the dropdown in a few minutes, or set Stock Status to Out of Stock to order it.']; }],
    [/already sent to Purchases \(([^)]*)\)/i, function (r) { return ['This row was already sent to Purchases (' + r[1] + ').', 'Nothing more to do here.']; }],
    [/reservation already requested \(([^)]*)\)/i, function (r) { return ['IMS already has the reservation request for this row (' + r[1] + ').', 'Nothing more to do; IMS confirms within about a minute.']; }],
    [/sheet busy|Busy|lock/i, function () { return ['Another update was running at the same moment.', 'Wait a few seconds and tick again.']; }],
    [/Code not found: (\S+)/i, function (r) { return ['IMS does not know the item "' + r[1] + '".', 'Pick the Material again from the dropdown, or add it with 🔗 Interlink → ➕ Add item to IMS.']; }],
    [/Qty must be > ?0/i, function () { return ['The quantity must be more than 0.', 'Enter the Required Qty and tick Reserve? again.']; }],
    [/not in allowed list/i, function () { return ['That value is not in IMS\'s list.', 'Choose it from the dropdown.']; }],
    [/Editors only|Admin only|not an editor/i, function () { return ['This account is not allowed to do that.', 'Ask the Masters admin to add your email as an editor.']; }],
    [/no access|Cannot open|You do not have permission|Access denied|Requested entity was not found/i, function () { return ['This account cannot open the other spreadsheet.', 'Ask the owner to share IMS / Finance with you as an editor.']; }],
    [/too many times|quota|Rate Limit/i, function () { return ['Google is slowing the sheet down for a moment.', 'Wait a minute and try again.']; }],
    [/IMS Reservation Register is not set up/i, function () { return ['IMS is not ready for PM requests yet.', 'Ask your administrator to run the IMS set-up once.']; }]
  ];
  for (var i = 0; i < rules.length; i++) {
    if ((x = rules[i][0].exec(m))) { var o = rules[i][1](x); return { text: o[0], hint: o[1] }; }
  }
  return { text: m, hint: '' };
}
/** One sentence: what happened and what to do. */
function IM_friendlyText_(msg) { var f = IM_friendlyError_(msg); return f.text + (f.hint ? ' ' + f.hint : ''); }

// ---------------------------------------------------------------- health checks
/** Minutes since the "dd/MM HH:mm[:ss]" in a heartbeat text (IST), or -1 when unreadable. */
function IM_heartbeatAge_(text) {
  var m = /(\d\d)\/(\d\d) (\d\d):(\d\d)(?::(\d\d))?/.exec(String(text));
  if (!m) return -1;
  var yr = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy');
  var when = Utilities.parseDate(yr + '-' + m[2] + '-' + m[1] + ' ' + m[3] + ':' + m[4] + ':' + (m[5] || '00'), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss');
  return (Date.now() - when.getTime()) / 60000;
}

/**
 * The fast checks. Returns [{ area, check, result: PASS|FAIL|WARN, detail, fix }]. `fix` is a plain sentence for a
 * person. A duplicated PM trigger is repaired on the spot (reported as WARN). Never throws.
 */
function IM_healthChecks_() {
  var out = [], A = 'Setup', ss = SpreadsheetApp.getActive();
  var add = function (check, ok, detail, fix, level) { out.push({ area: A, check: check, result: ok ? 'PASS' : (level || 'FAIL'), detail: detail || '', fix: ok ? '' : (fix || '') }); return ok; };
  try { add('Masters library reachable', true, 'version ' + Masters.getVersion() + ', ' + (Masters.getProjects() || []).length + ' project(s)'); }
  catch (e) { add('Masters library reachable', false, e.message, 'Masters cannot be reached. ' + IM_friendlyText_(e.message)); }

  try {
    var ims = SpreadsheetApp.openById(IM_IMS_SHEET_ID), reg = ims.getSheetByName('Reservation Register');
    add('IMS: Reservation Register has the Request ID column (PM requests can be queued)', !!reg && reg.getRange(1, 16).getValue() === 'Request ID', '',
        'IMS is not set up for PM requests. Ask your administrator to run the IMS set-up (IM_installReservationRegister).');
    add('IMS: Material Requisition Register and Item Master exist', !!ims.getSheetByName('Material Requisition Register') && !!ims.getSheetByName('Item Master'), '', 'An IMS sheet is missing or renamed. Tell your administrator.');
    var hb = reg ? String(reg.getRange('Q1').getValue()) : '', age = IM_heartbeatAge_(hb), tr = /triggers: (\d+)/.exec(hb);
    add('IMS: 1-minute reservation trigger is running (heartbeat < 5 min old)', age >= 0 && age < 5 && !!tr && Number(tr[1]) >= 1, hb || 'no heartbeat text',
        'IMS has stopped answering PM reservation requests. Open the IMS spreadsheet once (that restarts it), or in IMS use 🔗 Interlink → Process PM reservation requests now.');
  } catch (e) { add('IMS reachable', false, e.message, 'PM cannot open IMS. ' + IM_friendlyText_(e.message)); }

  try {
    var ids = Masters.getSystemSheetIds();
    add('Finance and IMS have registered their spreadsheet ids with Masters', !!ids.finance && !!ids.ims, 'finance: ' + (ids.finance ? 'yes' : 'NO'), 'Open the Finance spreadsheet once so it registers itself.');
    if (ids.finance) {
      var log = SpreadsheetApp.openById(ids.finance).getSheetByName('IMS_Cost_Log');
      var fhb = log ? String(log.getRange('M1').getValue()) : '', fage = IM_heartbeatAge_(fhb), ftr = /triggers: (\d+)/.exec(fhb);
      add('Finance: 10-minute IMS cost import is running (last run < 25 min ago)', fage >= 0 && fage < 25 && !!ftr && Number(ftr[1]) >= 1, fhb || 'no IMS_Cost_Log heartbeat',
          'The Finance import of IMS costs has stopped. Open the Finance spreadsheet once (that restarts it). If this stays red, ask your administrator to run IM_installImportTrigger in Finance.');
    }
  } catch (e) { add('Finance reachable', false, e.message, 'PM cannot open Finance. ' + IM_friendlyText_(e.message)); }

  var dupes = 0;
  try { dupes = IM_removeDuplicateTriggers_(); } catch (e) {}
  if (dupes) out.push({ area: A, check: 'No duplicate PM triggers', result: 'WARN', detail: dupes + ' duplicate trigger(s) were installed (every edit ran twice); the extra copy was removed', fix: '' });
  else add('No duplicate PM triggers (each edit runs once)', true, '');
  var handlers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  add('PM: edit triggers installed (onEditHandler, IM_onHardwareItemEdit_)', handlers.indexOf('onEditHandler') >= 0 && handlers.indexOf('IM_onHardwareItemEdit_') >= 0, handlers.join(', '),
      'PM\'s automatic edit handling is not installed. Ask your administrator to run IM_installItems and installTriggers.');
  var it = ss.getSheetByName(IM_ITEMS_SHEET), a2 = it ? String(it.getRange('A2').getDisplayValue()) : '';
  add('PM: IMS item list (_IMS_Items) loads, no #REF / "Allow access" error', !!it && a2 !== '' && a2.charAt(0) !== '#', a2 ? 'first item: ' + a2 : 'empty',
      'The IMS item list is not loading. Open View ▸ Hidden sheets ▸ _IMS_Items, click cell A2 and press "Allow access".');
  return out;
}

/** Menu: 🩺 Health check. One verdict, each problem with what to do. */
function IM_healthCheckNow() {
  var ui = SpreadsheetApp.getUi(), res = IM_healthChecks_();
  var bad = res.filter(function (r) { return r.result === 'FAIL'; }), warn = res.filter(function (r) { return r.result === 'WARN'; });
  var head = bad.length ? '🔴 ' + bad.length + ' problem(s) found' : '🟢 Everything is working';
  var lines = bad.map(function (r) { return '• ' + r.check.replace(/\s*\(.*$/, '') + '\n   → ' + (r.fix || r.detail); });
  warn.forEach(function (r) { lines.push('• ' + r.check + ': ' + r.detail); });
  ui.alert('🩺 Health check', head + '\n\n' + (lines.length ? lines.join('\n\n') + '\n\n' : '') + (res.length - bad.length) + ' of ' + res.length + ' checks are fine.', ui.ButtonSet.OK);
}

// ---------------------------------------------------------------- needs attention
var NA_SHEET = 'Needs attention';

/** Rebuild the "Needs attention" sheet. quiet = on open: no pop-up, only a toast when something is red. Returns { red, amber, blue }. */
function IM_needsAttentionNow(quiet) {
  var ss = SpreadsheetApp.getActive(), lines = [];
  var add = function (p, area, what, where, todo) { lines.push([p, area, what, where, todo]); };
  var capped = function (arr, n, makeLine, moreLine) { arr.slice(0, n).forEach(makeLine); if (arr.length > n) moreLine(arr.length - n); };

  // 1. background jobs / connections (the health checks)
  IM_healthChecks_().forEach(function (r) {
    if (r.result === 'FAIL') add('🔴', 'System', r.check.replace(/\s*\(.*$/, ''), '', r.fix || r.detail);
    else if (r.result === 'WARN') add('🔵', 'System', r.check, '', r.detail);
  });

  // 2. Hardware List
  var hw = ss.getSheetByName(SHEETS.HARDWARE), errs = [], oos = [], free = [], ready = [];
  if (hw && hw.getLastRow() >= 2) {
    hw.getRange(2, 1, hw.getLastRow() - 1, IM_HW_RSV_COL).getValues().forEach(function (r, i) {
      var id = r[0], row = i + 2;
      if (!id || !r[1] || !r[2]) return;
      var ticked = r[IM_HW_RESERVE_COL - 1] === true, pushed = r[IM_HW_PUSHED_COL - 1], m = String(r[IM_HW_RSV_COL - 1] || ''), st = r[IM_HW_STATUS_COL - 1], f = r[IM_STOCK_COL - 1];
      if (m.charAt(0) === '⚠') errs.push({ id: id, row: row, why: m.replace(/^⚠\s*/, '') });
      else if (!ticked && !pushed && !m) {
        if (f === 'n/a') free.push({ id: id, row: row, name: r[2] });
        else if (st === 'Out of Stock') oos.push({ id: id, row: row, name: r[2] });
        else if (st === 'In Stock') ready.push({ id: id, row: row, name: r[2] });
      }
    });
  }
  capped(errs, 30, function (e) { add('🔴', 'Hardware List', e.id + ': IMS could not reserve this row. ' + IM_friendlyError_(e.why).text, 'Hardware List, row ' + e.row, IM_friendlyError_(e.why).hint || 'Fix it, untick and tick Reserve? again.'); },
    function (n) { add('🔴', 'Hardware List', '+ ' + n + ' more rows with a ⚠', 'Hardware List', 'Filter the IMS Reservation column for ⚠.'); });
  capped(free, 30, function (e) { add('🟡', 'Hardware List', e.id + ': "' + e.name + '" is not an IMS item', 'Hardware List, row ' + e.row, 'Add it with 🔗 Interlink → ➕ Add item to IMS and pick it again, or set Stock Status to Out of Stock to order it.'); },
    function (n) { add('🟡', 'Hardware List', '+ ' + n + ' more free-text items', 'Hardware List', ''); });
  capped(oos, 30, function (e) { add('🟡', 'Hardware List', e.id + ': "' + e.name + '" is out of stock and not ordered yet', 'Hardware List, row ' + e.row, 'Tick Reserve? to reserve what exists and send the shortfall to Purchases.'); },
    function (n) { add('🟡', 'Hardware List', '+ ' + n + ' more out-of-stock rows', 'Hardware List', ''); });
  capped(ready, 30, function (e) { add('🔵', 'Hardware List', e.id + ': "' + e.name + '" is in stock and not reserved yet', 'Hardware List, row ' + e.row, 'Tick Reserve? if this project should hold it.'); },
    function (n) { add('🔵', 'Hardware List', '+ ' + n + ' more rows ready to reserve', 'Hardware List', ''); });

  // 3. IMS: PM requests waiting too long, requisitions Finance cannot cost
  try {
    var ims = SpreadsheetApp.openById(IM_IMS_SHEET_ID), reg = ims.getSheetByName('Reservation Register');
    if (reg && reg.getLastRow() >= 2) {
      var stuck = reg.getRange(2, 1, reg.getLastRow() - 1, 16).getValues().filter(function (r) { return r[8] === 'Requested' && r[11] instanceof Date && (Date.now() - r[11].getTime()) > 10 * 60000; });
      if (stuck.length) add('🔴', 'IMS', stuck.length + ' reservation request(s) from PM have waited more than 10 minutes', 'IMS Reservation Register (Status "Requested")', 'Open the IMS spreadsheet once, or use IMS → 🔗 Interlink → Process PM reservation requests now.');
    }
    var rq = ims.getSheetByName('Material Requisition Register'), im = ims.getSheetByName('Item Master');
    if (rq && im && rq.getLastRow() >= 2 && im.getLastRow() >= 2) {
      var head = im.getRange(1, 1, 1, im.getLastColumn()).getValues()[0], cRate = head.indexOf('Avg Rate (Rs.)'), cName = head.indexOf('Item Name'), rate = {}, name = {};
      im.getRange(2, 1, im.getLastRow() - 1, im.getLastColumn()).getValues().forEach(function (r) { if (r[0]) { rate[r[0]] = Number(r[cRate]) || 0; name[r[0]] = r[cName] || r[2] || r[0]; } });
      var miss = {};
      rq.getRange(2, 1, rq.getLastRow() - 1, 7).getValues().forEach(function (r) {
        if (/^\s*PRJ-\d+/i.test(String(r[3])) && r[4] && Number(r[6]) > 0 && !(rate[r[4]] > 0)) miss[r[4]] = (miss[r[4]] || 0) + 1;
      });
      Object.keys(miss).slice(0, 30).forEach(function (code) {
        add('🟡', 'IMS → Finance', 'Finance cannot cost ' + miss[code] + ' requisition line(s) of ' + code + ' (' + (name[code] || '') + '): it has no Avg Rate', 'IMS Item Master', 'Enter the Avg Rate for ' + code + ' in IMS Item Master. Finance picks the lines up by itself afterwards.');
      });
    }
  } catch (e) { add('🔴', 'IMS', 'PM could not read IMS', '', IM_friendlyText_(e.message)); }

  // 4. PM projects that Masters does not know (IMS and Finance cannot see them)
  try {
    var mine = ss.getSheetByName(SHEETS.PROJECTS), known = {};
    (Masters.getProjects() || []).forEach(function (p) { known[String(p.Project_ID).toUpperCase()] = true; });
    if (mine && mine.getLastRow() >= 2) mine.getRange(2, 1, mine.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (r[0] && r[1] && !known[String(r[0]).toUpperCase()]) add('🟡', 'Projects', r[0] + ' — ' + r[1] + ' is not in Masters yet', 'Projects sheet', 'It syncs by itself when PM opens. If it stays, ask your administrator (editor rights in Masters).');
    });
  } catch (e) {}

  // 5. employee names that disagree with Masters (item 8 test run)
  try { IM_employeeAttention_(add); } catch (e) {}

  // write the sheet
  var order = { '🔴': 0, '🟡': 1, '🔵': 2 };
  lines.sort(function (a, b) { return order[a[0]] - order[b[0]]; });
  var red = lines.filter(function (l) { return l[0] === '🔴'; }).length, amber = lines.filter(function (l) { return l[0] === '🟡'; }).length, blue = lines.filter(function (l) { return l[0] === '🔵'; }).length;
  var sh = ss.getSheetByName(NA_SHEET) || ss.insertSheet(NA_SHEET);
  sh.clear();
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm');
  sh.getRange(1, 1).setValue((lines.length ? 'Needs attention: 🔴 ' + red + ' broken · 🟡 ' + amber + ' waiting for you · 🔵 ' + blue + ' for information' : '✅ Nothing needs attention') + '   (updated ' + stamp + ')').setFontWeight('bold').setFontSize(13);
  sh.getRange(3, 1, 1, 5).setValues([['', 'Area', 'What is wrong', 'Where', 'What to do']]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
  if (lines.length) sh.getRange(4, 1, lines.length, 5).setValues(lines).setWrap(true).setVerticalAlignment('top');
  sh.setColumnWidth(1, 36); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 460); sh.setColumnWidth(4, 190); sh.setColumnWidth(5, 520);
  sh.setFrozenRows(3);
  var res = { red: red, amber: amber, blue: blue };
  if (quiet) { if (red) ss.toast('🔴 ' + red + ' problem(s) need attention. See the "' + NA_SHEET + '" sheet (🔗 Interlink → 📋 Needs attention).', '🔗 Interlink', 10); }
  else { ss.setActiveSheet(sh); }
  return res;
}

/**
 * Employee lists that disagree with Masters, for a person to decide. Never changes anything.
 *   not in Masters → 🟡 add them in Masters if they are real employees, or delete them from that list
 *   marked Left in Masters but still listed → 🔵 delete them from that list
 * Compared by IM_normName_ (case / punctuation / spaces only).
 */
function IM_employeeAttention_(add) {
  var emp = IM_mastersEmployees_();
  var shortList = function (a) { return a.slice(0, 8).join(', ') + (a.length > 8 ? ', … (+' + (a.length - 8) + ' more)' : ''); };
  // names = the people the list offers; allNames = everyone on it (inactive rows too) when that differs; fix = what to do if it is behind
  var check = function (area, where, names, allNames, fix) {
    var missing = [], left = [], seen = {}, present = {};
    (allNames || names).forEach(function (n) { present[IM_normName_(n)] = true; });
    names.forEach(function (n) {
      n = String(n).trim(); var k = IM_normName_(n);
      if (!k || seen[k]) return; seen[k] = true;
      if (!emp.all[k]) missing.push(n); else if (emp.left[k]) left.push(n);
    });
    var behind = emp.active.filter(function (e) { return !present[IM_normName_(e.name)]; }).map(function (e) { return e.name; });
    if (behind.length) add('🟡', area, behind.length + ' active Masters employee(s) are not in the ' + where + ' yet: ' + shortList(behind), where,
      fix || 'Open that system once; it adds them by itself. If they are still missing after that, ask your administrator.');
    if (missing.length) add('🟡', area, missing.length + ' name(s) in the ' + where + ' are not in Masters: ' + shortList(missing), where,
      'If they are real employees, add them in Masters (⚙ Masters → Open Admin UI). Otherwise delete them from this list.');
    if (left.length) add('🔵', area, left.length + ' employee(s) that Masters marks as Left are still in the ' + where + ': ' + shortList(left), where,
      'Delete them from this list so they stop appearing in dropdowns.');
  };
  var col = function (sh, c, head) {
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getRange(2, c, sh.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; }).filter(function (v) { return String(v).trim() !== ''; });
  };
  check('Employees', 'PM People list', col(SpreadsheetApp.getActive().getSheetByName(SHEETS.PEOPLE), 1), null,
    'Reload PM (or 🔗 Interlink → Refresh dropdowns from Masters). If they are still missing, ask your administrator.');
  try { check('Employees', 'IMS employee list', col(SpreadsheetApp.openById(IM_IMS_SHEET_ID).getSheetByName('Dropdowns'), 9), null,
    'Open IMS once, or in IMS use 🔗 Interlink → Refresh dropdowns from Masters.'); } catch (e) {}
  try {
    var fid = Masters.getSystemSheetIds().finance;
    if (fid) {
      var m = SpreadsheetApp.openById(fid).getSheetByName('M_Employee');
      if (m && m.getLastRow() >= 2) {
        var head = m.getRange(1, 1, 1, m.getLastColumn()).getValues()[0], cA = head.indexOf('Active');
        var rowsF = m.getRange(2, 1, m.getLastRow() - 1, m.getLastColumn()).getValues().filter(function (r) { return String(r[0]).trim() !== ''; });
        var names = rowsF.filter(function (r) { return cA < 0 || String(r[cA]) !== 'No'; }).map(function (r) { return r[0]; });
        check('Employees', 'Finance employee list (M_Employee)', names, rowsF.map(function (r) { return r[0]; }),
          'Open Finance once, or in Finance use 🔗 Interlink → Refresh employees from Masters (Finance also does it every 10 minutes).');
      }
    }
  } catch (e) {}
}

// ---------------------------------------------------------------- help on the sheet
/** Explanations on the Hardware List headers. Only written when the header is the one we expect; only when changed. */
function IM_installHelpNotes_() {
  var hw = SpreadsheetApp.getActive().getSheetByName(SHEETS.HARDWARE);
  if (!hw) return 0;
  var notes = [
    [IM_STOCK_COL, IM_STOCK_HDR, 'Free stock in IMS: what is in stock minus what is already reserved. "n/a" means this material is not an IMS item. Automatic: do not type here.'],
    [IM_HW_RESERVE_COL, IM_RESERVE_HDR, 'Tick to reserve this row in IMS. What exists is reserved; anything missing goes to Purchases. IMS confirms within about a minute. Ticking again does nothing.'],
    [IM_HW_ITEMSTATUS_COL, IM_ITEMSTATUS_HDR, 'Where this item is now: Not Reserved, Reserving…, Reserved, Partially Consumed, Consumed, Released, plus To order / Ordered / Delivered from Purchases. Automatic. A ⚠ means IMS refused: the reason is in this cell\'s note.'],
    [IM_HW_STATUS_COL, IM_STOCKSTATUS_HDR, 'In Stock or Out of Stock. Fills itself from IMS until you tick Reserve?. You can still choose by hand; setting Out of Stock sends the row to Purchases.'],
    [IM_HW_PUSHED_COL, 'Pushed to Purchases', 'The Purchases row created for this item (PUR-…). Filled by itself.'],
    [IM_HW_RSV_COL, 'IMS Reservation', 'IMS\'s answer for this row: RSV-… | reserved 5. Filled by itself.']
  ], n = 0;
  notes.forEach(function (x) {
    var c = hw.getRange(1, x[0]);
    if (String(c.getValue()) !== x[1]) return;     // not the column we think it is: leave it alone
    if (c.getNote() !== x[2]) { c.setNote(x[2]); n++; }
  });
  return n;
}

/** If someone typed over an automatic formula, put it back (Hardware In Stock + Item Status, Purchases In Stock). Returns cells restored. */
function IM_healFormulaColumns_() {
  var ss = SpreadsheetApp.getActive(), fixed = 0;
  var heal = function (sh, col, hdr, make) {
    if (!sh || String(sh.getRange(1, col).getValue()) !== hdr) return;
    var last = Math.min(sh.getMaxRows(), MAX_ROWS);
    var f = sh.getRange(2, col, last - 1, 1).getFormulas(), out = [], changed = false;
    f.forEach(function (r, i) { if (!r[0] || String(r[0]).charAt(0) !== '=') { out.push([make(i + 2)]); changed = true; fixed++; } else out.push([r[0]]); });
    if (changed) sh.getRange(2, col, last - 1, 1).setFormulas(out);
  };
  heal(ss.getSheetByName(SHEETS.HARDWARE), IM_STOCK_COL, IM_STOCK_HDR, IM_stockFormula_);
  heal(ss.getSheetByName(SHEETS.HARDWARE), IM_HW_ITEMSTATUS_COL, IM_ITEMSTATUS_HDR, IM_itemStatusFormula_);
  heal(ss.getSheetByName(SHEETS.PURCHASES), IM_STOCK_COL, IM_STOCK_HDR, IM_stockFormula_);
  return fixed;
}

/** Menu: ❓ Start here. Creates the tab the first time, then just opens it. */
function IM_showStartHere() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName('Start here');
  if (!sh) {
    sh = ss.insertSheet('Start here', 0);
    var rows = [
      ['Start here: Project Management', 'title'],
      ['The normal flow, step by step', 'h'],
      ['1.  ➕ Add Project (menu 🏗 Interior PM). It reaches Masters, IMS and Finance by itself within minutes. Fill in the Start and Handover dates; the Delay Status shows On Track, Near Deadline or Overdue on its own.', ''],
      ['2.  Hardware List: one row for each item the project needs. Pick the Project and the Material from the dropdowns and enter the Required Qty. Stock Status fills itself (In Stock or Out of Stock, from IMS).', ''],
      ['3.  Tick Reserve? when the project should hold the item. IMS reserves what exists; anything missing goes to Purchases. IMS confirms within about a minute: Item Status shows Reserving…, then Reserved.', ''],
      ['4.  Material not in IMS? Add it with 🔗 Interlink → ➕ Add item to IMS, then pick it from the Material dropdown after a few minutes. Or type it as plain text and set Stock Status to Out of Stock to order it.', ''],
      ['5.  When the store issues material, the IMS Material Requisition must use this project (pick it from the dropdown). The reservation reduces by itself, and Finance shows the cost within about 10 minutes.', ''],
      ['6.  🗑 Delete Project (menu 🏗 Interior PM) removes the project from PM, IMS, Finance and Masters after you type its ID. It cannot be undone.', ''],
      ['What the columns and words mean', 'h'],
      ['Reserve?  tick to make IMS hold the item for this project.', ''],
      ['Stock Status  In Stock or Out of Stock. Fills itself until you tick Reserve?.', ''],
      ['Item Status  where the item is now: Not Reserved, Reserving…, Reserved, Partially Consumed, Consumed, Released, and To order / Ordered / Delivered from Purchases.', ''],
      ['Free stock  what is in stock minus what is already reserved by projects.', ''],
      ['Requisition  the store issuing material to a project (done in IMS). Reservation  IMS holding stock for a project before it is issued.', ''],
      ['If something looks wrong', 'h'],
      ['•  🔗 Interlink → 📋 Needs attention lists what is stuck and what to do.', ''],
      ['•  🔗 Interlink → 🩺 Health check shows green or red for the connections and background jobs.', ''],
      ['•  🔗 Interlink → Refresh Hardware List brings reservations and stock status up to date now.', ''],
      ['•  Hover a column header on the Hardware List for a short explanation. Hover a ⚠ Item Status cell to see why.', '']
    ];
    sh.getRange(1, 1, rows.length, 1).setValues(rows.map(function (r) { return [r[0]]; })).setWrap(true).setVerticalAlignment('top');
    sh.setColumnWidth(1, 900);
    rows.forEach(function (r, i) {
      if (r[1] === 'title') sh.getRange(i + 1, 1).setFontSize(16).setFontWeight('bold');
      if (r[1] === 'h') sh.getRange(i + 1, 1).setFontWeight('bold').setBackground('#DDEBF7');
    });
  }
  ss.setActiveSheet(sh);
}
