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
 * H "Item Status" (right after Unit), a per-row formula over L (IMS Reservation)
 * and K (Pushed to Purchases):
 *   Hardware List: A ID | B Project | C Material | D Spec | E Req Qty | F In Stock |
 *                  G Unit | H Item Status | I Status | J Remarks | K Pushed to Purchases |
 *                  L IMS Reservation
 *   Purchases:     A ID | B Project | C Material | D Spec | E Req Qty | F In Stock |
 *                  G Ordered | H Unit | I Vendor | J PO Issued | K Expected | L Actual |
 *                  M Variance | N Status | O Remarks
 * Picking an item in C fills Unit (G on Hardware, H on Purchases) if blank.
 *
 * Hardware List col I is "Stock Status" (In Stock / Out of Stock dropdown); col H is "Item Status".
 *
 * Item Status (H) is a live per-row formula; no script has to run for it to change:
 *   Reservation: Not Reserved (IMS item, nothing reserved) / Reserving… / Reserved / Released /
 *     Consumed / ⚠ Reservation error. It follows IMS directly: _IMS_Items imports IMS's
 *     'PM Reservation Requests' (K:M) and 'Reservation Register' (O:P), so a release or consume
 *     in IMS shows up here by itself. Col L's text is only the fallback.
 *   Order (from the linked Purchases row's Status): To order / Ordered / Delivered.
 *   Joined as e.g. "Reserved + Ordered". Free-text (non-IMS) materials never show "Not Reserved".
 *
 * Rollout (push the code, then straight away, from the editor):
 *   1. IM_installItems()      idempotent; inserts col F in the live sheets (native
 *                             insert shifts existing formulas / validations / colours),
 *                             then re-seeds ARRAYFORMULAs from the updated code.
 *   2. IM_installItemStatus() inserts Hardware List col H, renames I to Stock Status,
 *                             and (re)writes the IMS imports (idempotent).
 *   3. Open _IMS_Items (View → Hidden sheets), click A2, press "Allow access".
 *   4. IM_verifyWiring() + IM_verifyItemStatus()   PASS/FAIL lists.
 * Re-run IM_installItems() and IM_installItemStatus() after "Create/Rebuild Hardware List Sheet".
 */

var IM_IMS_SHEET_ID = '1CglvPxUWZnIHNeToGEDz8-Xj3eNfvaF351j_ogpRTgk';
var IM_ITEMS_SHEET  = '_IMS_Items';
var IM_MAT_COL      = 3;   // C
var IM_STOCK_COL    = 6;   // F
var IM_STOCK_HDR    = 'In Stock (IMS)';
var IM_HW_ITEMSTATUS_COL = 8;   // H
var IM_ITEMSTATUS_HDR    = 'Item Status';
var IM_STOCKSTATUS_HDR   = 'Stock Status';   // Hardware List col I (was "Status")

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
 *   K:M  PM Reservation Requests → Request ID, Reservation No, Queue Status
 *   O:P  Reservation Register    → Reservation No, Status (Active / Released / Consumed)
 */
function IM_writeStatusImports_(it) {
  var id = IM_IMS_SHEET_ID;
  it.getRange('K1:M1').setValues([['Request ID', 'Res No', 'Queue Status']]);
  it.getRange('K2').setFormula(
    '=IFERROR(QUERY(IMPORTRANGE("' + id + '","PM Reservation Requests!A2:I"),"select Col1,Col9,Col8 where Col1 is not null",0),"")');
  it.getRange('O1:P1').setValues([['Res No', 'Register Status']]);
  it.getRange('O2').setFormula(
    '=IFERROR(QUERY(IMPORTRANGE("' + id + '","Reservation Register!A2:I"),"select Col2,Col9 where Col2 is not null",0),"")');
}

/**
 * Hardware List "Item Status" (col H) for one row, from L (IMS Reservation text), K (PUR code) and F (stock).
 * Reservation part, live from IMS via the helper sheet (falls back to L's text if IMS has no row for it):
 *   Consumed / Released / Reserved (Register status) · ⚠ Reservation error · Reserving… (request still pending)
 *   · Not Reserved (IMS item with no reservation yet) · blank for free-text materials.
 * Order part (linked Purchases Status): To order (Not Ordered) / Ordered (PO Issued, In Transit) /
 *   Delivered / otherwise that status text.  Both present → "Reserved + Ordered".
 */
function IM_itemStatusFormula_(r) {
  var I = IM_ITEMS_SHEET, P = SHEETS.PURCHASES;
  return '=LET(resvRaw,$L' + r + ',purch,$K' + r + ',isIms,ISNUMBER($F' + r + '),' +
    'reqId,IFERROR(REGEXEXTRACT(resvRaw,"^PMR-\\w+"),""),' +
    'rsvNo,IFERROR(REGEXEXTRACT(resvRaw,"^RSV-\\d+"),IF(reqId="","",IFERROR(VLOOKUP(reqId,' + I + '!$K:$M,2,FALSE),""))),' +
    'qStat,IF(reqId="","",IFERROR(VLOOKUP(reqId,' + I + '!$K:$M,3,FALSE),"")),' +
    'regStat,IF(rsvNo="","",IFERROR(VLOOKUP(rsvNo,' + I + '!$O:$P,2,FALSE),"")),' +
    'liveState,IF(regStat="Active","Reserved",IF(OR(regStat="Released",regStat="Consumed"),regStat,"")),' +
    'textState,IF(REGEXMATCH(resvRaw,"consumed"),"Consumed",IF(REGEXMATCH(resvRaw,"released"),"Released",' +
      'IF(REGEXMATCH(resvRaw,"reserved"),"Reserved",""))),' +
    'state,IF(liveState<>"",liveState,textState),' +
    'resvText,IF(state<>"",state,IF(OR(qStat="Error",LEFT(resvRaw,1)="⚠"),"⚠ Reservation error",' +
      'IF(reqId<>"","Reserving…",IF(isIms,"Not Reserved","")))),' +
    'pstat,IF(purch="","",IFERROR(VLOOKUP(purch,' + P + '!$A:$N,14,FALSE),"")),' +
    'ordText,IF(purch="","",IF(OR(pstat="PO Issued",pstat="In Transit"),"Ordered",IF(pstat="Delivered","Delivered",' +
      'IF(OR(pstat="Not Ordered",pstat=""),"To order",pstat)))),' +
    'IF(AND(resvText<>"",ordText<>""),resvText&" + "&ordText,resvText&ordText))';
}

/**
 * Hardware List: insert col H "Item Status" (right after Unit), rename col I to "Stock Status",
 * (re)write the IMS imports it reads, and fill its per-row formulas. Idempotent.
 */
function IM_installItemStatus() {
  var ss = SpreadsheetApp.getActive();
  var hw = ss.getSheetByName(SHEETS.HARDWARE), it = ss.getSheetByName(IM_ITEMS_SHEET);
  if (!hw) { IM_say_('Hardware List sheet not found.'); return; }
  if (!it) { IM_say_('Run IM_installItems() first (the hidden ' + IM_ITEMS_SHEET + ' sheet is missing).'); return; }
  if (hw.getRange(1, IM_HW_ITEMSTATUS_COL).getValue() !== IM_ITEMSTATUS_HDR) {
    if (hw.getRange(1, IM_HW_ITEMSTATUS_COL - 1).getValue() !== 'Unit') {
      IM_say_('Expected "Unit" in column G of Hardware List; nothing changed.');
      return;
    }
    hw.insertColumnAfter(IM_HW_ITEMSTATUS_COL - 1);  // new H; Status / Remarks / Pushed / IMS Reservation shift right
  }
  hw.getRange(1, IM_HW_ITEMSTATUS_COL + 1).setValue(IM_STOCKSTATUS_HDR);  // old "Status" → "Stock Status"
  IM_writeStatusImports_(it);
  hw.getRange(2, IM_HW_ITEMSTATUS_COL, MAX_ROWS - 1, 1).clearDataValidations();  // a new column can inherit Unit's dropdown
  hw.getRange(1, IM_HW_ITEMSTATUS_COL).setValue(IM_ITEMSTATUS_HDR)
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  hw.setColumnWidth(IM_HW_ITEMSTATUS_COL, 150);

  var f = [];
  for (var r = 2; r <= MAX_ROWS; r++) f.push([IM_itemStatusFormula_(r)]);
  hw.getRange(2, IM_HW_ITEMSTATUS_COL, MAX_ROWS - 1, 1).setFormulas(f).setHorizontalAlignment('center');
  IM_say_('Hardware List: H "Item Status" ready, I renamed "Stock Status". Open the hidden ' + IM_ITEMS_SHEET +
          ' sheet and click K2 and O2 to press "Allow access" if they show an error. Then run IM_verifyWiring() and IM_verifyItemStatus().');
}

function IM_say_(msg) {
  try { SpreadsheetApp.getUi().alert('Items', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log(msg); }  // run from editor: no UI context
}

/**
 * Edit trigger for Hardware List / Purchases.
 *  - Material (C) picked  → fill Unit if blank (Hardware G, Purchases H).
 *  - Hardware List Project (B), Material (C) or Required Qty (E) edited → IM_autoEvaluateHardwareRow_
 *    (Status auto-set, reservation queued, shortfall pushed to Purchases).
 */
function IM_onHardwareItemEdit_(e) {
  try {
    if (!e || !e.range) return;
    var r = e.range, sh = r.getSheet();
    var t = IM_itemSheets_().filter(function (x) { return x.name === sh.getName(); })[0];
    if (!t || r.getRow() < 2 || r.getNumRows() !== 1 || r.getNumColumns() !== 1) return;
    var col = r.getColumn();
    if (col === IM_MAT_COL) IM_fillUnit_(e, sh, t, r);
    if (t.name === SHEETS.HARDWARE && (col === 2 || col === IM_MAT_COL || col === 5)) {
      var res = IM_autoEvaluateHardwareRow_(sh, r.getRow());
      if (res.indexOf('skipped') !== 0) e.source.toast(res, 'Hardware List', 8);  // stay quiet while the row is still incomplete
    }
  } catch (err) {
    Logger.log('IM_onHardwareItemEdit_ error: ' + err.message);
  }
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
  [[H, 5, 'Required Qty'], [H, 6, IM_STOCK_HDR], [H, 7, 'Unit'], [H, 8, IM_ITEMSTATUS_HDR], [H, 9, IM_STOCKSTATUS_HDR],
   [H, 11, 'Pushed to Purchases'], [H, 12, 'IMS Reservation'],
   [P, 5, 'Required Qty'], [P, 6, IM_STOCK_HDR], [P, 7, 'Ordered Qty'], [P, 8, 'Unit'], [P, 9, 'Vendor'],
   [P, 10, 'PO Issued'], [P, 11, 'Expected Delivery'], [P, 12, 'Actual Delivery'], [P, 13, 'Variance (Days)'],
   [P, 14, 'Status'], [P, 15, 'Remarks']
  ].forEach(function (c) { check(c[0] + ' header col ' + c[1] + ' = ' + c[2], hdr(c[0], c[1]) === c[2]); });

  var pur = ss.getSheetByName(P);
  var mf = pur ? pur.getRange('M2').getFormula() : '';
  check('Purchases M2 variance = L − K', mf.indexOf('L2:L-K2:K') >= 0 && mf.indexOf('K2:K=""') >= 0);

  var hwSheet = ss.getSheetByName(H);
  check('Hardware H2 Item Status formula reads K and L', !!hwSheet && /\$L2/.test(hwSheet.getRange('H2').getFormula()) && /\$K2/.test(hwSheet.getRange('H2').getFormula()));
  check('Hardware I (Stock Status) dropdown', dvHas(H, 'I2', 'Out of Stock'));
  check('Hardware G (Unit) dropdown',   dvHas(H, 'G2', 'Pcs'));
  check('Hardware H (Item Status) has no dropdown', !!hwSheet && !hwSheet.getRange('H2').getDataValidation());
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
// Live check of the Item Status formula (Hardware List col H). Uses the first empty
// Hardware row and first empty Purchases row as scratch, then clears what it wrote.
// -----------------------------------------------------------------------------
function IM_verifyItemStatus() {
  var ss = SpreadsheetApp.getActive();
  var hw = ss.getSheetByName(SHEETS.HARDWARE), pur = ss.getSheetByName(SHEETS.PURCHASES);
  if (!hw || !pur) { IM_say_('Hardware List / Purchases sheet missing.'); return; }

  // [F stock (number = IMS item, "n/a" = free text), L IMS Reservation, K PUR code, linked Purchases status, expected]
  // These use made-up RSV/PMR numbers, so they exercise the text fallback. The live IMS lookup
  // (Register / queue imports) is checked by the first real reservation: reserve, then release in IMS.
  var cases = [
    [30,    '',                     '',         '',            'Not Reserved'],
    [30,    'PMR-t | pending | 5',  '',         '',            'Reserving…'],
    [30,    'RSV-1 | reserved 5',   '',         '',            'Reserved'],
    [30,    'RSV-1 | released 5',   '',         '',            'Released'],
    [30,    'RSV-1 | consumed 5',   '',         '',            'Consumed'],
    [30,    '⚠ IMS error',          '',         '',            '⚠ Reservation error'],
    ['n/a', '',                     '',         '',            ''],
    ['n/a', '',                     'PUR-TEST', 'Not Ordered', 'To order'],
    ['n/a', '',                     'PUR-TEST', 'In Transit',  'Ordered'],
    ['n/a', '',                     'PUR-TEST', 'Delivered',   'Delivered'],
    [30,    '',                     'PUR-TEST', 'Not Ordered', 'Not Reserved + To order'],
    [0,     '',                     'PUR-TEST', 'PO Issued',   'Not Reserved + Ordered'],
    [30,    'RSV-1 | reserved 5',   'PUR-TEST', 'PO Issued',   'Reserved + Ordered'],
    [30,    'RSV-1 | consumed 5',   'PUR-TEST', 'Delivered',   'Consumed + Delivered']
  ];
  var out = [];
  var lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  var hr = firstEmptyRow_(hw, 1), pr = firstEmptyRow_(pur, 1);
  var reset = function () {
    hw.getRange(hr, 11, 1, 2).clearContent();                 // K, L
    hw.getRange(hr, IM_STOCK_COL).setFormula(IM_stockFormula_(hr));
    pur.getRange(pr, 1).clearContent(); pur.getRange(pr, 14).clearContent();  // A ID, N Status
  };
  try {
    cases.forEach(function (c) {
      hw.getRange(hr, IM_STOCK_COL).setValue(c[0]);           // simulate an IMS item / a free-text one
      hw.getRange(hr, 11, 1, 2).setValues([[c[2], c[1]]]);
      pur.getRange(pr, 1).setValue(c[2] ? 'PUR-TEST' : '');
      pur.getRange(pr, 14).setValue(c[3]);
      SpreadsheetApp.flush();
      var got = String(hw.getRange(hr, IM_HW_ITEMSTATUS_COL).getValue());
      out.push((got === c[4] ? 'PASS  ' : 'FAIL  ') + 'F=' + c[0] + ' L="' + c[1] + '" K="' + c[2] + '" purchase="' + c[3] + '" → "' + got + '"' +
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
