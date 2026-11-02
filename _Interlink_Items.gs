/**
 * _Interlink_Items.gs  (v0.4 Phase 3)
 * -----------------------------------------------------------------------------
 * Hardware List ← IMS items. Additive only: appends cols J/K to Hardware List
 * and adds a hidden helper sheet. No existing PM code is modified.
 *
 *   _IMS_Items (hidden)  IMPORTRANGE of IMS Item Master + SFG Master + stock
 *   Hardware List J      Item Code dropdown (label list from _IMS_Items)
 *   Hardware List K      In Stock (IMS) = Available − Reserved, live formula
 *
 * Picking a code in J fills Material (C) and Unit (F) only if they're blank.
 *
 * One-time: run IM_installItems() from the editor, then open _IMS_Items,
 * click A2 and press "Allow access" (IMPORTRANGE needs it once).
 * Re-run IM_installItems() after "Create/Rebuild Hardware List Sheet".
 */

var IM_IMS_SHEET_ID = '1CglvPxUWZnIHNeToGEDz8-Xj3eNfvaF351j_ogpRTgk';
var IM_ITEMS_SHEET  = '_IMS_Items';
var IM_HW_ITEM_COL  = 10;  // J
var IM_HW_STOCK_COL = 11;  // K

function IM_installItems() {
  var ss = SpreadsheetApp.getActive();
  var hw = ss.getSheetByName(SHEETS.HARDWARE);
  if (!hw) { SpreadsheetApp.getUi().alert('Hardware List sheet not found.'); return; }

  // --- helper sheet ---------------------------------------------------------
  var it = ss.getSheetByName(IM_ITEMS_SHEET) || ss.insertSheet(IM_ITEMS_SHEET);
  it.clear();
  var id = IM_IMS_SHEET_ID;
  it.getRange('A1:E1').setValues([['Code', 'Label', 'UOM', 'Item Name', 'Net Stock']]);
  it.getRange('H1:I1').setValues([['Stock Code', 'Stock Net']]);
  // Items + SFGs (inactive rows dropped). Search Label = Item Master T / SFG Master N.
  it.getRange('A2').setFormula(
    '={QUERY(IMPORTRANGE("' + id + '","Item Master!A2:T"),"select Col1,Col20,Col7,Col4 where Col1 is not null and Col19 <> \'No\'",0);' +
    'QUERY(IMPORTRANGE("' + id + '","SFG Master!A2:N"),"select Col1,Col14,Col7,Col4 where Col1 is not null and Col11 <> \'No\'",0)}');
  // Stock: RM = Snapshot Available − Reserved; SFG = SFG Master Available (already net).
  it.getRange('H2').setFormula(
    '={QUERY(IMPORTRANGE("' + id + '","Stock Snapshot!A2:J"),"select Col1,Col9-Col10 where Col1 is not null",0);' +
    'QUERY(IMPORTRANGE("' + id + '","SFG Master!A2:M"),"select Col1,Col13 where Col1 is not null",0)}');
  it.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IFERROR(VLOOKUP(A2:A,H:I,2,FALSE),0)))');
  it.hideSheet();

  // --- Hardware List cols J/K -----------------------------------------------
  var hdr = hw.getRange(1, IM_HW_ITEM_COL, 1, 2);
  hdr.setValues([['Item Code (IMS)', 'In Stock (IMS)']])
     .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG)
     .setFontWeight('bold').setFontSize(11)
     .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  hw.setColumnWidth(IM_HW_ITEM_COL, 260);
  hw.setColumnWidth(IM_HW_STOCK_COL, 110);

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(it.getRange('B2:B5000'), true)
    .setAllowInvalid(false).build();
  hw.getRange(2, IM_HW_ITEM_COL, MAX_ROWS - 1, 1).setDataValidation(rule);

  hw.getRange(2, IM_HW_STOCK_COL).setFormula(
    '=ARRAYFORMULA(IF(J2:J="","",IFERROR(VLOOKUP(J2:J,' + IM_ITEMS_SHEET + '!B:E,4,FALSE),"n/a")))');

  // --- autofill trigger (separate from the existing onEditHandler) ---------
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'IM_onHardwareItemEdit_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('IM_onHardwareItemEdit_').forSpreadsheet(ss).onEdit().create();

  var msg = 'Now open the hidden "' + IM_ITEMS_SHEET + '" sheet (View → Hidden sheets), click A2 and press "Allow access".';
  try { SpreadsheetApp.getUi().alert('Items wired', msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { Logger.log('Items wired. ' + msg); }  // run from editor: no UI context
}

/** Item Code picked in J → fill Material (C) and Unit (F) if blank. */
function IM_onHardwareItemEdit_(e) {
  try {
    if (!e || !e.range) return;
    var r = e.range, sh = r.getSheet();
    if (sh.getName() !== SHEETS.HARDWARE || r.getRow() < 2) return;
    if (r.getColumn() !== IM_HW_ITEM_COL || r.getNumRows() !== 1) return;
    var label = String(r.getValue());
    if (!label) return;

    var rows = e.source.getSheetByName(IM_ITEMS_SHEET).getRange('A2:D').getValues();
    var m = rows.filter(function (x) { return x[1] === label; })[0];
    if (!m) return;

    var row = r.getRow();
    if (!sh.getRange(row, 3).getValue()) sh.getRange(row, 3).setValue(m[3]);   // Material ← Item Name
    var units = ['Sq.ft', 'Sq.m', 'Pcs', 'Box', 'Roll', 'Kg', 'Meter', 'Set', 'Other'];
    var u = units.filter(function (x) { return x.toLowerCase() === String(m[2]).toLowerCase(); })[0];
    if (u && !sh.getRange(row, 6).getValue()) sh.getRange(row, 6).setValue(u); // Unit ← UOM if in PM's list
  } catch (err) {
    Logger.log('IM_onHardwareItemEdit_ error: ' + err.message);
  }
}
