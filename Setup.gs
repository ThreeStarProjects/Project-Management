/**
 * ============================================================================
 * 3SP PROJECT MANAGEMENT — Setup
 * One-shot bootstrap for the whole spreadsheet. Run setupAll() from the editor
 * once. Rebuild* functions in Menu.gs let you refresh individual tabs later.
 * ============================================================================
 */

function setupAll() {
  setupSheets();
  setNamedRanges_();
  installDataValidations();
  installFormatting();
  installTriggers();
  removeScriptProtections_();  // nuke any lingering 3SP warning banners
  SpreadsheetApp.getUi().alert('✅ Setup complete. Reload the sheet to see the menu.');
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  removeScriptProtections_();

  // Keep a scratch tab so Google always has 1 sheet during rebuild.
  let scratch = ss.getSheetByName('__scratch__') || ss.insertSheet('__scratch__');
  SpreadsheetApp.flush();

  // Delete both current and legacy sheets — legacy covers tabs from earlier
  // webapp experiments + the Invoices/Log tabs dropped in this slim-down.
  [...Object.values(SHEETS), ...LEGACY_SHEETS].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) { try { ss.deleteSheet(sh); } catch (e) { Logger.log('skip ' + name + ': ' + e.message); } }
  });
  SpreadsheetApp.flush();

  // Order matters: reference sheets first so validations from other sheets can point to them.
  createPeopleSheet_(ss);
  createVendorsSheet_(ss);
  createProjectsSheet_(ss);
  createMilestonesSheet_(ss);
  createHardwareListSheet_(ss);
  createPurchasesSheet_(ss);
  createProductionSheet_(ss);
  createSnagsSheet_(ss);
  createTicketsSheet_(ss);
  createProjectViewSheet_(ss);
  createDashboardSheet_(ss);

  // Reorder tabs for a sensible left-to-right flow
  const order = [SHEETS.PROJECTS, SHEETS.VIEW, SHEETS.DASHBOARD, SHEETS.MILESTONES,
                 SHEETS.HARDWARE, SHEETS.PURCHASES, SHEETS.PRODUCTION,
                 SHEETS.SNAGS, SHEETS.TICKETS, SHEETS.PEOPLE, SHEETS.VENDORS];
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });

  const s2 = ss.getSheetByName('__scratch__');
  if (s2 && ss.getSheets().length > 1) { try { ss.deleteSheet(s2); } catch (e) {} }
  Logger.log('Sheets created.');
}

function removeScriptProtections_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sh => {
    try {
      sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
        if ((p.getDescription() || '').startsWith('3SP:')) { try { p.remove(); } catch (e) {} }
      });
      sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
        if ((p.getDescription() || '').startsWith('3SP:')) { try { p.remove(); } catch (e) {} }
      });
    } catch (e) {}
  });
}

// ── NAMED RANGES ───────────────────────────────────────────────────────────
// Wraps common cross-sheet targets. When we reshuffle columns later, only
// the named-range definition needs to move, not every formula.
function setNamedRanges_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = [
    ['PROJECT_LABELS', SHEETS.PROJECTS, 'I2:I' + MAX_ROWS],  // Projects.Label is col I now
    ['VENDOR_NAMES',   SHEETS.VENDORS,  'A2:A100'],
    ['PEOPLE_NAMES',   SHEETS.PEOPLE,   'A2:A100']
  ];
  ss.getNamedRanges().forEach(nr => { try { nr.remove(); } catch (e) {} });
  names.forEach(([name, sheetName, a1]) => {
    const sh = ss.getSheetByName(sheetName);
    if (sh) ss.setNamedRange(name, sh.getRange(a1));
  });
}

/** Public wrapper so it appears in the Apps Script Run dropdown + menu. */
function resetNamedRanges() { setNamedRanges_(); }

// ── TRIGGERS ───────────────────────────────────────────────────────────────
function installTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEditHandler') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss).onEdit().create();
  Logger.log('Trigger installed.');
}

// ── PROTECTIONS ────────────────────────────────────────────────────────────
/**
 * Strips every 3SP warning-only protection from the spreadsheet. Previous
 * builds installed protections on header rows, read-only tabs, and ARRAYFORMULA
 * seed cells — those triggered the "Heads up" dialog on every edit. Protection
 * is no longer installed; this function just cleans up any old ones.
 *
 * If a formula seed gets accidentally wiped, run menu → ⚙️ Setup → Reseed Formulas.
 */
function removeAllWarnings() {
  removeScriptProtections_();
  SpreadsheetApp.getUi().alert('✅ All 3SP warning banners removed.');
}
