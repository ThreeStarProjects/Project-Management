/**
 * ============================================================================
 * 3SP PROJECT MANAGEMENT — Menu
 * Custom "🏗 Interior PM" menu + the admin actions it exposes: Add Project,
 * Raise Ticket (opens Index.html modal), Open Project View, Delete Project,
 * and setup/maintenance utilities.
 * ============================================================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗 Interior PM')
    .addItem('➕ Add Project', 'menuAddProject')
    .addItem('🎫 Raise Ticket', 'menuRaiseTicket')
    .addSeparator()
    .addItem('📋 Open Project View', 'menuOpenProjectView')
    .addItem('🗑 Delete Project', 'menuDeleteProject')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⚙️ Setup')
      .addItem('Run Full Setup (nukes everything)', 'setupAll')
      .addSeparator()
      .addItem('Rebuild Dashboard only', 'rebuildDashboard')
      .addItem('Rebuild Project View only', 'rebuildProjectView')
      .addItem('Create/Rebuild Hardware List Sheet', 'rebuildHardwareList')
      .addItem('Import Purchases → Hardware List', 'importPurchasesIntoHardwareList')
      .addItem('🔍 Diagnose Dashboard Report', 'diagnoseReports')
      .addItem('🩹 Reseed Formulas (fix dead dropdowns)', 'reseedFormulas')
      .addSeparator()
      .addItem('Reinstall Data Validations', 'installDataValidations')
      .addItem('Reinstall Formatting', 'installFormatting')
      .addItem('Reinstall Triggers', 'installTriggers')
      .addItem('🧹 Remove All Warnings', 'removeAllWarnings')
      .addItem('Reset Named Ranges', 'resetNamedRanges'))
    .addToUi();
}

// ── CORE ACTIONS ───────────────────────────────────────────────────────────

function menuAddProject() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Add Project', 'Project Name:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const name = (r.getResponseText() || '').trim();
  if (!name) { ui.alert('Name required.'); return; }
  const id = addProject_(name);
  // Also seed a Milestones row so % Done starts computing
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mile = ss.getSheetByName(SHEETS.MILESTONES);
  const label = id + ' — ' + name;
  if (mile && !findRowByValue_(mile, 1, label)) {
    mile.getRange(firstEmptyRow_(mile, 1), 1).setValue(label);
  }
  ui.alert('✅ Added ' + id + '. Fill in the rest of the row in the Projects tab.');
}

function addProject_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const proj = ss.getSheetByName(SHEETS.PROJECTS);
  const id = 'PRJ-' + String(getMaxIdNumber_(proj, 'PRJ') + 1).padStart(3, '0');
  const row = firstEmptyRow_(proj, 1);
  proj.getRange(row, 1).setValue(id);
  proj.getRange(row, 2).setValue(name);
  proj.getRange(row, 7).setValue('In Progress');  // col G Overall Status (shifted after Owner/Contract removed)
  return id;
}

function menuRaiseTicket() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('Index').setWidth(720).setHeight(720),
    '🎫 Raise a Ticket'
  );
}

function menuOpenProjectView() {
  const match = findProjectByQuery_();
  if (!match) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const view = ss.getSheetByName(SHEETS.VIEW);
  view.getRange('B3').setValue(match.label);
  ss.setActiveSheet(view);
  view.setActiveRange(view.getRange('B3'));
}

function menuDeleteProject() {
  const match = findProjectByQuery_();
  if (!match) return;
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Delete project?',
    'Permanently delete ' + match.id + ' — ' + match.name + ' from Projects and every related row in Milestones, Hardware List, Purchases, Production, Snags, and Tickets. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let total = 0;
  // (sheet name, project-label column 1-based, match value)
  const targets = [
    [SHEETS.PROJECTS,   1,  match.id],
    [SHEETS.MILESTONES, 1,  match.label],
    [SHEETS.HARDWARE,   2,  match.label],
    [SHEETS.PURCHASES,  2,  match.label],
    [SHEETS.PRODUCTION, 2,  match.label],
    [SHEETS.SNAGS,      2,  match.label],
    [SHEETS.TICKETS,    11, match.label]
  ];
  targets.forEach(([n, c, v]) => { total += deleteAllRowsByValue_(ss.getSheetByName(n), c, v); });
  // deleteRow on row 2 can wipe ARRAYFORMULA seeds. Re-install unconditionally.
  reseedArrayFormulas_(ss);
  ui.alert('✅ Deleted ' + match.id + '. ' + total + ' row(s) removed.');
}

function findProjectByQuery_() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Find Project', 'Partial Project ID or Name:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const q = (r.getResponseText() || '').trim().toLowerCase();
  if (!q) return null;
  const proj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PROJECTS);
  const last = proj.getLastRow();
  if (last < 2) { ui.alert('No projects yet.'); return null; }
  // Read cols A..I (ID, Name, ..., Label at col 9).
  const data = proj.getRange(2, 1, last - 1, 9).getValues();
  const matches = [];
  for (let i = 0; i < data.length; i++) {
    const id = String(data[i][0] || '').toLowerCase();
    const name = String(data[i][1] || '').toLowerCase();
    if (!id) continue;
    if (id.includes(q) || name.includes(q)) matches.push({ row: i + 2, id: data[i][0], name: data[i][1], label: data[i][8] });
  }
  if (!matches.length) { ui.alert('No match: ' + q); return null; }
  if (matches.length === 1) return matches[0];
  const list = matches.slice(0, 10).map((m, i) => `${i+1}. ${m.id} — ${m.name}`).join('\n');
  const p = ui.prompt('Multiple matches', list + '\n\nEnter number:', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return null;
  const idx = parseInt(p.getResponseText(), 10) - 1;
  if (isNaN(idx) || idx < 0 || idx >= matches.length) { ui.alert('Invalid.'); return null; }
  return matches[idx];
}

// ── REBUILD / DIAGNOSE / RESEED ────────────────────────────────────────────

/** Read the Dashboard report cell and report OK / empty / errored. */
function diagnoseReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!dash) { SpreadsheetApp.getUi().alert('Dashboard sheet not found.'); return; }
  const targets = [
    ['A17', 'Report · Project Status']
  ];
  const lines = targets.map(([cell, name]) => {
    const v = String(dash.getRange(cell).getDisplayValue() || '');
    let status;
    if (v.startsWith('#'))            status = '❌ ERROR: ' + v;
    else if (v.startsWith('(pick '))  status = '⚪ waiting for input';
    else if (v.startsWith('No ') || v === '') status = '⚪ empty (no matching data)';
    else                              status = '✅ populated (first cell: "' + v.substring(0, 40) + '")';
    return `${name} (${cell}): ${status}`;
  });
  SpreadsheetApp.getUi().alert('Dashboard Report Health:\n\n' + lines.join('\n\n'));
}

function rebuildDashboard() {
  rebuildOneSheet_(SHEETS.DASHBOARD, createDashboardSheet_, formatDashboard_);
}

function rebuildProjectView() {
  rebuildOneSheet_(SHEETS.VIEW, createProjectViewSheet_, formatProjectView_);
  installDataValidations();  // restores the B3 project-selector dropdown
}

function rebuildHardwareList() {
  rebuildOneSheet_(SHEETS.HARDWARE, createHardwareListSheet_, null);
  installDataValidations();  // seeds Project/Unit/Status dropdowns
}

function rebuildOneSheet_(name, createFn, formatFn) {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('Rebuild ' + name + '?',
    'Deletes and recreates the ' + name + ' tab only. Every other sheet (and all seeded data) is untouched. Continue?',
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(name);
  if (old) {
    old.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => { try { p.remove(); } catch (e) {} });
    ss.deleteSheet(old);
  }
  createFn(ss);
  if (formatFn) formatFn(ss);
  ui.alert('✅ ' + name + ' rebuilt.');
}

/**
 * One-shot backfill: create a Hardware List row for every existing Purchases row.
 * Each backfilled row is marked Out of Stock and col I (Pushed) is stamped with the
 * PUR-code so the onEdit trigger will NOT create a duplicate Purchase later.
 */
function importPurchasesIntoHardwareList() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const purSheet = ss.getSheetByName(SHEETS.PURCHASES);
  const hwSheet  = ss.getSheetByName(SHEETS.HARDWARE);
  if (!purSheet || !hwSheet) { ui.alert('Missing Purchases or Hardware List sheet.'); return; }

  const purLast = purSheet.getLastRow();
  if (purLast < 2) { ui.alert('No purchases to import.'); return; }

  const purData = purSheet.getRange(2, 1, purLast - 1, 7).getValues();

  const alreadyPushed = new Set();
  const hwLast = hwSheet.getLastRow();
  if (hwLast >= 2) {
    hwSheet.getRange(2, 9, hwLast - 1, 1).getValues().forEach(r => {
      if (r[0]) alreadyPushed.add(String(r[0]));
    });
  }

  let nextHwNum = getMaxIdNumber_(hwSheet, 'HW') + 1;
  const toInsert = [];
  let skipped = 0;
  for (const p of purData) {
    const [purCode, project, material, spec, reqQty, ordQty, unit] = p;
    if (!purCode || !material) continue;
    if (alreadyPushed.has(String(purCode))) { skipped++; continue; }
    const hwCode = 'HW-' + String(nextHwNum++).padStart(3, '0');
    const qty = (reqQty !== '' && reqQty !== null) ? reqQty : ordQty;
    toInsert.push([hwCode, project, material, spec, qty, unit, 'Out of Stock', 'Imported from ' + purCode, purCode]);
  }

  if (toInsert.length === 0) {
    ui.alert('Nothing to import. ' + skipped + ' Purchase(s) already linked to Hardware List.');
    return;
  }

  const confirm = ui.alert(
    'Import ' + toInsert.length + ' Purchases into Hardware List?',
    (skipped ? skipped + ' already linked and will be skipped.\n\n' : '') +
    'Each new Hardware row will be marked Out of Stock and linked to its Purchase. No duplicate Purchase rows will be created.',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const startRow = firstEmptyRow_(hwSheet, 1);
  hwSheet.getRange(startRow, 1, toInsert.length, 9).setValues(toInsert);
  ui.alert('✅ Imported ' + toInsert.length + ' row(s). Skipped: ' + skipped + '.');
}

/**
 * Re-install all ARRAYFORMULA seed cells across every sheet that has them.
 * Fixes the case where deleteRow (via Delete Project OR a user right-click)
 * nuked the row-2 seed, killing the whole column's spill.
 */
function reseedArrayFormulas_(ss) {
  const PR = SHEETS.PROJECTS, M = SHEETS.MILESTONES;
  const proj = ss.getSheetByName(PR);
  if (proj) {
    proj.getRange('I2').setFormula(`=ARRAYFORMULA(IF(A2:A="","",A2:A&" — "&B2:B))`);
    proj.getRange('J2').setFormula(`=ARRAYFORMULA(IF(F2:F="","",IF(TODAY()>F2:F,"Overdue",IF(F2:F-TODAY()<=7,"Near Deadline","On Track"))))`);
    proj.getRange('K2').setFormula(`=ARRAYFORMULA(IF(F2:F="","",F2:F-TODAY()))`);
    STAGES.forEach(st => {
      const colLetter = colLetter_(st.col);
      const sumParts = [];
      for (let i = st.start; i <= st.end; i++) {
        const c = colLetter_(i + 3);
        sumParts.push(`COUNTIFS('${M}'!A:A,I2:I,'${M}'!${c}:${c},"Yes")`);
      }
      const total = st.end - st.start + 1;
      proj.getRange(colLetter + '2').setFormula(
        `=ARRAYFORMULA(IF(A2:A="","",IFERROR((${sumParts.join('+')})/${total},0)))`
      );
    });
    proj.getRange('Q2').setFormula(
      `=ARRAYFORMULA(IF(A2:A="","",IFERROR(COUNTIF(FILTER('${M}'!C:AA,'${M}'!A:A=I2:I),"Yes")/25,0)))`
    );
  }
  const mile = ss.getSheetByName(SHEETS.MILESTONES);
  if (mile) mile.getRange('B2').setFormula(
    `=ARRAYFORMULA(IFERROR(VLOOKUP(A2:A, {${PR}!I:I, ${PR}!B:B}, 2, FALSE), ""))`
  );
  const pur = ss.getSheetByName(SHEETS.PURCHASES);
  if (pur) pur.getRange('L2').setFormula(
    `=ARRAYFORMULA(IF(OR(J2:J="",K2:K=""),"",IFERROR(K2:K-J2:J,"")))`
  );
  const prod = ss.getSheetByName(SHEETS.PRODUCTION);
  if (prod) {
    prod.getRange('H2').setFormula(
      `=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,{${PR}!I:I,${PR}!F:F},2,FALSE)-1,"")))`
    );
    prod.getRange('J2').setFormula(`=ARRAYFORMULA(IF(OR(H2:H="",I2:I=""),"",IFERROR(I2:I-H2:H,"")))`);
    prod.getRange('M2').setFormula(
      `=ARRAYFORMULA(IF(A2:A="","",IF(I2:I<>"","Completed",IF(G2:G<>"","In Progress","Not Started"))))`
    );
  }
  const snag = ss.getSheetByName(SHEETS.SNAGS);
  if (snag) snag.getRange('K2').setFormula(`=ARRAYFORMULA(IF(OR(I2:I="",J2:J=""),"",IFERROR(J2:J-I2:I,"")))`);
  const tkt = ss.getSheetByName(SHEETS.TICKETS);
  if (tkt) {
    tkt.getRange('V2').setFormula(
      `=ARRAYFORMULA(IF(A2:A="","",IF(OR(I2:I="Resolved",I2:I="Closed"),"Done",IF(J2:J="","",IF(J2:J<TODAY(),"Overdue",IF(J2:J-TODAY()<=1,"Due Soon","On Track"))))))`
    );
    tkt.getRange('W2').setFormula(`=ARRAYFORMULA(IF(OR(B2:B="",S2:S=""),"",IFERROR(INT(S2:S-B2:B),"")))`);
  }
}

/** Public wrapper so users can self-heal dead dropdowns via the menu. */
function reseedFormulas() {
  reseedArrayFormulas_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('✅ All ARRAYFORMULA seed cells reinstalled.');
}
