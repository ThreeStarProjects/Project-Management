/**
 * ============================================================================
 * 3SP PROJECT MANAGEMENT — Sheets
 * Per-sheet builders + validations + formatting + conditional formatting +
 * project color bands. Everything that defines sheet shape or style lives here.
 * ============================================================================
 */

// ── SHEET BUILDERS ─────────────────────────────────────────────────────────

function createPeopleSheet_(ss) {
  const sh = ss.insertSheet(SHEETS.PEOPLE);
  sh.getRange(1, 1, 1, 5).setValues([['Name','Department','Role','Phone','Email']]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 180); sh.setColumnWidth(2, 140); sh.setColumnWidth(3, 160);
  sh.setColumnWidth(4, 120); sh.setColumnWidth(5, 200);
}

function createProjectsSheet_(ss) {
  const sh = ss.insertSheet(SHEETS.PROJECTS);
  const headers = [
    'Project ID',      // A  auto PRJ-###
    'Project Name',    // B
    'Client',          // C
    'Location',        // D
    'Start Date',      // E
    'Handover Date',   // F
    'Overall Status',  // G  In Progress / At Risk / Completed
    'Remarks',         // H
    // ── formulas below (positioned AFTER user cols) ──
    'Label',           // I  ARRAYFORMULA A & " — " & B
    'Delay Status',    // J  ARRAYFORMULA On Track / Near Deadline / Overdue
    'Days Left',       // K  ARRAYFORMULA F - TODAY()
    'Pre-Execution',   // L  ARRAYFORMULA stage % from Milestones
    'Procurement',     // M
    'Production',      // N
    'Delivery',        // O
    'Installation',    // P
    'Completion %'     // Q  ARRAYFORMULA count Yes / 25
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  // ARRAYFORMULA columns live in row 2 and spill down.
  const M = SHEETS.MILESTONES;
  sh.getRange('I2').setFormula(`=ARRAYFORMULA(IF(A2:A="","",A2:A&" — "&B2:B))`);
  sh.getRange('J2').setFormula(`=ARRAYFORMULA(IF(F2:F="","",IF(TODAY()>F2:F,"Overdue",IF(F2:F-TODAY()<=7,"Near Deadline","On Track"))))`);
  sh.getRange('K2').setFormula(`=ARRAYFORMULA(IF(F2:F="","",F2:F-TODAY()))`);

  // Stage % — count Yes across the stage's milestone cols. We sum per-column
  // COUNTIFS rather than passing a multi-col range (that was the parse error).
  STAGES.forEach(st => {
    const colLetter = colLetter_(st.col);
    const sumParts = [];
    for (let i = st.start; i <= st.end; i++) {
      const c = colLetter_(i + 3); // Milestones col C = milestone 0
      sumParts.push(`COUNTIFS('${M}'!A:A,I2:I,'${M}'!${c}:${c},"Yes")`);
    }
    const total = st.end - st.start + 1;
    sh.getRange(colLetter + '2').setFormula(
      `=ARRAYFORMULA(IF(A2:A="","",IFERROR((${sumParts.join('+')})/${total},0)))`
    );
  });

  sh.getRange('Q2').setFormula(
    `=ARRAYFORMULA(IF(A2:A="","",IFERROR(COUNTIF(FILTER('${M}'!C:AA,'${M}'!A:A=I2:I),"Yes")/25,0)))`
  );

  sh.getRange('E:F').setNumberFormat('dd/mm/yyyy');
  sh.getRange('L:P').setNumberFormat('0%');
  sh.getRange('Q:Q').setNumberFormat('0.0%');

  sh.setColumnWidth(1, 90); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 160);
  sh.setColumnWidth(4, 140); sh.setColumnWidth(9, 260);
  sh.hideColumns(9); // hide Label helper (I)
}

function createMilestonesSheet_(ss) {
  const sh = ss.insertSheet(SHEETS.MILESTONES);
  const headers = ['Project Label', 'Project Name'].concat(
    MILESTONES.map((m, i) => `${i + 1}. ${m}`)
  );
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  // Projects.Label is now col I.
  sh.getRange('B2').setFormula(
    `=ARRAYFORMULA(IFERROR(VLOOKUP(A2:A, {${SHEETS.PROJECTS}!I:I, ${SHEETS.PROJECTS}!B:B}, 2, FALSE), ""))`
  );

  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 180);
}

function createHardwareListSheet_(ss) {
  try { CacheService.getScriptCache().remove(LAYOUT_CACHE_KEY); } catch (e) {}
  const sh = ss.insertSheet(SHEETS.HARDWARE);
  const headers = [
    'Hardware ID',          // A  auto HW-###
    'Project Label',        // B  dropdown from Projects
    'Material',             // C
    'Specification',        // D
    'Required Qty',         // E
    'In Stock (IMS)',       // F  per-row formula, see _Interlink_Items.gs
    'Unit',                 // G  dropdown
    'Reserve?',             // H  checkbox: tick sends the reservation to IMS (see _Interlink_Items.gs)
    'Item Status',          // I  per-row formula: Reserved / Released / Consumed / Ordered (see _Interlink_Items.gs)
    'Stock Status',         // J  In Stock / Out of Stock — flip drives push to Purchases
    'Remarks',              // K
    'Pushed to Purchases',  // L  trigger stamps PUR-code; prevents duplicate push
    'IMS Reservation'       // M  reservation request / result text
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);

  sh.setColumnWidth(1, 90);   sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 200);  sh.setColumnWidth(4, 200);
  sh.setColumnWidth(5, 100);  sh.setColumnWidth(6, 110);
  sh.setColumnWidth(7, 90);   sh.setColumnWidth(8, 80);
  sh.setColumnWidth(9, 150);  sh.setColumnWidth(10, 110);
  sh.setColumnWidth(11, 220); sh.setColumnWidth(12, 120);
  sh.setColumnWidth(13, 200);
}

function createPurchasesSheet_(ss) {
  try { CacheService.getScriptCache().remove(LAYOUT_CACHE_KEY); } catch (e) {}
  const sh = ss.insertSheet(SHEETS.PURCHASES);
  const headers = [
    'Purchase ID',      // A  auto PUR-###
    'Project Label',    // B
    'Material',         // C
    'Specification',    // D
    'Required Qty',     // E
    'In Stock (IMS)',   // F  per-row formula, see _Interlink_Items.gs
    'Ordered Qty',      // G
    'Unit',             // H
    'Vendor',           // I  dropdown from Vendors
    'PO Issued',        // J  Yes / No dropdown
    'Expected Delivery',// K  date
    'Actual Delivery',  // L  date
    'Variance (Days)',  // M  ARRAYFORMULA = Actual - Expected (moved beside Actual)
    'Status',           // N  Not Ordered / PO Issued / In Transit / Delivered / Rejected / On Hold
    'Remarks'           // O
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);

  sh.getRange('M2').setFormula(`=ARRAYFORMULA(IF(OR(K2:K="",L2:L=""),"",IFERROR(L2:L-K2:K,"")))`);

  sh.getRange('K:L').setNumberFormat('dd/mm/yyyy');
  sh.setColumnWidth(1, 90); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 200);
  sh.setColumnWidth(4, 200); sh.setColumnWidth(6, 110); sh.setColumnWidth(9, 180); sh.setColumnWidth(15, 220);
}

function createProductionSheet_(ss) {
  try { CacheService.getScriptCache().remove(LAYOUT_CACHE_KEY); } catch (e) {}
  const sh = ss.insertSheet(SHEETS.PRODUCTION);
  const headers = [
    'Production ID',    // A  auto PROD-###
    'Project Label',    // B
    'Item',             // C
    'Category/Room',    // D  free text (no dropdown)
    'Quantity',         // E
    'Assigned To',      // F  dropdown from People
    'Start Date',       // G
    'Target Date',      // H  ARRAYFORMULA Handover-1 (beside Start)
    'Completion Date',  // I  user-entered when done
    'Variance (Days)',  // J  ARRAYFORMULA Completion - Target (beside Completion)
    'QC Status',        // K  Pending / Pass / Fail / Re-inspection
    'QC Date',          // L
    'Status',           // M  ARRAYFORMULA Not Started / In Progress / Completed
    'Remarks'           // N
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);

  const PR = SHEETS.PROJECTS;
  // Target Date = Projects.HandoverDate - 1 for this project (lookup by Label).
  sh.getRange('H2').setFormula(
    `=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,{${PR}!I:I,${PR}!F:F},2,FALSE)-1,"")))`
  );
  // Variance = Completion - Target
  sh.getRange('J2').setFormula(
    `=ARRAYFORMULA(IF(OR(H2:H="",I2:I=""),"",IFERROR(I2:I-H2:H,"")))`
  );
  // Status derived from Start / Completion
  sh.getRange('M2').setFormula(
    `=ARRAYFORMULA(IF(A2:A="","",IF(I2:I<>"","Completed",IF(G2:G<>"","In Progress","Not Started"))))`
  );

  sh.getRange('G:I').setNumberFormat('dd/mm/yyyy');
  sh.getRange('H:H').setNumberFormat('dd/mm/yyyy');
  sh.getRange('L:L').setNumberFormat('dd/mm/yyyy');
  sh.setColumnWidth(1, 90); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 220);
}

function createSnagsSheet_(ss) {
  const sh = ss.insertSheet(SHEETS.SNAGS);
  const headers = [
    'Snag ID',            // A  auto SNAG-###
    'Project Label',      // B
    'Description',        // C
    'Location',           // D
    'Priority',           // E
    'Assigned To',        // F
    'Status',             // G  Open / In Progress / Resolved
    'Raised Date',        // H
    'Target Resolve Date',// I
    'Resolved Date',      // J
    'Variance (Days)',    // K  ARRAYFORMULA Resolved - Target (beside Resolved)
    'Resolution Notes',   // L
    'Remarks'             // M
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('K2').setFormula(`=ARRAYFORMULA(IF(OR(I2:I="",J2:J=""),"",IFERROR(J2:J-I2:I,"")))`);
  sh.getRange('H:J').setNumberFormat('dd/mm/yyyy');
  sh.setColumnWidth(1, 90); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 260);
}

function createProjectViewSheet_(ss) {
  const sh = ss.insertSheet(SHEETS.VIEW);
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 240);
  sh.setColumnWidth(3, 30);  sh.setColumnWidth(4, 170); sh.setColumnWidth(5, 240);

  sh.getRange('A1').setValue('PROJECT VIEW');
  sh.getRange('A3').setValue('Select Project:');

  const PR = SHEETS.PROJECTS;
  // Label is col I now; MATCH against I:I.
  const idx = col => `=IFERROR(INDEX('${PR}'!${col}:${col},MATCH($B$3,'${PR}'!I:I,0)),"")`;

  const layout = [
    ['── PROJECT INFO ──', '', '', '', ''],
    ['Project ID:',     idx('A'), '', 'Client:',      idx('C')],
    ['Project Name:',   idx('B'), '', 'Location:',    idx('D')],
    ['Start Date:',     idx('E'), '', 'Handover:',    idx('F')],
    ['Overall Status:', idx('G'), '', 'Days Left:',   idx('K')],
    ['', '', '', '', ''],
    ['── PROGRESS ──', '', '', '', ''],
    ['Completion %:',   idx('Q'), '', 'Delay Status:', idx('J')],
    ['', '', '', '', ''],
    ['── STAGES ──', '', '', '', ''],
    ['Pre-Execution:', idx('L'), '', '', ''],
    ['Procurement:',   idx('M'), '', '', ''],
    ['Production:',    idx('N'), '', '', ''],
    ['Delivery:',      idx('O'), '', '', ''],
    ['Installation:',  idx('P'), '', '', ''],
    ['', '', '', '', ''],
    ['── COUNTS (per project) ──', '', '', '', ''],
    ['Purchases:',    `=COUNTIF('${SHEETS.PURCHASES}'!B:B,$B$3)`,  '', 'Delivered:', `=COUNTIFS('${SHEETS.PURCHASES}'!B:B,$B$3,'${SHEETS.PURCHASES}'!N:N,"Delivered")`],
    ['Production Items:', `=COUNTIF('${SHEETS.PRODUCTION}'!B:B,$B$3)`, '', 'Completed:', `=COUNTIFS('${SHEETS.PRODUCTION}'!B:B,$B$3,'${SHEETS.PRODUCTION}'!M:M,"Completed")`],
    ['Snags Total:',  `=COUNTIF('${SHEETS.SNAGS}'!B:B,$B$3)`,      '', 'Snags Open:',  `=COUNTIFS('${SHEETS.SNAGS}'!B:B,$B$3,'${SHEETS.SNAGS}'!G:G,"Open")`]
  ];
  sh.getRange(5, 1, layout.length, 5).setValues(layout);

  // Milestones table (row 30+)
  sh.getRange('A30').setValue('── MILESTONES ──');
  sh.getRange(31, 1, 1, 3).setValues([['#', 'Milestone', 'Status']]).setFontWeight('bold');
  const M = SHEETS.MILESTONES;
  const mRows = MILESTONES.map((name, i) => [
    i + 1, name,
    `=IFERROR(INDEX('${M}'!$C:$AA,MATCH($B$3,'${M}'!$A:$A,0),${i + 1}),"")`
  ]);
  sh.getRange(32, 1, 25, 3).setValues(mRows);

  // Purchases detail for the selected project (row 58+)
  sh.getRange('A58').setValue('── PURCHASES FOR THIS PROJECT ──');
  const PU = SHEETS.PURCHASES;
  // Purchases cols: A ID, B Project, C Material, D Spec, G Ordered Qty, H Unit, I Vendor, N Status
  sh.getRange('A59').setFormula(
    `=IF($B$3="","(pick a project in B3)",IFERROR(QUERY({'${PU}'!A2:D,'${PU}'!G2:H,'${PU}'!I2:I,'${PU}'!N2:N}, "SELECT Col1, Col3, Col4, Col5, Col6, Col7, Col8 WHERE Col2='"&$B$3&"' LABEL Col1 'ID', Col3 'Material', Col4 'Spec', Col5 'Qty', Col6 'Unit', Col7 'Vendor', Col8 'Status'", 1), "No purchases yet for this project"))`
  );

  // Formats
  sh.getRange('B8').setNumberFormat('dd/mm/yyyy');  // Start Date
  sh.getRange('E8').setNumberFormat('dd/mm/yyyy');  // Handover
  sh.getRange('B12').setNumberFormat('0.0%');       // Completion %
  sh.getRange('B22:B25').setNumberFormat('0');      // Counts
  sh.getRange('E22:E25').setNumberFormat('0');
  sh.setFrozenRows(3);
}

function createDashboardSheet_(ss) {
  const sh = ss.insertSheet(SHEETS.DASHBOARD);
  sh.getRange('A1').setValue('3SP DASHBOARD').setFontWeight('bold').setFontSize(14);

  const PR = SHEETS.PROJECTS, SN = SHEETS.SNAGS;

  // KPI stack (rows 2-14) + one report (rows 16-36)
  const kpi = [
    ['', ''],
    ['PROJECT HEALTH', ''],
    ['Active Projects',   `=COUNTIF('${PR}'!G:G,"In Progress")`],
    ['At Risk Projects',  `=COUNTIF('${PR}'!G:G,"At Risk")`],
    ['Completed Projects',`=COUNTIF('${PR}'!G:G,"Completed")`],
    ['Overdue Projects',  `=COUNTIF('${PR}'!J:J,"Overdue")`],
    ['Avg Completion %',  `=IFERROR(AVERAGE('${PR}'!Q2:Q${MAX_ROWS}),0)`],
    ['', ''],
    ['SNAGS', ''],
    ['', ''],
    ['', ''],
    ['Open Snags',      `=COUNTIF('${SN}'!G:G,"Open")`],
    ['Resolved Snags',  `=COUNTIF('${SN}'!G:G,"Resolved")`]
  ];
  for (let i = 0; i < kpi.length; i++) {
    sh.getRange(i + 2, 1).setValue(kpi[i][0]);
    if (kpi[i][1] && String(kpi[i][1]).startsWith('=')) sh.getRange(i + 2, 2).setFormula(kpi[i][1]);
    else                                                sh.getRange(i + 2, 2).setValue(kpi[i][1]);
  }

  // REPORT · Project Status (rows 16-36)
  sh.getRange('A16').setValue('REPORT · PROJECT STATUS (Top 20 by earliest handover)');
  sh.getRange('A17').setFormula(
    `=IFERROR(QUERY({'${PR}'!A2:B,'${PR}'!G2:G,'${PR}'!Q2:Q,'${PR}'!J2:J,'${PR}'!K2:K}, "SELECT * WHERE Col1 IS NOT NULL ORDER BY Col6 ASC LIMIT 20 LABEL Col1 'ID', Col2 'Name', Col3 'Status', Col4 '% Done', Col5 'Delay', Col6 'Days Left'", 0), "No projects yet.")`
  );
  sh.getRange('D18:D36').setNumberFormat('0.0%');

  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 130); sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 130); sh.setColumnWidth(6, 150);
}

// ── DATA VALIDATIONS ───────────────────────────────────────────────────────

function installDataValidations() {
  ensureTopRows_();   // the three table sheets first get their control rows, so every range below lands on real data rows
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const R = MAX_ROWS;
  const proj = ss.getRangeByName('PROJECT_LABELS');
  const ppl  = ss.getRangeByName('PEOPLE_NAMES');

  const dd = (sheetName, a1, source) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const b = SpreadsheetApp.newDataValidation();
    if (Array.isArray(source)) b.requireValueInList(source, true).setAllowInvalid(false);
    else                        b.requireValueInRange(source, true).setAllowInvalid(true);
    sh.getRange(shiftA1_(sh, a1)).setDataValidation(b.build());
  };
  const date = (sheetName, a1) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    sh.getRange(shiftA1_(sh, a1)).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());
  };

  // Projects (Owner col removed, Contract col removed)
  dd(SHEETS.PROJECTS, `G2:G${R}`, ['In Progress','At Risk','Completed']);
  date(SHEETS.PROJECTS, `E2:F${R}`);

  // Milestones — 25 milestone cols get Yes/No/Pending
  if (proj) dd(SHEETS.MILESTONES, `A2:A${R}`, proj);
  for (let c = 3; c <= 27; c++) {
    dd(SHEETS.MILESTONES, `${colLetter_(c)}2:${colLetter_(c)}${R}`, ['Yes','No','Pending']);
  }

  // Hardware List
  if (proj) dd(SHEETS.HARDWARE, `B2:B${R}`, proj);
  dd(SHEETS.HARDWARE, `H2:H${R}`, ['Sq.ft','Sq.m','Pcs','Box','Roll','Kg','Meter','Set','Other']);   // Unit (H: the Photo column sits at D)
  dd(SHEETS.HARDWARE, `K2:K${R}`, ['In Stock','Out of Stock']);                                          // Stock Status (K)

  // Purchases (Invoice# + Amount removed; PO Issued now Y/N; Variance moved; Status now col M)
  if (proj) dd(SHEETS.PURCHASES, `B2:B${R}`, proj);
  dd(SHEETS.PURCHASES, `H2:H${R}`, ['Sq.ft','Sq.m','Pcs','Box','Roll','Kg','Meter','Set','Other']);
  dd(SHEETS.PURCHASES, `J2:J${R}`, ['Yes','No']);
  dd(SHEETS.PURCHASES, `N2:N${R}`, ['Not Ordered','PO Issued','In Transit','Delivered','Rejected','On Hold']);
  date(SHEETS.PURCHASES, `K2:L${R}`);

  // Production (Category/Room is now free text; cols reshuffled)
  if (proj) dd(SHEETS.PRODUCTION, `B2:B${R}`, proj);
  if (ppl) dd(SHEETS.PRODUCTION, `F2:F${R}`, ppl);
  dd(SHEETS.PRODUCTION, `K2:K${R}`, ['Pending','Pass','Fail','Re-inspection']);
  date(SHEETS.PRODUCTION, `G2:G${R}`);
  date(SHEETS.PRODUCTION, `I2:I${R}`);
  date(SHEETS.PRODUCTION, `L2:L${R}`);

  // Snags (Variance moved to K; notes/remarks shifted)
  if (proj) dd(SHEETS.SNAGS, `B2:B${R}`, proj);
  dd(SHEETS.SNAGS, `E2:E${R}`, ['Low','Medium','High','Critical']);
  if (ppl) dd(SHEETS.SNAGS, `F2:F${R}`, ppl);
  dd(SHEETS.SNAGS, `G2:G${R}`, ['Open','In Progress','Resolved']);
  date(SHEETS.SNAGS, `H2:J${R}`);

  // Project View selector
  const view = ss.getSheetByName(SHEETS.VIEW);
  if (view && proj) view.getRange('B3').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(proj, true).setAllowInvalid(true).build()
  );

  Logger.log('Validations installed.');
}

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ── FORMATTING ─────────────────────────────────────────────────────────────

function installFormatting() {
  ensureTopRows_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = (name, cols) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const h = hdrRow_(sh), f = firstRow_(sh), n = Math.min(MAX_ROWS - 1, sh.getMaxRows() - f + 1);   // header row, first data row, how many rows
    sh.getRange(h, 1, 1, cols)
      .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG)
      .setFontWeight('bold').setFontSize(11)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sh.setRowHeight(h, 36);
    sh.getBandings().forEach(b => b.remove());
    sh.getRange(f, 1, n, cols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    sh.getRange(f, 1, n, cols).setVerticalAlignment('middle').setFontSize(10);
  };
  base(SHEETS.PROJECTS,   17);
  base(SHEETS.MILESTONES, 27);
  base(SHEETS.HARDWARE,   14);
  base(SHEETS.PURCHASES,  15);
  base(SHEETS.PRODUCTION, 14);
  base(SHEETS.SNAGS,      13);
  base(SHEETS.PEOPLE,     5);

  formatDashboard_(ss);
  formatProjectView_(ss);
  applyAllCF_(ss);
  applyProjectColorBands_(ss);
  Logger.log('Formatting installed.');
}

function formatDashboard_(ss) {
  const sh = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!sh) return;

  // Title bar
  sh.getRange('A1:F1').merge()
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG)
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 44);

  // Section header rows: PROJECT HEALTH (3), SNAGS (10), REPORT (16)
  [3, 10, 16].forEach(r => {
    sh.getRange(r, 1, 1, 6).merge()
      .setBackground(COLORS.SECTION_BG).setFontColor(COLORS.HEADER_FG)
      .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.setRowHeight(r, 30);
  });

  sh.getRange('A4:A8').setFontWeight('bold').setBackground('#F8F9FA');
  sh.getRange('A11:A14').setFontWeight('bold').setBackground('#F8F9FA');
  sh.getRange('B4:B14').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('B8').setNumberFormat('0.0%');          // Avg Completion
  sh.getRange('B7').setFontColor(COLORS.DANGER);      // Overdue Projects
}

function formatProjectView_(ss) {
  const sh = ss.getSheetByName(SHEETS.VIEW);
  if (!sh) return;
  sh.getRange('A1:E1').merge()
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG)
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 44);
  sh.getRange('A3').setFontWeight('bold').setBackground('#F8F9FA');
  sh.getRange('B3').setFontWeight('bold').setFontSize(12).setBackground('#FFF7E6');

  // Section header rows after Owner/Contract removed: Info row 5, Progress 11, Stages 14, Counts 21, Milestones 30, Purchases 58
  [5, 11, 14, 21, 30, 58].forEach(r => {
    sh.getRange(r, 1, 1, 5).merge()
      .setBackground(COLORS.SECTION_BG).setFontColor(COLORS.HEADER_FG)
      .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.setRowHeight(r, 30);
  });
  sh.getRange('A31:C31').setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange('A32:A56').setHorizontalAlignment('center').setFontWeight('bold');
}

function cfRules_(range, triples) {
  return triples.map(([value, bg, fg]) =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(value)
      .setBackground(bg).setFontColor(fg)
      .setBold(true)                 // status cells read like HUD badges
      .setRanges([range]).build()
  );
}

function applyAllCF_(ss) {
  const R = MAX_ROWS;
  // Projects: Overall Status (G), Delay Status (J)
  let sh = ss.getSheetByName(SHEETS.PROJECTS);
  if (sh) sh.setConditionalFormatRules([
    ...cfRules_(sh.getRange(`G2:G${R}`), [
      ['Completed',   COLORS.DONE_BG,     COLORS.DONE_FG],
      ['In Progress', COLORS.PROGRESS_BG, COLORS.PROGRESS_FG],
      ['At Risk',     COLORS.WARN_BG,     COLORS.WARN_FG]
    ]),
    ...cfRules_(sh.getRange(`J2:J${R}`), [
      ['On Track',      COLORS.DONE_BG, COLORS.DONE_FG],
      ['Near Deadline', COLORS.WARN_BG, COLORS.WARN_FG],
      ['Overdue',       COLORS.BAD_BG,  COLORS.BAD_FG]
    ])
  ]);
  // Milestones: Yes/No/Pending on C-AA
  sh = ss.getSheetByName(SHEETS.MILESTONES);
  if (sh) sh.setConditionalFormatRules(cfRules_(sh.getRange(`C2:AA${R}`), [
    ['Yes', COLORS.DONE_BG, COLORS.DONE_FG],
    ['No',  COLORS.BAD_BG,  COLORS.BAD_FG],
    ['Pending', COLORS.PENDING_BG, COLORS.PENDING_FG]
  ]));
  // Hardware List: Stock Status (K)
  sh = ss.getSheetByName(SHEETS.HARDWARE);
  if (sh) sh.setConditionalFormatRules(cfRules_(sh.getRange(shiftA1_(sh, `K2:K${R}`)), [
    ['In Stock',     COLORS.DONE_BG, COLORS.DONE_FG],
    ['Out of Stock', COLORS.BAD_BG,  COLORS.BAD_FG]
  ]));
  // Purchases: PO Issued (J), Status (N)
  sh = ss.getSheetByName(SHEETS.PURCHASES);
  if (sh) sh.setConditionalFormatRules([
    ...cfRules_(sh.getRange(shiftA1_(sh, `J2:J${R}`)), [
      ['Yes', COLORS.DONE_BG, COLORS.DONE_FG],
      ['No',  COLORS.PENDING_BG, COLORS.PENDING_FG]
    ]),
    ...cfRules_(sh.getRange(shiftA1_(sh, `N2:N${R}`)), [
      ['Delivered',   COLORS.DONE_BG,     COLORS.DONE_FG],
      ['In Transit',  COLORS.PROGRESS_BG, COLORS.PROGRESS_FG],
      ['PO Issued',   COLORS.PROGRESS_BG, COLORS.PROGRESS_FG],
      ['Not Ordered', COLORS.PENDING_BG,  COLORS.PENDING_FG],
      ['On Hold',     COLORS.WARN_BG,     COLORS.WARN_FG],
      ['Rejected',    COLORS.BAD_BG,      COLORS.BAD_FG]
    ])
  ]);
  // Production: Status (M), QC Status (K)
  sh = ss.getSheetByName(SHEETS.PRODUCTION);
  if (sh) sh.setConditionalFormatRules([
    ...cfRules_(sh.getRange(shiftA1_(sh, `M2:M${R}`)), [
      ['Completed',   COLORS.DONE_BG,     COLORS.DONE_FG],
      ['In Progress', COLORS.PROGRESS_BG, COLORS.PROGRESS_FG],
      ['Not Started', COLORS.PENDING_BG,  COLORS.PENDING_FG]
    ]),
    ...cfRules_(sh.getRange(shiftA1_(sh, `K2:K${R}`)), [
      ['Pass',          COLORS.DONE_BG,     COLORS.DONE_FG],
      ['Fail',          COLORS.BAD_BG,      COLORS.BAD_FG],
      ['Pending',       COLORS.PENDING_BG,  COLORS.PENDING_FG],
      ['Re-inspection', COLORS.WARN_BG,     COLORS.WARN_FG]
    ])
  ]);
  // Snags: Priority (E), Status (G)
  sh = ss.getSheetByName(SHEETS.SNAGS);
  if (sh) sh.setConditionalFormatRules([
    ...cfRules_(sh.getRange(`E2:E${R}`), [
      ['Critical', COLORS.BAD_BG,     COLORS.BAD_FG],
      ['High',     COLORS.WARN_BG,    COLORS.WARN_FG],
      ['Medium',   COLORS.PROGRESS_BG,COLORS.PROGRESS_FG],
      ['Low',      COLORS.DONE_BG,    COLORS.DONE_FG]
    ]),
    ...cfRules_(sh.getRange(`G2:G${R}`), [
      ['Resolved',    COLORS.DONE_BG,     COLORS.DONE_FG],
      ['In Progress', COLORS.PROGRESS_BG, COLORS.PROGRESS_FG],
      ['Open',        COLORS.WARN_BG,     COLORS.WARN_FG]
    ])
  ]);
}

/**
 * For every sheet with a Project Label column, install 8 CF rules that color
 * the row background based on MOD(project_number, 8). Same project → same
 * color across every sheet. Appended AFTER status/priority CF so those still
 * paint specific status cells (both bg + fg).
 */
function applyProjectColorBands_(ss) {
  const targets = [
    [SHEETS.MILESTONES, 'A', 'AA'],
    [SHEETS.HARDWARE,   'B', 'N'],
    [SHEETS.PURCHASES,  'B', 'O'],  // O (15) after In Stock (IMS) was added
    [SHEETS.PRODUCTION, 'B', 'N'],
    [SHEETS.SNAGS,      'B', 'M']
  ];
  const R = MAX_ROWS;
  targets.forEach(([sheetName, projCol, lastCol]) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    const existing = sh.getConditionalFormatRules();
    const f = firstRow_(sh), rowRange = sh.getRange(`A${f}:${lastCol}${Math.min(R + f - 2, sh.getMaxRows())}`);
    const colorRules = PROJECT_PALETTE.map((color, i) =>
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=AND($${projCol}${f}<>"", MOD(IFERROR(VALUE(REGEXEXTRACT($${projCol}${f}, "\\d+")),0), 8)=${i})`
        )
        .setBackground(color)
        .setRanges([rowRange])
        .build()
    );
    sh.setConditionalFormatRules(existing.concat(colorRules));
  });
}

// ── PROJECT FILTER (control rows above the table) ────────────────────────────────────────────────
// Hardware List / Purchases / Production: rows 1-2 hold a merged dropdown (C1:C2, "All projects" or one project) with the words "Project filter"
// beside it (B1:B2); the header is row 3, the data starts on row 4. Picking a project hides every other project's rows (a normal sheet filter on
// the Project column, so formulas, new rows and the scripts keep working); "All projects" shows everything again. Rows with no project stay
// visible so a new row can be typed. Typing in a new row while a project is picked fills its Project in.
function ensureTopRows_() {
  const cache = CacheService.getScriptCache();
  try { if (cache.get(LAYOUT_CACHE_KEY) === '1') return 0; } catch (e) {}
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return 0;
  const lock = LockService.getDocumentLock();
  try { if (!lock.tryLock(20000)) return 0; } catch (e) { return 0; }
  let moved = 0, ready = true;
  try {
    // cleanup: old PUR_* named ranges (PUR_MATERIAL, PUR_AMOUNT ...) point at a deleted area (#REF!); nothing uses them
    try { ss.getNamedRanges().forEach(function (nr) { if (/^PUR_/.test(nr.getName())) nr.remove(); }); } catch (e) {}
    FILTER_SHEETS.forEach(function (name) {
      const sh = ss.getSheetByName(name);
      if (!sh) return;
      if (String(sh.getRange(1, 1).getValue()).trim() === FILTER_HEADERS[name]) {   // still the old layout: the header is on row 1
        sh.insertRowsBefore(1, TOP_ROWS);                                              // native insert: formulas, validations, colours, banding move with their rows
        moved++;
      }
      if (String(sh.getRange(hdrRow_(sh), 1).getValue()).trim() !== FILTER_HEADERS[name]) { ready = false; return; }   // not a layout we know: leave it alone
      if (name === SHEETS.HARDWARE) ensurePhotoColumn_(sh);   // Hardware List only: the IMS photo between Material and Specification
      setupProjectFilter_(sh);
    });
    // heal: an older reseed left array-formula seeds on row 2 (the filter rows) of Purchases / Production: they show #REF!. Re-seed on the data row.
    try {
      const stale = [[SHEETS.PURCHASES, ['M']], [SHEETS.PRODUCTION, ['H', 'J', 'M']]].some(function (x) {
        const sh = ss.getSheetByName(x[0]);
        return !!sh && x[1].some(function (c) { return String(sh.getRange(c + '2').getFormula()).toUpperCase().indexOf('ARRAYFORMULA') >= 0; });
      });
      if (stale) reseedArrayFormulas_(ss);
    } catch (e) {}
  } finally { try { lock.releaseLock(); } catch (e) {} }
  if (ready) { try { cache.put(LAYOUT_CACHE_KEY, '1', 21600); } catch (e) {} }
  return moved;
}

/** Idempotent: the merged cells, the dropdown, its list, the frozen header and the sheet filter. */
function setupProjectFilter_(sh) {
  const ss = sh.getParent();
  let list = ss.getSheetByName(FILTER_LIST_SHEET);
  if (!list) {
    list = ss.insertSheet(FILTER_LIST_SHEET);
    list.getRange('A1').setValue(FILTER_ALL);
    list.getRange('A2').setFormula('=IFERROR(FILTER(PROJECT_LABELS,PROJECT_LABELS<>""),"")');
    list.hideSheet();
  }
  // The two rows were inserted above the header and took ITS look (dark blue, white text, and a dropdown copied from the cell beside).
  // Rows 1-2 are plain white; only the label and the dropdown are styled below.
  sh.getRange(1, 1, TOP_ROWS, Math.max(sh.getLastColumn(), FILTER_CELL_COL + 1)).clearDataValidations().setBackground('#FFFFFF').setFontColor('#000000').setFontWeight('normal');
  const lab = sh.getRange(1, FILTER_LABEL_COL, TOP_ROWS, 1), cell = sh.getRange(1, FILTER_CELL_COL, TOP_ROWS, 1);
  if (!lab.isPartOfMerge()) lab.merge();
  if (!cell.isPartOfMerge()) cell.merge();
  lab.setValue('Project filter ▶').setFontWeight('bold').setFontColor('#1F3864').setHorizontalAlignment('right').setVerticalAlignment('middle').setBackground('#EAF1FB');
  const cur = String(sh.getRange(1, FILTER_CELL_COL).getValue()).trim();
  cell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(list.getRange('A1:A500'), true).setAllowInvalid(false).build())
      .setFontWeight('bold').setFontSize(11).setFontColor('#000000').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#FFF2CC')
      .setBorder(true, true, true, true, null, null, '#BF9000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  if (!cur) sh.getRange(1, FILTER_CELL_COL).setValue(FILTER_ALL);
  sh.getRange(1, FILTER_CELL_COL).setNote('Pick a project to show only its rows. "All projects" shows everything. Rows with no project stay visible so you can add one.');
  sh.setRowHeights(1, TOP_ROWS, 24);
  if (sh.getFrozenRows() !== hdrRow_(sh)) sh.setFrozenRows(hdrRow_(sh));
  if (!sh.getFilter()) {
    const h = hdrRow_(sh);
    sh.getRange(h, 1, Math.max(sh.getMaxRows() - h + 1, 1), Math.max(sh.getLastColumn(), 1)).createFilter();
  }
}

/** Show only `label`'s rows (blank-project rows stay); '' / "All projects" removes the filter. */
function applyProjectFilter_(sh, label) {
  const f = sh.getFilter();
  if (!f) return;
  const want = String(label || '').trim();
  if (!want || want === FILTER_ALL) { f.removeColumnFilterCriteria(FILTER_PROJECT_COL); return; }
  const first = firstRow_(sh), last = sh.getLastRow(), seen = {}, hidden = [];
  if (last >= first) sh.getRange(first, FILTER_PROJECT_COL, last - first + 1, 1).getValues().forEach(function (r) {
    const v = r[0];
    if (v !== '' && v !== null && String(v) !== want && !seen[v]) { seen[v] = true; hidden.push(v); }
  });
  f.setColumnFilterCriteria(FILTER_PROJECT_COL, SpreadsheetApp.newFilterCriteria().setHiddenValues(hidden.map(String)).build());
}

/** Edit trigger: the dropdown changed -> filter. A typed cell in a new row while a project is picked -> that Project fills in. */
function handleProjectFilterEdit_(e) {
  const range = e.range, sh = range.getSheet();
  if (FILTER_SHEETS.indexOf(sh.getName()) < 0) return;
  const row = range.getRow(), col = range.getColumn();
  if (row <= TOP_ROWS) {
    if (col === FILTER_CELL_COL) applyProjectFilter_(sh, sh.getRange(1, FILTER_CELL_COL).getValue());
    return;
  }
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1 || col === FILTER_PROJECT_COL || col === 1 || row < firstRow_(sh)) return;
  const picked = String(sh.getRange(1, FILTER_CELL_COL).getValue()).trim();
  if (!picked || picked === FILTER_ALL || range.getValue() === '') return;
  const p = sh.getRange(row, FILTER_PROJECT_COL);
  if (p.getValue() === '') p.setValue(picked);
}

/** Hardware List: insert the Photo column between Material (C) and Specification, once. Everything from Specification on moves one to the right. */
function ensurePhotoColumn_(sh) {
  const h = hdrRow_(sh), cur = String(sh.getRange(h, IM_HW_PHOTO_COL).getValue()).trim();
  if (cur === IM_PHOTO_HDR) { cleanPhotoColumn_(sh); return false; }
  if (cur !== 'Specification') return false;   // not the layout we know: leave it alone
  sh.insertColumnAfter(IM_HW_PHOTO_COL - 1);   // native insert: formulas, dropdowns, checkboxes, colours move with their columns
  sh.getRange(h, IM_HW_PHOTO_COL).setValue(IM_PHOTO_HDR)
    .setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_FG).setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sh.setColumnWidth(IM_HW_PHOTO_COL, IM_PHOTO_COL_PX);
  cleanPhotoColumn_(sh);
  return true;
}

/** The inserted column copied Material's dropdown (and any label picked from it). The Photo column holds pictures only: no dropdown, no text. Pictures are never touched. */
function cleanPhotoColumn_(sh) {
  const f = firstRow_(sh), n = Math.max(sh.getMaxRows() - f + 1, 1), col = sh.getRange(f, IM_HW_PHOTO_COL, n, 1);
  col.clearDataValidations().setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(1, IM_HW_PHOTO_COL, TOP_ROWS, 1).clearDataValidations();
  col.getValues().forEach(function (r, i) { if (typeof r[0] === 'string' && r[0] !== '') sh.getRange(f + i, IM_HW_PHOTO_COL).clearContent(); });   // text only (a picture is an object)
}
