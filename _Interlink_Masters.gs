/**
 * _Interlink_Masters.gs
 * -----------------------------------------------------------------------------
 * Thin wrapper around the Masters library (added as `Masters` dependency in
 * this script's Libraries). Caches reads in CacheService for 5 min so repeated
 * dropdown reads don't thrash the library.
 *
 * Paste this file verbatim into Finance / IMS / PM. It's self-contained —
 * no touches to any existing production code.
 *
 * One-time setup per system (do once after pasting this file):
 *   1. Script editor → Libraries → Add library
 *      Script ID: 1zWL2xvoZIZ1HAk-RrngKIeA4Ps4fdByYW9_hsMEm8lMPA2P8rOgTzSxC
 *      Version:   1   (or latest)
 *      Identifier: Masters
 *   2. Run IM_installInterlinkMenu() once from the editor — adds the
 *      🔗 Interlink menu to the sheet (via installable onOpen trigger, so
 *      the system's existing onOpen stays untouched).
 */

var IM_VERSION    = '0.4.0';
var IM_CACHE_TTL  = 300;   // 5 min
var IM_CACHE_KEY  = 'IM_v' + IM_VERSION + '_';

// -----------------------------------------------------------------------------
// Reads (cached)
// -----------------------------------------------------------------------------
function IM_getProjects()  { return IM_cache_('projects',  function () { return Masters.getProjects(); }); }
function IM_getCustomers() { return IM_cache_('customers', function () { return Masters.getCustomers(); }); }
function IM_getEmployees() { return IM_cache_('employees', function () { return Masters.getEmployees(); }); }
function IM_getSites()     { return IM_cache_('sites',     function () { return Masters.getSites(); }); }

/**
 * Parties with role flags. Call with a flag object to narrow:
 *   IM_getParties({ isSupplier:  true })  → suppliers only
 *   IM_getParties({ isJobWorker: true })  → job workers only
 *   IM_getParties({})                     → all parties
 */
function IM_getParties(filter) {
  var key = 'parties_' + JSON.stringify(filter || {});
  return IM_cache_(key, function () { return Masters.getParties(filter || {}); });
}

/** Freeform name → Masters record. For migration-time resolution. */
function IM_resolveName(entity, text) {
  return Masters.resolveFreeformName(entity, text);
}

// -----------------------------------------------------------------------------
// Dropdown-friendly label lists — "{ID} — {Name}"
// -----------------------------------------------------------------------------
function IM_projectLabels()  { return IM_asLabels_(IM_getProjects(),  'Project_ID',  'Name'); }
function IM_customerLabels() { return IM_asLabels_(IM_getCustomers(), 'Customer_ID', 'Name'); }
function IM_partyLabels(f)   { return IM_asLabels_(IM_getParties(f),  'Party_ID',    'Name'); }
function IM_employeeLabels() { return IM_asLabels_(IM_getEmployees(), 'Emp_ID',      'Name'); }
function IM_siteLabels()     { return IM_asLabels_(IM_getSites(),     'Site_ID',     'Name'); }

function IM_asLabels_(rows, idCol, nameCol) {
  return (rows || [])
    .filter(function (r) { return String(r['Active?'] || 'Yes').toLowerCase() !== 'no'; })
    .map(function (r) { return r[idCol] + ' — ' + r[nameCol]; });
}

// Just the IDs (for lookups where you don't want the label)
function IM_projectIds()  { return IM_getProjects().map(function (r)  { return r.Project_ID;  }); }
function IM_customerIds() { return IM_getCustomers().map(function (r) { return r.Customer_ID; }); }

// -----------------------------------------------------------------------------
// Cache plumbing
// -----------------------------------------------------------------------------
function IM_cache_(key, factory) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(IM_CACHE_KEY + key);
  if (hit) {
    try { return JSON.parse(hit); } catch (_) {}
  }
  var fresh = factory();
  try {
    var str = JSON.stringify(fresh);
    if (str.length < 95000) cache.put(IM_CACHE_KEY + key, str, IM_CACHE_TTL);
  } catch (_) {}
  return fresh;
}

function IM_clearCache() {
  var cache = CacheService.getScriptCache();
  cache.removeAll([
    IM_CACHE_KEY + 'projects',
    IM_CACHE_KEY + 'customers',
    IM_CACHE_KEY + 'employees',
    IM_CACHE_KEY + 'sites',
    IM_CACHE_KEY + 'parties_{}',
    IM_CACHE_KEY + 'parties_{"isSupplier":true}',
    IM_CACHE_KEY + 'parties_{"isJobWorker":true}',
    IM_CACHE_KEY + 'parties_{"isTransporter":true}',
    IM_CACHE_KEY + 'parties_{"isServiceProv":true}'
  ]);
  SpreadsheetApp.getActive().toast('Masters cache cleared', '🔗 Interlink', 3);
}

// -----------------------------------------------------------------------------
// Menu — attached via installable onOpen trigger so the system's existing
// simple onOpen stays untouched. Run IM_installInterlinkMenu() ONCE.
// -----------------------------------------------------------------------------
function IM_installInterlinkMenu() {
  var ss = SpreadsheetApp.getActive();
  // De-dup: remove any existing IM_onOpen_ trigger first
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'IM_onOpen_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('IM_onOpen_').forSpreadsheet(ss).onOpen().create();
  // Also run once now so the menu appears without reopening
  IM_onOpen_();
  SpreadsheetApp.getUi().alert(
    'Interlink menu installed',
    'A 🔗 Interlink menu will appear every time this sheet is opened. ' +
    'The system\'s existing menu is unchanged.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function IM_onOpen_() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔗 Interlink')
    .addItem('➕ Add item to IMS',             'IM_addItemDialog')
    .addItem('➕ Add employee',                'IM_addEmployeeDialog')
    .addItem('Refresh dropdowns from Masters', 'IM_refreshDropdowns')
    .addItem('Refresh Hardware List',          'IM_refreshHardwareNow')
    .addSeparator()
    .addItem('📋 Needs attention',             'IM_needsAttentionNow')
    .addItem('🩺 Health check',                'IM_healthCheckNow')
    .addSeparator()
    .addSubMenu(ui.createMenu('🧪 Self-test')
      .addItem('Run system self-test',         'IM_selfTestStart')
      .addItem('Show last self-test report',   'IM_selfTestShowReport')
      .addItem('Stop self-test and clean up',  'IM_selfTestStop'))
    .addSeparator()
    .addItem('❓ Start here',                  'IM_showStartHere')
    .addItem('About',                          'IM_about_')
    .addToUi();
  try { IM_removeDuplicateTriggers_(); } catch (e) {}       // a doubled edit trigger would run every edit twice
  try { IM_syncProjectsToMasters_(); } catch (e) {}         // a project added or renamed in PM reaches Masters (and IMS / Finance)
  try {                                                      // the People list (ticket dropdowns) is rewritten from Masters
    var pr = IM_syncPeopleFromMasters_();
    if (pr.blocked) SpreadsheetApp.getActive().toast('The People list was not refreshed: details typed in PM could not be copied into Masters first (' + pr.error + '). Open PM as a Masters editor once.', '🔗 Interlink', 12);
    else if (pr.removed.length) SpreadsheetApp.getActive().toast(pr.removed.length + ' name(s) that are not active employees in Masters were removed from the People list: ' + pr.removed.slice(0, 5).join(', ') + (pr.removed.length > 5 ? ', …' : ''), '🔗 Interlink', 12);
  } catch (e) {}
  try { IM_syncCustomerList_(); } catch (e) {}              // the Client dropdown on the Projects sheet (Finance's customers, via Masters)
  try { IM_syncVendorList_(); } catch (e) {}                // the Vendors sheet behind the vendor dropdowns (Masters' suppliers)
  try { IM_selfTestAnnounce_(); } catch (e) {}              // how the last self-test ended
  try { IM_installHelpNotes_(); } catch (e) {}              // explanations on the Hardware List headers
  try { IM_healFormulaColumns_(); } catch (e) {}            // an automatic formula someone typed over comes back
  try { IM_needsAttentionNow(true); } catch (e) {}          // refresh the "Needs attention" list; a toast only if something is red
  try { IM_refreshReservationStatus(true); } catch (e) {}  // keep Hardware List col M current on open
  try { IM_refreshStockStatus_(null); } catch (e) {}        // keep Stock Status (col J) of unticked rows current on open
}

function IM_about_() {
  var pv = '?';
  try { pv = Masters.getVersion(); } catch (e) {}
  SpreadsheetApp.getUi().alert(
    '🔗 Interlink v' + IM_VERSION,
    'This system reads master data (Projects, Customers, Parties, Employees, ' +
    'Sites) from the Masters library.\n\n' +
    'Masters version: ' + pv + '\n' +
    'Cache TTL: ' + IM_CACHE_TTL + ' sec',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// -----------------------------------------------------------------------------
// Employees from Masters (item 8, test run). Masters' Employees is the master list. PM's People sheet (the
// ticket "Raised By / Assigned To" dropdowns) is kept in step ADDITIVELY: an active Masters employee that the
// list lacks is appended; nothing already there is changed or removed. Names are compared ignoring case,
// punctuation and extra spaces, and never fuzzily ("Arjun" is not "Arjun Singh"). Differences the other way
// (a name in PM / IMS / Finance that Masters does not have, or an employee Masters marks as Left) are listed in
// 📋 Needs attention for a person to decide.
// -----------------------------------------------------------------------------
function IM_normName_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }

/**
 * The People sheet is GENERATED from Masters' Employees: it is rewritten completely each time (name, department, role, phone, email
 * of every ACTIVE Masters employee, A to Z). It exists only because the in-sheet dropdowns (Raised By / Assigned To) need a range to
 * read; it is not a list to edit, and names that are not active in Masters disappear from it (tickets already raised keep their text).
 *
 * One-time safety: if the sheet still holds a department / role / phone / email for someone that Masters has blank, that detail is
 * first copied INTO Masters (blank fields only; needs a Masters editor). If that cannot be done, the sheet is left exactly as it is
 * and { blocked: true } says so. Returns { rows, added:[names], removed:[names], moved, blocked, error, overflow }.
 */
function IM_syncPeopleFromMasters_(fresh) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.PEOPLE);
  var res = { rows: 0, added: [], removed: [], moved: 0, blocked: false, overflow: false };
  if (!sh) return res;
  CacheService.getScriptCache().remove(IM_CACHE_KEY + 'employees');   // people are always read fresh: someone added in Masters a minute ago must show up now
  var all = IM_getEmployees() || [];
  if (!all.length) return res;          // Masters returned nothing: never wipe the list
  var last = sh.getLastRow();
  var cur = last >= 2 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
  var txt = function (v) { return String(v == null ? '' : v).trim(); };
  var byKey = {}; all.forEach(function (e) { var k = IM_normName_(e.Name); if (k && !byKey[k]) byKey[k] = e; });
  // 1. details typed here that Masters lacks go up first
  var up = [];
  cur.forEach(function (r) {
    var e = byKey[IM_normName_(r[0])]; if (!e) return;
    var u = { Name: e.Name };
    if (txt(e.Department) === '' && txt(r[1]) !== '') u.Department = txt(r[1]);
    if (txt(e.Role) === '' && txt(r[2]) !== '') u.Role = txt(r[2]);
    if (txt(e.Phone) === '' && txt(r[3]) !== '') u.Phone = txt(r[3]);
    if (txt(e.Email) === '' && txt(r[4]) !== '') u.Email = txt(r[4]);
    if (Object.keys(u).length > 1) up.push(u);
  });
  if (up.length) {
    try {
      res.moved = Masters.fillEmployeeBlanks(up).filled;
      CacheService.getScriptCache().remove(IM_CACHE_KEY + 'employees');
      all = IM_getEmployees() || all;
    } catch (e) {
      res.blocked = true; res.error = IM_friendlyText_ ? IM_friendlyText_(e.message) : e.message; res.pending = up.length;
      return res;                      // leave the sheet exactly as it is
    }
  }
  // 2. rewrite from Masters
  var seen = {}, rows = [];
  all.forEach(function (e) {
    var n = txt(e.Name).replace(/\s+/g, ' '), k = IM_normName_(n);
    if (!k || seen[k] || String(e.Status) === 'Left') return; seen[k] = true;
    rows.push([n, txt(e.Department), txt(e.Role), txt(e.Phone), txt(e.Email)]);
  });
  rows.sort(function (a, b) { return a[0].toLowerCase() < b[0].toLowerCase() ? -1 : a[0].toLowerCase() > b[0].toLowerCase() ? 1 : 0; });
  var had = {}; cur.forEach(function (r) { var k = IM_normName_(r[0]); if (k) had[k] = txt(r[0]); });
  var now = {}; rows.forEach(function (r) { now[IM_normName_(r[0])] = true; });
  rows.forEach(function (r) { if (!had[IM_normName_(r[0])]) res.added.push(r[0]); });
  Object.keys(had).forEach(function (k) { if (!now[k]) res.removed.push(had[k]); });
  var same = cur.length >= rows.length && rows.every(function (r, i) { return r.every(function (v, j) { return txt(cur[i][j]) === v; }); }) &&
    cur.slice(rows.length).every(function (r) { return r.every(function (v) { return txt(v) === ''; }); });
  if (!same) {
    if (last >= 2) sh.getRange(2, 1, Math.max(last - 1, 1), 5).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, 5).setValues(rows);
  }
  try { sh.getRange('A1').setNote('Generated from Masters (Employees). Do not type here: this sheet is rewritten automatically. Add or change people in Masters (🔗 Interlink → ➕ Add employee).'); } catch (e) {}
  res.rows = rows.length; res.overflow = rows.length + 1 > 100;   // the dropdown range is A2:A100
  return res;
}

/**
 * CUSTOMERS: Finance is the master (Masters keeps a copy). The Projects sheet's Client column gets a dropdown fed by a hidden, GENERATED
 * sheet 'Customers' (active Masters customers, A to Z). Typing a client that is not in the list is still allowed (it only shows a warning);
 * the list is rewritten completely each time, never wiped when Masters returns nobody. Returns { rows, added, removed } or null.
 */
function IM_syncCustomerList_(fresh) {
  var ss = SpreadsheetApp.getActive(), proj = ss.getSheetByName(SHEETS.PROJECTS);
  if (!proj) return null;
  if (fresh) { try { CacheService.getScriptCache().remove(IM_CACHE_KEY + 'customers'); } catch (e) {} }
  var names = [], seen = {};
  (IM_getCustomers() || []).forEach(function (c) {
    var n = String(c.Name == null ? '' : c.Name).replace(/\s+/g, ' ').trim(), k = IM_normName_(n);
    if (!k || seen[k] || String(c['Active?'] || 'Yes').toLowerCase() === 'no') return; seen[k] = true; names.push(n);
  });
  if (!names.length) return null;
  names.sort(function (a, b) { return a.toLowerCase() < b.toLowerCase() ? -1 : a.toLowerCase() > b.toLowerCase() ? 1 : 0; });
  var sh = ss.getSheetByName('Customers'), made = false;
  if (!sh) { sh = ss.insertSheet('Customers'); made = true; sh.hideSheet(); }
  var last = sh.getLastRow(), cur = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }).filter(String) : [];
  var res = { rows: names.length, added: names.filter(function (n) { return cur.indexOf(n) < 0; }).length, removed: cur.filter(function (n) { return names.indexOf(n) < 0; }).length };
  if (made || res.added || res.removed) {
    if (last >= 2) sh.getRange(2, 1, last - 1, 1).clearContent();
    sh.getRange(1, 1).setValue('Customer Name').setFontWeight('bold');
    sh.getRange(2, 1, names.length, 1).setValues(names.map(function (n) { return [n]; }));
    try { sh.getRange('A1').setNote('Generated from Masters (Customers, copied from Finance). Do not type here: rewritten automatically.'); } catch (e) {}
  }
  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(sh.getRange(2, 1, Math.max(names.length, 1), 1), true).setAllowInvalid(true).build();
  proj.getRange(2, 3, Math.max(MAX_ROWS - 1, 1), 1).setDataValidation(rule);   // Projects: Client (col C)
  return res;
}

/**
 * VENDORS: Masters (Parties flagged Supplier) is the master list. The 'Vendors' sheet (the dropdowns on Purchases and Tickets read it) is a
 * GENERATED COPY: active Masters suppliers A to Z (name, contact, category, phone, email). A vendor that only PM knows (typed here by hand,
 * not in Masters) is KEPT below them, never deleted; it is named in the toast so someone can add it to Masters.
 * Returns { masters:[rows], onlyPm:[rows] } (rows = [name, contact, category, phone, email]) or null when Masters returned no supplier.
 */
function IM_vendorRows_() {
  var t = function (v) { return String(v == null ? '' : v).trim(); }, seen = {}, all = {}, masters = [];
  (IM_getParties({ isSupplier: true }) || []).forEach(function (p) {
    var n = t(p.Name).replace(/\s+/g, ' '), k = IM_normName_(n);
    if (!k) return;
    all[k] = true;
    if (seen[k] || String(p['Active?'] || 'Yes').toLowerCase() === 'no') return; seen[k] = true;
    masters.push([n, t(p.Contact_Person), t(p.Category), t(p.Phone), t(p.Email)]);
  });
  if (!masters.length) return null;
  masters.sort(function (a, b) { return a[0].toLowerCase() < b[0].toLowerCase() ? -1 : a[0].toLowerCase() > b[0].toLowerCase() ? 1 : 0; });
  var onlyPm = [], sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.VENDORS), last = sh ? sh.getLastRow() : 0;
  if (last >= 2) sh.getRange(2, 1, last - 1, 5).getValues().forEach(function (r) {
    var k = IM_normName_(r[0]); if (k && !all[k]) { all[k] = true; onlyPm.push(r.map(t)); }   // not in Masters at all (active or not)
  });
  return { masters: masters, onlyPm: onlyPm };
}

/** Rewrites the Vendors sheet from Masters and points the vendor dropdowns (Purchases col I, Tickets col N) at it. Returns { rows, added, removed, onlyPm:[names] } or null. */
function IM_syncVendorList_(fresh) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(SHEETS.VENDORS);
  if (!sh) return null;
  if (fresh) { try { CacheService.getScriptCache().remove(IM_CACHE_KEY + 'parties_{"isSupplier":true}'); } catch (e) {} }
  var v = IM_vendorRows_();
  if (!v) return null;                                  // Masters returned nobody: never wipe the list
  var rows = v.masters.concat(v.onlyPm), last = sh.getLastRow();
  var cur = last >= 2 ? sh.getRange(2, 1, last - 1, 5).getValues().map(function (r) { return r.map(function (x) { return String(x).trim(); }); }) : [];
  var same = cur.length >= rows.length && rows.every(function (r, i) { return r.every(function (x, j) { return cur[i][j] === x; }); }) &&
    cur.slice(rows.length).every(function (r) { return r.every(function (x) { return x === ''; }); });
  var curNames = {}; cur.forEach(function (r) { if (r[0]) curNames[IM_normName_(r[0])] = true; });
  var now = {}; rows.forEach(function (r) { now[IM_normName_(r[0])] = true; });
  var res = { rows: rows.length, added: rows.filter(function (r) { return !curNames[IM_normName_(r[0])]; }).length,
              removed: Object.keys(curNames).filter(function (k) { return !now[k]; }).length, onlyPm: v.onlyPm.map(function (r) { return r[0]; }) };
  if (!same) {
    if (sh.getMaxRows() < rows.length + 1) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
    if (last >= 2) sh.getRange(2, 1, last - 1, 5).clearContent();
    sh.getRange(2, 1, rows.length, 5).setValues(rows);
    try { sh.getRange('A1').setNote('Generated from Masters (suppliers). Do not type here: this sheet is rewritten automatically. A vendor not in Masters is kept at the bottom; add it in Masters to make it official.'); } catch (e) {}
  }
  if (!same || fresh) {                                 // point the dropdowns at A2:A1000 (the old range stopped at row 100)
    var rng = sh.getRange(2, 1, 999, 1);
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(rng, true).setAllowInvalid(true).build();
    [[SHEETS.PURCHASES, 9], [SHEETS.TICKETS, 14]].forEach(function (x) {
      var s2 = ss.getSheetByName(x[0]); if (s2) s2.getRange(2, x[1], Math.max(MAX_ROWS - 1, 1), 1).setDataValidation(rule);
    });
    try { ss.getNamedRanges().forEach(function (nr) { if (nr.getName() === 'VENDOR_NAMES') nr.remove(); }); ss.setNamedRange('VENDOR_NAMES', rng); } catch (e) {}
  }
  return res;
}

/** Menu: 🔗 Refresh dropdowns from Masters. Projects and employees, with a plain summary. */
function IM_refreshDropdowns() {
  var ss = SpreadsheetApp.getActive(), parts = [];
  try { var p = IM_syncProjectsToMasters_(); parts.push(p.pushed ? p.pushed + ' project(s) sent to Masters' : 'projects already in step'); }
  catch (e) { parts.push('projects: ' + IM_friendlyText_(e.message)); }
  try {
    var e2 = IM_syncPeopleFromMasters_(true);
    if (e2.blocked) parts.push('People list NOT refreshed: details typed in PM could not be copied into Masters (' + e2.error + ')');
    else {
      parts.push(e2.added.length || e2.removed.length ? 'People list refreshed from Masters (' + e2.rows + ' people' + (e2.added.length ? ', ' + e2.added.length + ' added' : '') + (e2.removed.length ? ', ' + e2.removed.length + ' not in Masters removed: ' + e2.removed.slice(0, 4).join(', ') + (e2.removed.length > 4 ? ', …' : '') : '') + ')' : 'People list already matches Masters (' + e2.rows + ' people)');
      if (e2.overflow) parts.push('more than 99 people: the dropdown range (A2:A100) is too short, ask your administrator');
    }
  } catch (e) { parts.push('employees: ' + IM_friendlyText_(e.message)); }
  try {
    var cl = IM_syncCustomerList_(true);
    if (cl) parts.push('Client list: ' + cl.rows + ' customers' + (cl.added || cl.removed ? ' (' + cl.added + ' added, ' + cl.removed + ' removed)' : ''));
  } catch (e) { parts.push('customers: ' + IM_friendlyText_(e.message)); }
  try {
    var vl = IM_syncVendorList_(true);
    if (vl) parts.push('Vendor list: ' + vl.rows + ' vendors' + (vl.added || vl.removed ? ' (' + vl.added + ' added, ' + vl.removed + ' removed)' : '') +
      (vl.onlyPm.length ? ' · ' + vl.onlyPm.length + ' only in PM, not in Masters: ' + vl.onlyPm.slice(0, 4).join(', ') + (vl.onlyPm.length > 4 ? ', …' : '') : ''));
  } catch (e) { parts.push('vendors: ' + IM_friendlyText_(e.message)); }
  ss.toast(parts.join(' · '), '🔗 Interlink', 10);
}

// -----------------------------------------------------------------------------
// ➕ Add employee. The person is created in Masters (the one official list) and this system's own list is updated at
// once. Masters editors only (every Masters write is). Masters refuses an exact duplicate and warns on a similar name.
// -----------------------------------------------------------------------------
function IM_addEmployeeDialog() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(IM_addEmployeeHtml_()).setWidth(430).setHeight(500), '➕ Add employee');
}
function IM_addEmployeeHtml_() { return "<div style=\"font:14px Arial,sans-serif;padding:14px;color:#222\">\n<p style=\"margin:0 0 12px;color:#555\">The person is saved in <b>Masters</b>, the official employee list. The other systems pick the name up from there.</p>\n<label>Full name *</label><input id=\"n\" maxlength=\"80\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<label>Role (optional)</label><input id=\"r\" maxlength=\"60\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<label>Department (optional)</label><input id=\"d\" maxlength=\"60\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<label>Phone (optional)</label><input id=\"p\" maxlength=\"30\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<div id=\"m\" style=\"min-height:42px;margin:6px 0;font-size:13px\"></div>\n<button id=\"go\" onclick=\"send(false)\" style=\"padding:8px 16px;background:#1F4E79;color:#fff;border:0;border-radius:4px;cursor:pointer\">Add employee</button>\n<button id=\"any\" onclick=\"send(true)\" style=\"display:none;padding:8px 16px;background:#b45309;color:#fff;border:0;border-radius:4px;cursor:pointer\">Add anyway (different person)</button>\n<button onclick=\"google.script.host.close()\" style=\"padding:8px 16px;margin-left:6px;border:1px solid #bbb;background:#fff;border-radius:4px;cursor:pointer\">Close</button>\n</div>\n<script>\nfunction $(i){return document.getElementById(i);}\nfunction say(t,c){var m=$('m');m.textContent=t;m.style.color=c||'#222';}\nfunction busy(b){$('go').disabled=b;$('any').disabled=b;}\nfunction send(confirm){\n  var name=$('n').value.replace(/\\s+/g,' ').trim();\n  if(!name){say('Please type the name.','#b91c1c');return;}\n  busy(true);say('Saving\u2026');\n  google.script.run.withSuccessHandler(done).withFailureHandler(function(e){busy(false);say('Could not save: '+(e&&e.message||e),'#b91c1c');})\n    .IM_addEmployee({name:name,role:$('r').value,department:$('d').value,phone:$('p').value,confirmSimilar:confirm===true});\n}\nfunction done(r){\n  busy(false);$('any').style.display='none';\n  if(r&&r.ok){say('Added '+r.name+' ('+r.id+'). '+(r.note||''),'#166534');$('go').style.display='none';$('n').disabled=$('r').disabled=$('d').disabled=$('p').disabled=true;}\n  else if(r&&r.duplicate){say(r.duplicate.name+' is already in Masters ('+r.duplicate.id+', '+r.duplicate.status+'). Nothing was added.','#b45309');}\n  else if(r&&r.similar){say('Similar name(s) already in Masters: '+r.similar.map(function(s){return s.name+' ('+s.id+', '+s.status+')';}).join('; ')+'. If this is the same person, close this window. If it is a different person, click Add anyway.','#b45309');$('any').style.display='inline-block';}\n  else{say((r&&r.error)||'Something went wrong.','#b91c1c');}\n}\n</script>"; }
/** Called by the dialog. Returns { ok, id, name, note } | { ok:false, duplicate | similar | error }. Never throws. */
function IM_addEmployee(d) {
  try {
    var r = Masters.addEmployeeChecked({ Name: d.name, Role: d.role, Department: d.department, Phone: d.phone, confirmSimilar: d.confirmSimilar === true });
    if (r.ok) { try { IM_clearCache(); } catch (e) {} r.note = IM_addEmployeePost_(r, d); }
    return r;
  } catch (e) {
    var m = String((e && e.message) || e);
    return { ok: false, error: /Editors only|Admin only|permission|access/i.test(m)
      ? 'Only Masters editors can add employees. Ask the administrator to add you as an editor, or ask them to add this person.' : m };
  }
}
/** PM: the name goes into the People list straight away (ticket Raised By / Assigned To). */
function IM_addEmployeePost_(r, d) {
  try {
    var x = IM_syncPeopleFromMasters_(true);
    if (x.blocked) return 'Saved in Masters. The People list is refreshed the next time PM is opened.';
    return r.name + ' now appears in the People list (ticket Raised By / Assigned To). Payroll: the accountant adds the salary in Finance.';
  } catch (e) { return 'Saved in Masters. The People list updates the next time PM is opened.'; }
}
