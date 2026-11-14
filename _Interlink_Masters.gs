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
  try { IM_syncPeopleFromMasters_(); } catch (e) {}         // new Masters employees appear in the People list (ticket dropdowns)
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

/** Masters employees: { active: [{name, role}], left: {norm: true}, all: {norm: true} }. Throws if Masters cannot be read. */
function IM_mastersEmployees_() {
  var out = { active: [], left: {}, all: {} };
  (IM_getEmployees() || []).forEach(function (r) {
    var n = String(r.Name || '').trim();
    if (!n) return;
    var k = IM_normName_(n);
    out.all[k] = true;
    if (String(r.Status) === 'Left') out.left[k] = true;
    else out.active.push({ name: n, role: String(r.Role || '').trim() });
  });
  return out;
}

/** Append the active Masters employees PM's People list does not have yet. Returns { added, names, overflow }. */
function IM_syncPeopleFromMasters_(fresh) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.PEOPLE);
  if (!sh) return { added: 0, names: [], overflow: false };
  if (fresh) { try { IM_clearCache(); } catch (e) {} }
  var emp = IM_mastersEmployees_(), have = {}, last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { if (String(r[0]).trim() !== '') have[IM_normName_(r[0])] = true; });
  var add = emp.active.filter(function (e) { return !have[IM_normName_(e.name)]; });
  if (!add.length) return { added: 0, names: [], overflow: false };
  var start = firstEmptyRow_(sh, 1);
  sh.getRange(start, 1, add.length, 3).setValues(add.map(function (e) { return [e.name, '', e.role]; }));   // Name | Department (blank) | Role
  return { added: add.length, names: add.map(function (e) { return e.name; }), overflow: start + add.length - 1 > 100 };   // the dropdown range is A2:A100
}

/** Menu: 🔗 Refresh dropdowns from Masters. Projects and employees, with a plain summary. */
function IM_refreshDropdowns() {
  var ss = SpreadsheetApp.getActive(), parts = [];
  try { var p = IM_syncProjectsToMasters_(); parts.push(p.pushed ? p.pushed + ' project(s) sent to Masters' : 'projects already in step'); }
  catch (e) { parts.push('projects: ' + IM_friendlyText_(e.message)); }
  try {
    var e2 = IM_syncPeopleFromMasters_(true);
    parts.push(e2.added ? e2.added + ' employee(s) added to the People list (' + e2.names.slice(0, 4).join(', ') + (e2.names.length > 4 ? ', …' : '') + ')' : 'People list already has every active employee');
    if (e2.overflow) parts.push('the People list is longer than its dropdown range (A2:A100): ask your administrator');
  } catch (e) { parts.push('employees: ' + IM_friendlyText_(e.message)); }
  ss.toast(parts.join(' · '), '🔗 Interlink', 10);
}

// -----------------------------------------------------------------------------
// ➕ Add employee. The person is created in Masters (the one official list) and this system's own list is updated at
// once. Masters editors only (every Masters write is). Masters refuses an exact duplicate and warns on a similar name.
// -----------------------------------------------------------------------------
function IM_addEmployeeDialog() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(IM_addEmployeeHtml_()).setWidth(430).setHeight(430), '➕ Add employee');
}
function IM_addEmployeeHtml_() { return "<div style=\"font:14px Arial,sans-serif;padding:14px;color:#222\">\n<p style=\"margin:0 0 12px;color:#555\">The person is saved in <b>Masters</b>, the official employee list. The other systems pick the name up from there.</p>\n<label>Full name *</label><input id=\"n\" maxlength=\"80\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<label>Role (optional)</label><input id=\"r\" maxlength=\"60\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<label>Phone (optional)</label><input id=\"p\" maxlength=\"30\" style=\"width:100%;padding:7px;margin:3px 0 10px;box-sizing:border-box\">\n<div id=\"m\" style=\"min-height:42px;margin:6px 0;font-size:13px\"></div>\n<button id=\"go\" onclick=\"send(false)\" style=\"padding:8px 16px;background:#1F4E79;color:#fff;border:0;border-radius:4px;cursor:pointer\">Add employee</button>\n<button id=\"any\" onclick=\"send(true)\" style=\"display:none;padding:8px 16px;background:#b45309;color:#fff;border:0;border-radius:4px;cursor:pointer\">Add anyway (different person)</button>\n<button onclick=\"google.script.host.close()\" style=\"padding:8px 16px;margin-left:6px;border:1px solid #bbb;background:#fff;border-radius:4px;cursor:pointer\">Close</button>\n</div>\n<script>\nfunction $(i){return document.getElementById(i);}\nfunction say(t,c){var m=$('m');m.textContent=t;m.style.color=c||'#222';}\nfunction busy(b){$('go').disabled=b;$('any').disabled=b;}\nfunction send(confirm){\n  var name=$('n').value.replace(/\\s+/g,' ').trim();\n  if(!name){say('Please type the name.','#b91c1c');return;}\n  busy(true);say('Saving\u2026');\n  google.script.run.withSuccessHandler(done).withFailureHandler(function(e){busy(false);say('Could not save: '+(e&&e.message||e),'#b91c1c');})\n    .IM_addEmployee({name:name,role:$('r').value,phone:$('p').value,confirmSimilar:confirm===true});\n}\nfunction done(r){\n  busy(false);$('any').style.display='none';\n  if(r&&r.ok){say('Added '+r.name+' ('+r.id+'). '+(r.note||''),'#166534');$('go').style.display='none';$('n').disabled=$('r').disabled=$('p').disabled=true;}\n  else if(r&&r.duplicate){say(r.duplicate.name+' is already in Masters ('+r.duplicate.id+', '+r.duplicate.status+'). Nothing was added.','#b45309');}\n  else if(r&&r.similar){say('Similar name(s) already in Masters: '+r.similar.map(function(s){return s.name+' ('+s.id+', '+s.status+')';}).join('; ')+'. If this is the same person, close this window. If it is a different person, click Add anyway.','#b45309');$('any').style.display='inline-block';}\n  else{say((r&&r.error)||'Something went wrong.','#b91c1c');}\n}\n</script>"; }
/** Called by the dialog. Returns { ok, id, name, note } | { ok:false, duplicate | similar | error }. Never throws. */
function IM_addEmployee(d) {
  try {
    var r = Masters.addEmployeeChecked({ Name: d.name, Role: d.role, Phone: d.phone, confirmSimilar: d.confirmSimilar === true });
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
  try { IM_syncPeopleFromMasters_(true); return r.name + ' now appears in the People list (ticket Raised By / Assigned To). Payroll: the accountant adds the salary in Finance.'; }
  catch (e) { return 'Saved in Masters. The People list updates the next time PM is opened.'; }
}
