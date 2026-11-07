/**
 * _Interlink_PushProjects.gs  (PM only)
 * -----------------------------------------------------------------------------
 * Keeps Masters' Projects in step with PM's Projects sheet. Masters uses the SAME
 * PRJ-### as PM (passed in the payload), so every push is an upsert: same ID → update.
 *
 * It runs by itself:
 *   - right after ➕ Add Project (menuAddProject),
 *   - every time PM opens (installable onOpen, IM_onOpen_), so a missed push heals itself.
 * Only projects Masters lacks, or whose name / active flag differ, are sent, so opening PM
 * writes nothing when everything is already in step. Projects removed from PM are NOT
 * removed from Masters here (the cross-system Delete project does that).
 * 🔗 Interlink → "Send all projects to Masters" does the same on demand and reports.
 *
 * Depends on _Interlink_Masters.gs (Masters library + IM_clearCache).
 *
 * Mappings (PM column → Masters field):
 *   Project ID      → Project_ID        (upsert key)
 *   Project Name    → Name
 *   Client          → Customer_ID       (resolved via Masters.resolveFreeformName)
 *   Location        → Notes             (plain text; Masters.Sites is dropdown-only)
 *   Start Date      → Start_Date
 *   Handover Date   → Target_Handover
 *   Overall Status  → Status             (In Progress → In Progress,
 *                                         At Risk     → In Progress,
 *                                         Completed   → Delivered)
 *   Remarks         → Notes (joined with Location)
 *   Overall Status=Completed → Active?=No
 */

/** One PM project row (columns A–H) → Masters payload, or null for a blank row. */
function IM_projectPayload_(r, tz) {
  var projectId = String(r[0] || '').trim();
  var name      = String(r[1] || '').trim();
  if (!projectId || !name) return null;

  var client   = String(r[2] || '').trim();
  var location = String(r[3] || '').trim();
  var startDt  = r[4];
  var handover = r[5];
  var pmStatus = String(r[6] || '').trim();
  var remarks  = String(r[7] || '').trim();

  var statusMap = {
    'In Progress': 'In Progress',
    'At Risk':     'In Progress',   // closest active, see Masters SCHEMA
    'Completed':   'Delivered'
  };
  var customerId = '';
  if (client) {
    try {
      var hit = Masters.resolveFreeformName('customer', client);
      if (hit) customerId = hit.id;
    } catch (e) {}
  }
  var notes = [location && ('Location: ' + location), remarks].filter(Boolean).join(' · ');
  return {
    Project_ID:      projectId,
    Name:            name,
    Customer_ID:     customerId,
    Start_Date:      startDt  instanceof Date ? Utilities.formatDate(startDt,  tz, 'yyyy-MM-dd') : '',
    Target_Handover: handover instanceof Date ? Utilities.formatDate(handover, tz, 'yyyy-MM-dd') : '',
    Status:          statusMap[pmStatus] || 'Enquiry',
    Notes:           notes,
    'Active?':       pmStatus !== 'Completed' ? 'Yes' : 'No'
  };
}

/**
 * Send to Masters every PM project Masters lacks or has under another name / active flag.
 * Returns { pushed, skipped, errors[] }. Never shows UI.
 */
function IM_syncProjectsToMasters_() {
  var out = { pushed: 0, skipped: 0, errors: [] };
  var sh = SpreadsheetApp.getActive().getSheetByName('Projects');
  if (!sh || sh.getLastRow() < 2) return out;
  var tz = Session.getScriptTimeZone();
  var there = {};  // what Masters already has: PRJ-### → { name, active }
  (Masters.getProjects() || []).forEach(function (p) {
    there[String(p.Project_ID).toUpperCase()] = { name: String(p.Name), active: String(p['Active?'] || 'Yes') };
  });
  // Columns A–H are the user-editable project fields; I–Q are formulas we skip.
  sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().forEach(function (r, i) {
    var payload = IM_projectPayload_(r, tz);
    if (!payload) { out.skipped++; return; }
    var m = there[payload.Project_ID.toUpperCase()];
    if (m && m.name === payload.Name && m.active === payload['Active?']) return;   // already in step
    try { Masters.createProject(payload); out.pushed++; }                           // upsert on Project_ID
    catch (e) { out.errors.push('Row ' + (i + 2) + ' (' + payload.Project_ID + '): ' + e.message); }
  });
  if (out.pushed) { try { IM_clearCache(); } catch (e) {} }   // IM_getProjects() must see the change next call
  return out;
}

/** Menu / editor: send everything now and say what happened. */
function IM_pushAllProjectsToMasters() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActive().getSheetByName('Projects');
  if (!sh) { ui.alert('Projects sheet not found in PM.'); return; }
  if (sh.getLastRow() < 2) { ui.alert('No projects in PM to push.'); return; }
  var res = IM_syncProjectsToMasters_();
  ui.alert('Pushed to Masters:\n' +
    '  new or changed: ' + res.pushed + '\n' +
    '  blank rows (skipped): ' + res.skipped +
    (res.errors.length ? '\n\n  errors (' + res.errors.length + '):\n    ' + res.errors.slice(0, 10).join('\n    ') : '\n\nEverything else was already in step.'));
}
