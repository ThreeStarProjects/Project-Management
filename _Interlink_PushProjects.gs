/**
 * _Interlink_PushProjects.gs  (PM only)
 * -----------------------------------------------------------------------------
 * Reads PM's Projects sheet and upserts every project into Masters.
 * Masters uses the SAME PRJ-### as PM — passed in the payload, so repeated
 * pushes are idempotent (same ID → update, not duplicate).
 *
 * Depends on _Interlink_Masters.gs being pasted first (for the Masters library
 * dependency and IM_clearCache cache-bust).
 *
 * Usage: run IM_pushAllProjectsToMasters() from the Script editor's Run menu,
 * or wire it to a menu item once you want it one-click.
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

function IM_pushAllProjectsToMasters() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Projects');
  if (!sh) {
    SpreadsheetApp.getUi().alert('Projects sheet not found in PM.');
    return;
  }
  var last = sh.getLastRow();
  if (last < 2) {
    SpreadsheetApp.getUi().alert('No projects in PM to push.');
    return;
  }

  // Columns A–H are the user-editable project fields; I–Q are formulas we skip.
  var rows = sh.getRange(2, 1, last - 1, 8).getValues();

  var pushed = 0, skipped = 0, errors = [];
  var tz = Session.getScriptTimeZone();

  rows.forEach(function (r, i) {
    var projectId = String(r[0] || '').trim();
    var name      = String(r[1] || '').trim();
    if (!projectId || !name) { skipped++; return; }

    var client   = String(r[2] || '').trim();
    var location = String(r[3] || '').trim();
    var startDt  = r[4];
    var handover = r[5];
    var pmStatus = String(r[6] || '').trim();
    var remarks  = String(r[7] || '').trim();

    // Status map
    var statusMap = {
      'In Progress': 'In Progress',
      'At Risk':     'In Progress',   // closest active, see Masters SCHEMA
      'Completed':   'Delivered'
    };
    var mStatus = statusMap[pmStatus] || 'Enquiry';
    var isActive = pmStatus !== 'Completed';

    // Resolve Client name → Customer_ID (blank if no match)
    var customerId = '';
    if (client) {
      try {
        var hit = Masters.resolveFreeformName('customer', client);
        if (hit) customerId = hit.id;
      } catch (e) {}
    }

    // Date coercion (Masters CRUD coerces again on save, but formatting here
    // keeps Date objects from being lost over google.script.run boundaries)
    var startStr    = startDt  instanceof Date ? Utilities.formatDate(startDt,  tz, 'yyyy-MM-dd') : '';
    var handoverStr = handover instanceof Date ? Utilities.formatDate(handover, tz, 'yyyy-MM-dd') : '';

    // Notes = Location + Remarks (if any)
    var notes = [
      location && ('Location: ' + location),
      remarks
    ].filter(Boolean).join(' · ');

    var payload = {
      Project_ID:      projectId,
      Name:            name,
      Customer_ID:     customerId,
      Start_Date:      startStr,
      Target_Handover: handoverStr,
      Status:          mStatus,
      Notes:           notes,
      'Active?':       isActive ? 'Yes' : 'No'
    };

    try {
      Masters.createProject(payload);   // upsert — Masters uses Project_ID as key
      pushed++;
    } catch (e) {
      errors.push('Row ' + (i + 2) + ' (' + projectId + '): ' + e.message);
    }
  });

  // Bust the local cache so IM_getProjects() returns fresh data next call
  try { IM_clearCache(); } catch (e) {}

  var msg = 'Pushed to Masters:\n' +
    '  pushed/updated: ' + pushed + '\n' +
    '  blank rows (skipped): ' + skipped +
    (errors.length ? '\n\n  errors (' + errors.length + '):\n    ' + errors.slice(0, 10).join('\n    ') : '');
  SpreadsheetApp.getUi().alert(msg);
}
