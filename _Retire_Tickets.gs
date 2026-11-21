/**
 * _Retire_Tickets.gs - one time. Ticketing moved to its own system, so the empty 'Tickets' sheet and the ticket numbers on the
 * Dashboard / Project View go. Runs when PM opens; does nothing again afterwards (document property TICKETING_RETIRED).
 * Safety: the Tickets sheet is only deleted if it holds no ticket rows. If it does, nothing is deleted and the user is told.
 */
function retireTicketing_() {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('TICKETING_RETIRED') === 'v1') return 'already';
  var ss = SpreadsheetApp.getActiveSpreadsheet(), tk = ss.getSheetByName('Tickets');
  var rows = tk && tk.getLastRow() >= 2 ? tk.getRange(2, 1, tk.getLastRow() - 1, 1).getValues().filter(function (r) { return String(r[0]).trim() !== ''; }).length : 0;
  if (tk && rows > 0) {
    ss.toast('The Tickets sheet still has ' + rows + ' ticket(s), so it was NOT deleted. Move them to the new ticketing system, then delete the sheet.', 'Ticketing', 15);
    return 'has tickets';
  }
  // 1. the ticket numbers that read the sheet: clear them first so nothing turns into #REF!
  var dash = ss.getSheetByName(SHEETS.DASHBOARD);
  if (dash && dash.getLastRow() >= 1) {
    var a = dash.getRange(1, 1, Math.min(dash.getLastRow(), 30), 2).getValues();
    a.forEach(function (r, i) {
      var t = String(r[0]).trim();
      if (t === 'TICKETS & SNAGS') dash.getRange(i + 1, 1).setValue('SNAGS');
      if (t === 'Open Tickets' || t === 'Overdue Tickets') dash.getRange(i + 1, 1, 1, 2).clearContent();
    });
  }
  var view = ss.getSheetByName(SHEETS.VIEW);
  if (view && view.getLastRow() >= 1) {
    view.getRange(1, 1, Math.min(view.getLastRow(), 40), 5).getValues().forEach(function (r, i) {
      if (String(r[0]).trim() === 'Tickets Total:') view.getRange(i + 1, 1, 1, 5).clearContent();
    });
  }
  // 2. the empty sheet
  if (tk) { try { ss.deleteSheet(tk); } catch (e) { Logger.log('Tickets sheet not deleted: ' + e.message); return 'delete failed'; } }
  props.setProperty('TICKETING_RETIRED', 'v1');
  ss.toast('Tickets now live in the new 3SP Ticketing system. The old Tickets sheet was removed.', 'Ticketing', 10);
  return 'retired';
}
