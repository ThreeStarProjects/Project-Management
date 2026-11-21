/**
 * ============================================================================
 * 3SP PROJECT MANAGEMENT - Api
 * Ticketing is no longer part of PM: it is its own system (3SP Ticketing). The old worker link (this web app) now only shows
 * a "we moved" page that sends people to the new form.
 * ============================================================================
 */

function doGet() {
  const t = HtmlService.createTemplateFromFile('Moved');
  t.url = TICKETING_URL;
  return t.evaluate()
    .setTitle('3SP Tickets have moved')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
