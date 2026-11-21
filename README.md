# Project-Management (PM)

Google Apps Script bound to the "3SP Project Management" spreadsheet. Tracks every project from order to handover and connects to the other 3SP systems. **Ticketing is not part of PM any more** (it moved to the separate 3SP Ticketing system; the old worker link shows a "we moved" page).

> Part of the 3SP family: Masters-System, IMS, Finance-System, Ticketing, 3SP_DocGen. See "How the systems fit together" at the bottom.

## Sheets

| Sheet | What it is |
|---|---|
| **Projects** | One row per project: ID `PRJ-###`, name, client, location, dates, status; formulas for label, delay status, days left and stage percentages. A new project is sent to Masters automatically (on edit, on open, and with the Add Project menu). |
| **Milestones** | 25 milestones per project (Yes / No / Pending) feeding the completion percentages. |
| **Hardware List** | What each project needs: project, material (IMS item dropdown), photo, specification, required qty, live In Stock (IMS available), unit, **Reserve?** checkbox, Item Status, Stock Status, remarks, pushed-to-Purchases code, IMS reservation. |
| **Purchases** | The shopping list: imported from the Hardware List (shortfall only: project, material, specification, required qty, In Stock, unit). Ordered qty, vendor, PO issued, delivery dates, status and remarks are filled by hand; the vendor dropdown is the IMS suppliers. |
| **Production**, **Snags** | Production items and site snags. |
| **Dashboard**, **Project View** | KPIs and a per-project view (pick a project in B3). |
| **People** | Generated copy of Masters employees (kept for dropdowns). |

Hardware List, Purchases and Production have a **project filter** in rows 1-2 (merged dropdown, header on row 3, data from row 4): pick a project to see only its rows, "All projects" shows everything.

## The hardware flow

1. Pick a material (IMS item or free text). In Stock shows the IMS Available figure; Unit follows the material.
2. Tick **Reserve?**: PM asks IMS to reserve what exists (up to the required qty); the Reservation Register in IMS is the queue (a one-minute IMS trigger turns the request into a reservation). The row shows `RSV-0012 | reserved 5`, then follows IMS (partially consumed / consumed / released).
3. Anything short is pushed to **Purchases** as the shortfall (Required = shortfall, Ordered left empty). A row already reserved or pushed is never done twice.
4. A Material Requisition or a Delivery Challan in IMS for that project **consumes** the matching reservation; Item Status reads the Register live.

## Other automation

- Project delete everywhere: menu Delete Project (preview, type the PRJ id, Masters removes the project from PM, IMS and Finance).
- Add item / Add supplier / Add employee to IMS or Masters from PM (commands picked up by IMS).
- **Self-test** (Interlink menu): creates a throw-away project `ZZ-SELFTEST...`, walks the whole chain (Masters, reservations, requisition, delivery challan with a return, Finance import), writes a PASS / FAIL report sheet and deletes everything. About 15-20 minutes.
- Health check and *Needs attention* sheets show anything broken in plain words; many things heal themselves when the sheet opens.

## Files

`Config.gs` constants and column maps · `Util.gs` row helpers (project-filter layout: `topRows_`, `firstRow_`) · `Sheets.gs` sheet builders, validations, formatting, project filter, photo column · `Triggers.gs` edit trigger (auto IDs, Hardware to Purchases, project sync) · `Menu.gs` menu, Add / Delete project · `Setup.gs` bootstrap · `Api.gs` + `Moved.html` the "moved" page · `_Interlink_*.gs` the connections: `Masters` (library, dropdowns, health, self-test hooks), `PushProjects`, `Items` (IMS items, photos, units, self-test), `HardwareFlow`, `Reserve` · `_Retire_Tickets.gs` one-time removal of the old Tickets sheet.

## Deploy / development

- Work from this folder: `clasp push --force`, then commit and push to `main`. The GitHub Action (`.github/workflows/deploy.yml`) repeats the clasp push as a safety net and needs the repo secret `CLASPRC_JSON` (contents of `~/.clasprc.json`; if it fails with `invalid_rapt`, run `clasp login` and refresh the secret).
- `appsscript.json` pins the **Masters library** version; bump it deliberately after a Masters release.
- Foot-guns: top-level object literals must not use computed keys; Date objects must be turned into text before `google.script.run` returns them; ARRAYFORMULA seeds live on the first data row (row 4 on the filtered sheets).
- Offline tests (Node `vm`) are kept in `_tests` next to the repos, not in git.

## How the systems fit together

```
              Masters-System  (library: projects, customers, parties, employees, sites)
             /     |      |      \
     Finance   IMS   Project-Mgmt  Ticketing         3SP_DocGen (Next.js) --web app--> IMS
        ^        ^        |                              (documents, numbers, PDFs)
        |        |--reservations / requests-----|
        |--- 10-minute import of project material costs from IMS
```
