/**
 * ============================================================================
 * 3SP PROJECT MANAGEMENT — Config
 * All constants: sheet names, milestones, stages, ticket config, colors.
 * Edit here to tune anything that other files reference.
 * ============================================================================
 */

// ── SHEET NAMES ────────────────────────────────────────────────────────────
const SHEETS = {
  // Entry
  PROJECTS:   'Projects',
  MILESTONES: 'Milestones',
  HARDWARE:   'Hardware List',
  PURCHASES:  'Purchases',
  PRODUCTION: 'Production',
  SNAGS:      'Snags',
  // Reference
  PEOPLE:     'People',
  VENDORS:    'Vendors',
  // Read-only
  VIEW:       'Project View',
  DASHBOARD:  'Dashboard'
};

// Hardware List / Purchases / Production start with TOP_ROWS control rows (a merged "Project filter" dropdown), then the header row, then the
// data. So their header is row 1 + TOP_ROWS and their first data row is 2 + TOP_ROWS; every other sheet keeps header row 1 / data from row 2.
// ensureTopRows_() (Sheets.gs) moves a sheet that is still in the old layout, by itself, the first time anything runs.
const TOP_ROWS = 2;
const FILTER_SHEETS = ['Hardware List', 'Purchases', 'Production'];
const FILTER_HEADERS = { 'Hardware List': 'Hardware ID', 'Purchases': 'Purchase ID', 'Production': 'Production ID' };  // what A holds on the header row
const FILTER_ALL = 'All projects';
const LAYOUT_CACHE_KEY = 'layout_ok_v5';   // forgotten whenever a Hardware List / Purchases / Production sheet is (re)created, so the layout is checked again
const FILTER_LIST_SHEET = '_Filter';   // hidden: "All projects" + the project labels, the dropdown's list
const FILTER_LABEL_COL = 2;            // B1:B2 merged: the words "Project filter"
const FILTER_CELL_COL  = 3;            // C1:C2 merged: the dropdown (A is frozen on these sheets and a merge cannot cross that line)
const FILTER_PROJECT_COL = 2;          // B: the Project column the filter works on

/** Sheets from earlier builds (webapp era + v2 pre-slim). Deleted on setupAll if present. */
const LEGACY_SHEETS = [
  'Invoices',       // dropped — no financial tracking
  'Log',            // dropped — activity logging removed
  'ActivityLog',    // webapp era
  'Employees',      // webapp era (v2 uses People)
  'VendorPayments', // webapp era
  'Procurement',    // webapp era (v2 uses Purchases)
  'Snag',           // webapp era singular (v2 uses Snags)
  'Project Card'    // webapp era (v2 uses Project View)
];

// 25 milestones. Order matches Milestones sheet cols C-AA.
const MILESTONES = [
  'Requirements Received', 'Site Visit Done', 'Floor Plan Approved',
  '3D Design Approved', 'BOQ Approved', 'Working Drawings Ready',
  'Cutting Drawings Ready', 'Cut List Ready', 'Advance Received',
  'PO Raised', 'Material Received', 'Incoming QC Done',
  'Study Drawings & Tally Done', 'Production Planning Done',
  '1st Piece Cutting Done', 'Mockup Made', 'Mockup QC Done',
  'Bulk Production Done', 'Final QC Done', 'Store & Packed',
  'Dispatched', 'Site Delivery Done', 'Installation Done',
  'Snag List Made', 'Handover Done'
];

// Stage → milestone index range (0-indexed, inclusive) + target column on Projects.
// Cols shifted after Owner (E) and Contract (H) were removed.
const STAGES = [
  { name: 'Pre-Execution', start: 0,  end: 8,  col: 12 }, // Projects col L
  { name: 'Procurement',   start: 9,  end: 11, col: 13 }, // M
  { name: 'Production',    start: 12, end: 18, col: 14 }, // N
  { name: 'Delivery',      start: 19, end: 21, col: 15 }, // O
  { name: 'Installation',  start: 22, end: 24, col: 16 }  // P
];

// Sheet → ID prefix. Handled by onEdit auto-ID.
const ID_PREFIX_MAP = {
  [SHEETS.PROJECTS]:   'PRJ',
  [SHEETS.HARDWARE]:   'HW',
  [SHEETS.PURCHASES]:  'PUR',
  [SHEETS.PRODUCTION]: 'PROD',
  [SHEETS.SNAGS]:      'SNAG'
};

const MAX_ROWS = 1000;

// (Ticketing moved to its own system: 3SP Ticketing.)
const TICKETING_URL = 'https://script.google.com/macros/s/AKfycbwf8I9DTWuc9pvTzJJXiME1g3J8lIJ6L3UIkyxWsl6h8RA93rhc4PIIhD5-G9f1GA-02A/exec';

// ── COLORS (gamified / bold) ───────────────────────────────────────────────
// Headers deep indigo, section bands electric purple. Status cells are
// saturated-fill + white bold text so they read like game HUD badges.
const COLORS = {
  HEADER_BG:  '#1E3A8A', HEADER_FG:  '#FFFFFF',  // deep indigo
  SECTION_BG: '#7C3AED',                         // electric purple
  ACCENT:     '#F59E0B',                         // vivid amber
  DANGER:     '#DC2626',                         // true red
  DONE_BG:    '#10B981', DONE_FG:    '#FFFFFF',  // emerald
  PROGRESS_BG:'#3B82F6', PROGRESS_FG:'#FFFFFF',  // sky
  PENDING_BG: '#64748B', PENDING_FG: '#FFFFFF',  // slate
  WARN_BG:    '#F59E0B', WARN_FG:    '#FFFFFF',  // amber
  BAD_BG:     '#EF4444', BAD_FG:     '#FFFFFF'   // red
};

// Punchy row tints for per-project banding (readable with dark text on top).
const PROJECT_PALETTE = [
  '#FEE2E2', // rose
  '#DBEAFE', // sky
  '#D1FAE5', // mint
  '#FED7AA', // peach
  '#E9D5FF', // lavender
  '#FEF3C7', // butter
  '#CFFAFE', // cyan
  '#E5E7EB'  // silver
];
