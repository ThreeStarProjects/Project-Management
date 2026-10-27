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
  TICKETS:    'Tickets',
  // Reference
  PEOPLE:     'People',
  VENDORS:    'Vendors',
  // Read-only
  VIEW:       'Project View',
  DASHBOARD:  'Dashboard'
};

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
  [SHEETS.SNAGS]:      'SNAG',
  [SHEETS.TICKETS]:    'TKT'
};

const MAX_ROWS = 1000;

// ── TICKET CONFIG ──────────────────────────────────────────────────────────
const TICKET_CATEGORIES = {
  'Machine Breakdown':          { priority: 'High',     dueDays: 1, fields: ['machineId'] },
  'Material Shortage':          { priority: 'High',     dueDays: 2, fields: ['batchCode','relatedProject'] },
  'Quality Defect (In-Process)':{ priority: 'High',     dueDays: 1, fields: ['batchCode','relatedProject'] },
  'Design Change Request':      { priority: 'Medium',   dueDays: 3, fields: ['drawingRev','relatedProject'] },
  'Sample/Mockup Approval':     { priority: 'Medium',   dueDays: 5, fields: ['drawingRev','relatedProject'] },
  'Technical Query':            { priority: 'Medium',   dueDays: 2, fields: ['drawingRev','relatedProject'] },
  'Vendor Issue':               { priority: 'Medium',   dueDays: 3, fields: ['vendorName','amount'] },
  'Purchase Request':           { priority: 'Medium',   dueDays: 3, fields: ['vendorName','amount','relatedProject'] },
  'Site Issue':                 { priority: 'High',     dueDays: 1, fields: ['relatedProject'] },
  'Installation Snag':          { priority: 'Medium',   dueDays: 3, fields: ['relatedProject'] },
  'Delivery Coordination':      { priority: 'Medium',   dueDays: 2, fields: ['relatedProject'] },
  'Client Complaint':           { priority: 'High',     dueDays: 1, fields: ['relatedProject'] },
  'Client Change Request':      { priority: 'Medium',   dueDays: 5, fields: ['amount','relatedProject'] },
  'Warranty Claim':             { priority: 'Medium',   dueDays: 5, fields: ['relatedProject'] },
  'HR / Leave':                 { priority: 'Low',      dueDays: 3, fields: [] },
  'Safety Incident':            { priority: 'Critical', dueDays: 0, fields: ['machineId'] },
  'General Admin':              { priority: 'Low',      dueDays: 5, fields: [] },
  'Payment Issue':              { priority: 'Medium',   dueDays: 3, fields: ['amount','relatedProject'] },
  'Expense Reimbursement':      { priority: 'Low',      dueDays: 5, fields: ['amount'] }
};
const TICKET_STATUSES   = ['Open','In Progress','On Hold','Resolved','Closed'];
const TICKET_PRIORITIES = ['Low','Medium','High','Critical'];
const TICKET_IMPACTS    = ['Low','Medium','High','Critical'];

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
