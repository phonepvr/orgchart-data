// --- Template Schema ---
export const REQUIRED_COLUMNS = [
    "Employee's Position Code", 'Employee name', "Line Manager's Position Code",
    'Position Text', 'Level', 'Function 1', 'Function/Plant', 'Location Name',
    'Asset',
];
export const RECOMMENDED_COLUMNS = [
    'Line Manager Name', 'Employee Class',
    'Cluster', 'Gender', 'Date of Birth', 'HR Manager EID', 'HR Manager Name', 'Management Board EID'
];
export const OPTIONAL_COLUMNS = [
    'Date of Joining', 'Date in Role', 'Date Promoted', 'Manager Since',
    'Email', 'Photo URL', 'Matrix Manager EID(s)', 'Cohort Tags', 'Current Status/Tag'
];

// Status color tokens shared by card, spotlight, table, print
export const STATUS_STYLES = {
    Active:   { chip: 'bg-leaf/20 text-leaf border-leaf/60',         rule: '#3F9460', label: 'Active' },
    WIP:      { chip: 'bg-ember/20 text-ember border-ember/60',      rule: '#D9761E', label: 'WIP' },
    Offered:  { chip: 'bg-signal/20 text-signal border-signal/60',   rule: '#1B5EA6', label: 'Offered' },
    Vacant:   { chip: 'bg-red-100 text-red-800 border-red-300',      rule: '#B81F1F', label: 'Vacant' },
};
export const NAME_STATUS_TINT = {
    approved:   { card: 'bg-graphite-200 border-graphite-400',                printTile: 'bg-[#D6DAE2] border-[#8892A3]',   label: 'Approved' },
    unapproved: { card: 'bg-signal/20 border-signal/60 text-graphite-900',    printTile: 'bg-[#D2E3F2] border-[#1B5EA6]',   label: 'Unapproved' },
};

// --- Access Gate ---
// SHA-256 of the access password. Default: "amns2026".
// To change: run `node -e "console.log(require('crypto').createHash('sha256').update('NEWPASS').digest('hex'))"`
// and replace this constant. NOTE: this is client-side obfuscation, not real
// authentication — anyone with devtools can read the source. Use it to gate the
// public Pages URL from casual visitors, not to protect sensitive data.
export const ACCESS_HASH = 'bc321ef4abcfa473576619373b465dd491e5bced3b34b0b22dd1d54786b46f58';

// --- Filter Field Definitions ---
export const FILTER_FIELD_MAP = {
    'Current Status/Tag': 'currentStatus',
    'Level': 'level',
    'Function 1': 'function1',
    'Function/Plant': 'functionPlant',
    'Location': 'location',
    'Asset': 'asset',
    'Cluster': 'cluster',
    'Employee Class': 'employeeClass',
    'Gender': 'gender',
};
export const MULTI_SELECT_FIELDS = [...Object.keys(FILTER_FIELD_MAP), 'Cohort Tag', 'Mgmt Committee'];
export const NUMERIC_FIELDS = ['DR Size', 'Total Reportees', 'Team Size'];

// Compare color slots (also safelisted as hover:ring-{id}-400 in the frozen CSS)
export const COMPARE_COLORS = [
    { id: 'blue', bg: 'bg-blue-500', text: 'text-blue-700', border: 'border-blue-500', light: 'bg-blue-50' },
    { id: 'green', bg: 'bg-green-500', text: 'text-green-700', border: 'border-green-500', light: 'bg-green-50' },
    { id: 'red', bg: 'bg-red-500', text: 'text-red-700', border: 'border-red-500', light: 'bg-red-50' },
    { id: 'orange', bg: 'bg-orange-500', text: 'text-orange-700', border: 'border-orange-500', light: 'bg-orange-50' },
    { id: 'purple', bg: 'bg-purple-500', text: 'text-purple-700', border: 'border-purple-500', light: 'bg-purple-50' },
];

// --- Print layout constants ---
// Deterministic per-page capacity budget (no DOM measurement).
// A4 landscape inner area ~281x194mm; tile ~30mm tall + gap => ~4 rows/col.
export const NORMAL_PAGE_CAP = 8;        // subject page: 2 cols x 4 rows in the side pane
export const CONTINUATION_PAGE_CAP = 16; // continuation page: 4 cols x 4 rows full-width

// How many subject levels below the print root get their own pages. Each
// page shows its subject plus one level below, so the printed tree reaches
// root (n) → n-1 → n-2 → n-3 with a depth of 2.
export const PRINT_SUBJECT_DEPTH = 2;

// Static class strings for print grid columns.
export const GRID_COLS = { 1: 'grid-cols-1', 2: 'grid-cols-2', 3: 'grid-cols-3' };
