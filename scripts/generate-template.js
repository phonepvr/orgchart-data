// Regenerates orglens_sample_template.xlsx at the repo root.
// Single source of truth for the spreadsheet schema the app expects.
// Run with: npm run build:template
//
// Row 1 = column headers
// Row 2 = tier labels (Required / Recommended / Optional) — handleFileUpload
//         filters this row out at upload time
// Rows 3+ = ~500 deterministically-seeded sample rows exercising every filter
//          dimension and visual state (Active / WIP / Offered / Vacant,
//          Approved / Unapproved placeholders, matrix relationships, cohorts).

import * as XLSX from 'xlsx';
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const outPath = resolve(here, '..', 'orglens_sample_template.xlsx');

const REQUIRED = ["Employee's Position Code", "Employee name", "Line Manager's Position Code"];
const RECOMMENDED = [
    'Line Manager Name', 'Position Text', 'Level', 'Employee Class',
    'Function 1', 'Function/Plant', 'Location Name', 'Asset', 'Cluster',
    'Gender', 'Date of Birth', 'HR Manager Name', 'HR Manager EID', 'Management Board EID',
];
const OPTIONAL = [
    'Date of Joining', 'Date in Role', 'Date Promoted', 'Manager Since',
    'Email', 'Photo URL', 'Matrix Manager EID(s)', 'Cohort Tags', 'Current Status/Tag',
];
const HEADERS = [...REQUIRED, ...RECOMMENDED, ...OPTIONAL];

const tierRow = HEADERS.map(h =>
    REQUIRED.includes(h) ? 'Required' : RECOMMENDED.includes(h) ? 'Recommended' : 'Optional'
);

// --- Deterministic PRNG (mulberry32) ---
function mulberry32(seed) {
    let a = seed >>> 0;
    return function () {
        a = (a + 0x6D2B79F5) >>> 0;
        let t = a;
        t = Math.imul(t ^ (t >>> 15), t | 1);
        t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}
const rand = mulberry32(0xA17C0DE);
const pick = (arr) => arr[Math.floor(rand() * arr.length)];
const weighted = (entries) => {
    const total = entries.reduce((s, [, w]) => s + w, 0);
    let r = rand() * total;
    for (const [val, w] of entries) {
        if ((r -= w) <= 0) return val;
    }
    return entries[entries.length - 1][0];
};

// --- Pools ---
// First names tagged with gender so the gender field is consistent
const FIRST_NAMES = [
    ['Asha', 'F'], ['Rohit', 'M'], ['Priya', 'F'], ['Karan', 'M'], ['Sneha', 'F'],
    ['Vikram', 'M'], ['Maya', 'F'], ['Ananya', 'F'], ['Aditya', 'M'], ['Pooja', 'F'],
    ['Arjun', 'M'], ['Neha', 'F'], ['Rahul', 'M'], ['Kavya', 'F'], ['Deepak', 'M'],
    ['Meena', 'F'], ['Suresh', 'M'], ['Lakshmi', 'F'], ['Anil', 'M'], ['Geeta', 'F'],
    ['Sanjay', 'M'], ['Divya', 'F'], ['Ramesh', 'M'], ['Sunita', 'F'], ['Manoj', 'M'],
    ['Riya', 'F'], ['Nikhil', 'M'], ['Aarti', 'F'], ['Vivek', 'M'], ['Anjali', 'F'],
    ['Sandeep', 'M'], ['Ritu', 'F'], ['Harsh', 'M'], ['Tara', 'F'], ['Yash', 'M'],
    ['Pallavi', 'F'], ['Akash', 'M'], ['Shweta', 'F'], ['Tanmay', 'M'], ['Isha', 'F'],
    ['Pranav', 'M'], ['Swati', 'F'], ['Varun', 'M'], ['Madhuri', 'F'], ['Kunal', 'M'],
    ['Megha', 'F'], ['Aman', 'M'], ['Roshni', 'F'], ['Siddharth', 'M'], ['Kirti', 'F'],
];
const LAST_NAMES = [
    'Sharma', 'Iyer', 'Menon', 'Mehta', 'Pillai', 'Singh', 'Krishnan', 'Reddy',
    'Patel', 'Gupta', 'Nair', 'Banerjee', 'Chatterjee', 'Verma', 'Joshi', 'Bhatt',
    'Desai', 'Kapoor', 'Malhotra', 'Agarwal', 'Kulkarni', 'Mukherjee', 'Roy', 'Das',
    'Bhattacharya', 'Trivedi', 'Choudhury', 'Saxena', 'Tiwari', 'Bhandari',
];

const FUNCTIONS = [
    'Iron Making', 'Steel Making', 'Rolling Mills', 'Pellet Plant', 'Coke Plant',
    'Mines', 'Power Plant', 'Finance', 'HR', 'IT', 'Legal', 'Procurement',
    'Supply Chain', 'Sales & Marketing', 'Strategy', 'Operations Excellence',
    'Safety & Sustainability', 'R&D',
];

// Function/Plant sub-bins by Function 1 (keeps the data internally coherent)
const PLANTS_BY_FUNCTION = {
    'Iron Making': ['Blast Furnace', 'Sinter Plant', 'Sponge Iron'],
    'Steel Making': ['BOF', 'EAF', 'Continuous Casting'],
    'Rolling Mills': ['HSM', 'CRM', 'Pickling Line', 'Galvanizing', 'Color Coating'],
    'Pellet Plant': ['Pellet Plant 1', 'Pellet Plant 2'],
    'Coke Plant': ['Coke Oven 1', 'Coke Oven 2', 'By-Products'],
    'Mines': ['Iron Ore Mining', 'Limestone Mining'],
    'Power Plant': ['Captive Power', 'Renewable'],
    'Finance': ['Treasury', 'Controlling', 'Tax', 'Internal Audit', 'Investor Relations'],
    'HR': ['Talent Acquisition', 'L&D', 'C&B', 'HR Operations'],
    'IT': ['IT Infra', 'IT Apps', 'SAP', 'Cybersecurity'],
    'Legal': ['Corporate Legal', 'Compliance'],
    'Procurement': ['Direct Procurement', 'Indirect Procurement'],
    'Supply Chain': ['Logistics', 'Imports', 'Warehousing'],
    'Sales & Marketing': ['Domestic Sales', 'Exports', 'OEM Sales', 'Brand Marketing'],
    'Strategy': ['Corporate Strategy', 'M&A'],
    'Operations Excellence': ['Lean Six Sigma', 'Quality'],
    'Safety & Sustainability': ['EHS', 'Sustainability'],
    'R&D': ['Product R&D', 'Process R&D'],
};

// Each location pins to a cluster + reasonable asset set
const LOCATIONS = [
    { name: 'Hazira', cluster: 'West',      assets: ['Hazira Plant', 'Hazira Port'] },
    { name: 'Paradip', cluster: 'East',      assets: ['Paradip Plant', 'Paradip Port'] },
    { name: 'Kolkata HO', cluster: 'East',   assets: ['Corporate'] },
    { name: 'Mumbai HO', cluster: 'West',    assets: ['Corporate'] },
    { name: 'New Delhi', cluster: 'North',   assets: ['Corporate'] },
    { name: 'Pune', cluster: 'West',         assets: ['Corporate'] },
    { name: 'Bengaluru', cluster: 'South',   assets: ['Corporate'] },
    { name: 'Indore', cluster: 'West',       assets: ['Corporate'] },
    { name: 'Jamshedpur', cluster: 'East',   assets: ['Corporate'] },
    { name: 'Vizag', cluster: 'South',       assets: ['Corporate'] },
    { name: 'Bhubaneswar', cluster: 'East',  assets: ['Odisha Mines'] },
    { name: 'Raipur', cluster: 'East',       assets: ['Karnataka Mines'] },
    { name: 'Anjar', cluster: 'West',        assets: ['Anjar Plant'] },
];

const COHORTS = [
    'HiPo', 'Emerging Leader', 'Women in Leadership', 'FastTrack',
    'Digital Champion', 'Safety Steward', 'Sustainability Council',
];

const LEVELS = ['CEO', 'CXO', 'VP', 'GM', 'DGM', 'Manager', 'Sr Exec', 'Exec'];

// Each CXO can oversee multiple Function 1 values; subordinates pick from this
// pool, which is how all 18 functions get represented in the tree.
const CXO_TITLES = [
    { title: 'Chief Financial Officer',          fn: 'Finance',
      fnPool: ['Finance', 'Procurement', 'Supply Chain'],
      loc: 'Mumbai HO' },
    { title: 'Chief Human Resources Officer',    fn: 'HR',
      fnPool: ['HR'],
      loc: 'Mumbai HO' },
    { title: 'Chief Technology Officer',         fn: 'IT',
      fnPool: ['IT', 'R&D'],
      loc: 'Bengaluru' },
    { title: 'Chief Operating Officer',          fn: 'Steel Making',
      fnPool: ['Iron Making', 'Steel Making', 'Rolling Mills', 'Pellet Plant', 'Coke Plant', 'Mines', 'Power Plant', 'Operations Excellence'],
      loc: 'Hazira' },
    { title: 'Chief Marketing Officer',          fn: 'Sales & Marketing',
      fnPool: ['Sales & Marketing', 'Strategy'],
      loc: 'Mumbai HO' },
    { title: 'Chief Sustainability Officer',     fn: 'Safety & Sustainability',
      fnPool: ['Safety & Sustainability', 'Legal'],
      loc: 'New Delhi' },
];

// Function type drives which locations are realistic for a row.
const PLANT_FUNCTIONS = new Set([
    'Iron Making', 'Steel Making', 'Rolling Mills', 'Pellet Plant', 'Coke Plant',
    'Mines', 'Power Plant',
]);
const PLANT_LOCATIONS = ['Hazira', 'Paradip', 'Anjar', 'Jamshedpur', 'Vizag', 'Indore', 'Bhubaneswar', 'Raipur'];
const CORPORATE_LOCATIONS = ['Mumbai HO', 'Kolkata HO', 'New Delhi', 'Pune', 'Bengaluru'];
const locationsFor = (fn) => (PLANT_FUNCTIONS.has(fn) ? PLANT_LOCATIONS : CORPORATE_LOCATIONS);

// --- Date helpers ---
const TODAY = new Date('2026-05-11');
const DAY_MS = 24 * 60 * 60 * 1000;
const fmtDate = (d) => d.toISOString().slice(0, 10);
const dateBetween = (startYearsAgo, endYearsAgo) => {
    const start = new Date(TODAY.getTime() - startYearsAgo * 365.25 * DAY_MS);
    const end = new Date(TODAY.getTime() - endYearsAgo * 365.25 * DAY_MS);
    const t = start.getTime() + rand() * (end.getTime() - start.getTime());
    return new Date(t);
};

// --- Title generator ---
const titleFor = (level, fn, plant) => {
    switch (level) {
        case 'CEO': return 'Managing Director & CEO';
        case 'CXO': return fn;
        case 'VP': return `VP - ${fn}`;
        case 'GM': return `GM - ${plant || fn}`;
        case 'DGM': return `DGM - ${plant || fn}`;
        case 'Manager': return `Manager - ${plant || fn}`;
        case 'Sr Exec': return `Sr Executive - ${plant || fn}`;
        case 'Exec': return `Executive - ${plant || fn}`;
        default: return fn;
    }
};

const empClassFor = (level) => {
    switch (level) {
        case 'CEO':
        case 'CXO':
        case 'VP':
        case 'GM':
        case 'DGM': return 'Management';
        case 'Manager': return weighted([['Management', 7], ['Senior Officer', 3]]);
        case 'Sr Exec': return weighted([['Senior Officer', 5], ['Officer', 4], ['Management', 1]]);
        case 'Exec':    return weighted([['Officer', 4], ['Workman', 4], ['Trainee', 2]]);
        default: return 'Officer';
    }
};

const levelLabelFor = (level) => {
    switch (level) {
        case 'CEO': return 'MD - Managing Director';
        case 'CXO': return 'SVP - Sr. Vice President';
        case 'VP':  return 'VP - Vice President';
        case 'GM':  return 'GM - General Manager';
        case 'DGM': return 'DGM - Deputy General Manager';
        case 'Manager': return 'M3';
        case 'Sr Exec': return 'M2';
        case 'Exec':    return 'M1';
        default: return level;
    }
};

const statusFor = () => weighted([
    ['Active', 72], ['WIP', 12], ['Offered', 8], ['Vacant', 8],
]);

const cohortsFor = (level) => {
    if (rand() > 0.15) return '';
    const n = rand() < 0.3 ? 2 : 1;
    const seen = new Set();
    const picks = [];
    while (picks.length < n) {
        const c = pick(COHORTS);
        if (!seen.has(c)) { seen.add(c); picks.push(c); }
    }
    // Mgmt Committee is reserved for the top 7; do not surface here.
    if (level === 'CXO' && rand() < 0.5) picks.push('HiPo');
    return picks.join('; ');
};

// --- Code generation ---
let counters = { C: 0, X: 0, V: 0, G: 0, M: 0, E: 0 };
const codeFor = (level) => {
    const prefixMap = { CEO: 'C', CXO: 'X', VP: 'V', GM: 'G', DGM: 'G', Manager: 'M', 'Sr Exec': 'E', Exec: 'E' };
    const p = prefixMap[level];
    counters[p] += 1;
    return `AMNS-${p}${String(counters[p]).padStart(6, '0')}`;
};

// --- Row builder ---
// Photo URL is intentionally left blank for every row so the sample template
// makes zero external requests at render time. Customers who include their
// own Photo URLs at upload still get them loaded by the browser as before.

const realName = () => {
    const [fn, gender] = pick(FIRST_NAMES);
    const ln = pick(LAST_NAMES);
    return { name: `${fn} ${ln}`, gender };
};

const ROWS = [];
const byLevel = { CEO: [], CXO: [], VP: [], GM: [], DGM: [], Manager: [], 'Sr Exec': [], Exec: [] };

const buildRow = ({ level, manager, fnOverride, locOverride, statusOverride }) => {
    const status = statusOverride || statusFor();
    const isVacant = status === 'Vacant';
    let name, gender;
    if (isVacant) {
        // Placeholder seat — alternate Approved/Unapproved deterministically.
        name = rand() < 0.5 ? 'Approved' : 'Unapproved';
        gender = '';
    } else {
        ({ name, gender } = realName());
    }
    // Function: VPs pick from their CXO's pool; below inherits the VP's fn.
    let fn = fnOverride;
    if (!fn) {
        if (level === 'VP' && manager && manager._fnPool) fn = pick(manager._fnPool);
        else if (manager) fn = manager._fn;
        else fn = pick(FUNCTIONS);
    }
    const plant = pick(PLANTS_BY_FUNCTION[fn] || [fn]);
    // Location: VPs pick from realistic candidates for their function; lower
    // levels prefer the same location as their manager, but ~25% diverge.
    let locName;
    if (locOverride) locName = locOverride;
    else if (level === 'VP') locName = pick(locationsFor(fn));
    else if (manager && rand() < 0.75) locName = manager.location;
    else locName = pick(locationsFor(fn));
    const locDef = LOCATIONS.find((l) => l.name === locName) || LOCATIONS[0];
    const location = locDef.name;
    const cluster = locDef.cluster;
    const asset = pick(locDef.assets);

    const eid = codeFor(level);
    const dob = ['CEO', 'CXO'].includes(level)
        ? dateBetween(60, 50)
        : level === 'VP' ? dateBetween(55, 42)
        : level === 'GM' || level === 'DGM' ? dateBetween(50, 36)
        : level === 'Manager' ? dateBetween(45, 30)
        : dateBetween(38, 24);
    const doj = isVacant ? '' : fmtDate(dateBetween(25, 0.5));
    const dir = isVacant ? '' : fmtDate(dateBetween(8, 0.1));
    const prom = dir;

    const row = {
        eid,
        name,
        mgrEid: manager ? manager.eid : '',
        mgrName: manager ? manager.name : '',
        title: titleFor(level, fn, plant),
        level: levelLabelFor(level),
        empClass: empClassFor(level),
        function1: fn,
        plant,
        location,
        asset,
        cluster,
        gender,
        dob: isVacant ? '' : fmtDate(dob),
        hrName: '',
        hrEid: '',
        boardEid: level === 'CEO' ? '' : (byLevel.CEO[0] ? byLevel.CEO[0].eid : ''),
        doj,
        dir,
        prom,
        mgrSince: '',
        email: isVacant ? '' : `${name.toLowerCase().replace(/\s+/g, '.')}.${eid.slice(-4)}@example.com`,
        photo: '',
        matrix: '',
        cohorts: isVacant ? '' : cohortsFor(level),
        status,
        // internal
        _level: level,
        _fn: fn,
        _fnPool: undefined,
    };
    ROWS.push(row);
    byLevel[level].push(row);
    return row;
};

// --- Build the tree ---
// CEO
const ceo = buildRow({ level: 'CEO', manager: null, fnOverride: 'Strategy', locOverride: 'Mumbai HO', statusOverride: 'Active' });

// 6 CXOs
const cxos = CXO_TITLES.map((spec) => {
    const row = buildRow({ level: 'CXO', manager: ceo, fnOverride: spec.fn, locOverride: spec.loc, statusOverride: 'Active' });
    row._fnPool = spec.fnPool;
    return row;
});

// VPs — at least one per function in the CXO's pool, then top up to 5 per CXO.
for (const cxo of cxos) {
    const pool = cxo._fnPool;
    const vpCount = Math.max(5, pool.length);
    for (let i = 0; i < vpCount; i++) {
        const fn = pool[i % pool.length];
        buildRow({ level: 'VP', manager: cxo, fnOverride: fn });
    }
}

// GMs and DGMs under VPs — 3 per VP, randomly GM/DGM
for (const vp of byLevel.VP.slice()) {
    for (let i = 0; i < 3; i++) {
        const lv = rand() < 0.5 ? 'GM' : 'DGM';
        buildRow({ level: lv, manager: vp });
    }
}

// Managers under GMs/DGMs
const gmPool = [...byLevel.GM, ...byLevel.DGM];
for (const gm of gmPool) {
    const n = 2 + Math.floor(rand() * 2); // 2-3 managers per GM
    for (let i = 0; i < n; i++) {
        buildRow({ level: 'Manager', manager: gm });
    }
}

// Sr Execs / Execs under Managers
for (const mgr of byLevel.Manager.slice()) {
    if (rand() < 0.15) continue; // some managers stay leaf
    const n = 1 + Math.floor(rand() * 2); // 1-2 reports
    for (let i = 0; i < n; i++) {
        const lv = rand() < 0.55 ? 'Sr Exec' : 'Exec';
        buildRow({ level: lv, manager: mgr });
    }
}

// Pad / trim to exactly 500
while (ROWS.length < 500) {
    const mgr = pick(byLevel.Manager);
    const lv = rand() < 0.55 ? 'Sr Exec' : 'Exec';
    buildRow({ level: lv, manager: mgr });
}
if (ROWS.length > 500) ROWS.length = 500;

// HR partner assignment: pick a GM/DGM in the HR function as HR Manager for
// everyone in the same cluster; fallback to first HR GM otherwise.
const hrGms = [...byLevel.GM, ...byLevel.DGM].filter((r) => r._fn === 'HR');
for (const r of ROWS) {
    if (r._level === 'CEO') continue;
    const partner = hrGms.find((g) => g.cluster === r.cluster) || hrGms[0];
    if (partner && partner.eid !== r.eid) {
        r.hrName = partner.name;
        r.hrEid = partner.eid;
    }
}

// MC membership marker: CEO + CXOs already covered by tree position; nothing to
// add. The app derives MC from Mgmt Board EID + level — Mgmt Board EID is
// already pointing at the CEO for everyone except the CEO.

// Matrix relationships — ~6% of Manager+ levels get one matrix manager from a
// different function's GM/DGM/VP pool. Deterministic via rand().
const matrixCandidates = [...byLevel.VP, ...byLevel.GM, ...byLevel.DGM];
for (const r of ROWS) {
    if (!['Manager', 'Sr Exec', 'Exec'].includes(r._level)) continue;
    if (rand() > 0.06) continue;
    const candidate = matrixCandidates.find((c) => c._fn !== r._fn && c.eid !== r.mgrEid);
    if (candidate) r.matrix = candidate.eid;
}

// Manager Since for anyone with direct reports
const reportsByMgr = new Map();
for (const r of ROWS) {
    if (!r.mgrEid) continue;
    reportsByMgr.set(r.mgrEid, (reportsByMgr.get(r.mgrEid) || 0) + 1);
}
for (const r of ROWS) {
    if (reportsByMgr.get(r.eid)) {
        r.mgrSince = r.dir || fmtDate(dateBetween(10, 0.5));
    }
}

const sample = ROWS.map((r) => ({
    "Employee's Position Code": r.eid,
    "Employee name": r.name,
    "Line Manager's Position Code": r.mgrEid,
    'Line Manager Name': r.mgrName,
    'Position Text': r.title,
    'Level': r.level,
    'Employee Class': r.empClass,
    'Function 1': r.function1,
    'Function/Plant': r.plant,
    'Location Name': r.location,
    'Asset': r.asset,
    'Cluster': r.cluster,
    'Gender': r.gender,
    'Date of Birth': r.dob,
    'HR Manager Name': r.hrName,
    'HR Manager EID': r.hrEid,
    'Management Board EID': r.boardEid,
    'Date of Joining': r.doj,
    'Date in Role': r.dir,
    'Date Promoted': r.prom,
    'Manager Since': r.mgrSince,
    'Email': r.email,
    'Photo URL': r.photo,
    'Matrix Manager EID(s)': r.matrix,
    'Cohort Tags': r.cohorts,
    'Current Status/Tag': r.status,
}));

const aoa = [HEADERS, tierRow, ...sample.map(row => HEADERS.map(h => row[h]))];

const ws = XLSX.utils.aoa_to_sheet(aoa);
ws['!cols'] = HEADERS.map(h => ({ wch: Math.max(14, h.length + 2) }));
ws['!freeze'] = { xSplit: 0, ySplit: 2 };

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Employees');

const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
writeFileSync(outPath, buf);
console.log(`wrote ${outPath} (${aoa.length} rows, ${HEADERS.length} cols)`);
