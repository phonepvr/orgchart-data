// Pure data logic: parsing, normalization, insights, print pagination.
// Ported verbatim from the former React App.jsx — no DOM access here.
import {
    REQUIRED_COLUMNS, RECOMMENDED_COLUMNS, OPTIONAL_COLUMNS,
    NORMAL_PAGE_CAP, CONTINUATION_PAGE_CAP,
} from './constants.js';

// --- Format Helpers ---
export const formatNum = (num) => (num === 0 || num === '0' || !num) ? '-' : num;

export const formatJobTitle = (title) => {
    if (!title) return '';
    let t = String(title)
        .replace(/Ã¢Â€Â[”“-]/g, '-')
        .replace(/Ã¢[^\w\s]*/g, '-')
        .replace(/â€“|â€”/g, '-')
        .replace(/[–—]/g, '-')
        .replace(/\s*-\s*/g, ' - ')
        .replace(/\s+/g, ' ');

    return t
        .replace(/\bSenior Vice President\b/ig, 'SVP')
        .replace(/\bSenior VP\b/ig, 'SVP')
        .replace(/\bVice President\b/ig, 'VP')
        .replace(/\bGeneral Manager\b/ig, 'GM')
        .replace(/\bDeputy GM\b/ig, 'DGM')
        .replace(/\bSenior Manager\b/ig, 'Sr Manager')
        .replace(/Sr\.\s*Manager/ig, 'Sr Manager')
        .replace(/\bAssistant Manager\b/ig, 'Asst Manager')
        .trim();
};

export const splitSemicolonList = (v) => {
    if (v === undefined || v === null || v === '') return [];
    return String(v).split(';').map(s => s.trim()).filter(Boolean);
};

export const buildInitials = (name) => {
    return String(name || '?').split(/\s+/).filter(Boolean).map(n => n[0]).join('').substring(0, 2).toUpperCase();
};

export const deriveAge = (dob) => {
    if (!dob || !(dob instanceof Date) || isNaN(dob.getTime())) return null;
    const today = new Date();
    let age = today.getFullYear() - dob.getFullYear();
    const m = today.getMonth() - dob.getMonth();
    if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) age--;
    return age >= 0 && age < 120 ? age : null;
};

export const sortEmployees = (a, b, ceoId) => {
    if (a._id === ceoId) return -1;
    if (b._id === ceoId) return 1;
    const mcA = a._isMgmtCommittee ? 1 : 0;
    const mcB = b._isMgmtCommittee ? 1 : 0;
    if (mcA !== mcB) return mcB - mcA;
    const teamA = a._insights?.totalTeam || 0;
    const teamB = b._insights?.totalTeam || 0;
    if (teamA !== teamB) return teamB - teamA;
    return (a._formattedName || '').localeCompare(b._formattedName || '');
};

export const getMedian = (arr) => {
    if (!arr || arr.length === 0) return 0;
    const s = [...arr].sort((a,b) => a - b);
    const mid = Math.floor(s.length / 2);
    return s.length % 2 !== 0 ? s[mid] : s[mid - 1];
};

export const toProperCase = (str) => str ? str.replace(/\b\w+/g, txt => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase()) : '';

export const formatDisplayFirstLast = (name) => {
    if (!name) return '';
    let cleanName = String(name).replace(/\./g, '').trim();
    const parts = cleanName.split(/\s+/);

    // If it's only two parts, keep both even if one is an initial or prefix
    if (parts.length === 2) {
        return toProperCase(parts.join(' '));
    }

    let startIdx = 0;
    const skipRegex = /^(?:[A-Za-z]|mohd|mohammad|mohamad|mohammed|mohamed|muhammad|muhamad|muhammed|muhamed|md)$/i;

    while (startIdx < parts.length - 1 && skipRegex.test(parts[startIdx])) {
        startIdx++;
    }
    const sigParts = parts.slice(startIdx);

    let display = sigParts.join(' ');
    if (sigParts.length >= 3) display = `${sigParts[0]} ${sigParts[sigParts.length - 1]}`;
    return toProperCase(display);
};

export const parseExcelDate = (excelDate) => {
    if (excelDate === undefined || excelDate === null || excelDate === '') return null;
    if (typeof excelDate === 'number') return new Date(Math.round((excelDate - 25569) * 86400 * 1000));

    let dateStr = String(excelDate).trim();
    const parts = dateStr.split(/[-/]/);
    if (parts.length === 3) {
        let y, m, d;
        if (parts[2].length === 4) {
            d = parseInt(parts[0], 10);
            m = parseInt(parts[1], 10);
            y = parseInt(parts[2], 10);
        } else if (parts[0].length === 4) {
            y = parseInt(parts[0], 10);
            m = parseInt(parts[1], 10);
            d = parseInt(parts[2], 10);
        }
        if (y && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
            return new Date(y, m - 1, d);
        }
    }
    const fallbackDate = new Date(excelDate);
    return isNaN(fallbackDate.getTime()) ? null : fallbackDate;
};

export const formatDuration = (start, end) => {
    if (!start) return '-';
    let s = start;
    if (s instanceof Date && s.getFullYear() >= 9999) s = new Date();
    let e = end || new Date();
    if (e instanceof Date && e.getFullYear() >= 9999) e = new Date();
    let months = (e - s) / (1000 * 60 * 60 * 24 * 30.4375);
    if (months <= 0) return '< 1 mo';
    if (months < 1) return '< 1 mo';
    if (months < 12) return Math.round(months) + ' mos';
    return (months / 12).toFixed(1) + ' yrs';
};

export const isEA = (e) => {
    if (!e) return false;
    const title = String(e.jobTitle || '').toLowerCase();
    return title.includes('executive assistant') || title.includes('executive secretary') || title.includes('confidential secretary');
};

// --- Template Header Validation ---
export const validateHeaders = (rawRows) => {
    if (!rawRows || rawRows.length === 0) {
        return { ok: false, missingRequired: [...REQUIRED_COLUMNS], missingRecommended: [], missingOptional: [] };
    }
    const headers = Object.keys(rawRows[0]);
    // Accept legacy header names so files exported before the v3 rename still
    // pass the Required-column gate.
    const LEGACY_ALIASES = {
        'Position Text':       'Job Title',
        'Current Status/Tag':  'Current Status',
    };
    const has = (col) => headers.includes(col) || (LEGACY_ALIASES[col] && headers.includes(LEGACY_ALIASES[col]));
    return {
        ok: REQUIRED_COLUMNS.every(has),
        missingRequired: REQUIRED_COLUMNS.filter(c => !has(c)),
        missingRecommended: RECOMMENDED_COLUMNS.filter(c => !has(c)),
        missingOptional: OPTIONAL_COLUMNS.filter(c => !has(c)),
    };
};

// --- Row Normalizer ---
const NAME_STATUS_VALUES = new Set(['approved', 'unapproved']);
const VALID_STATUSES = new Set(['Active', 'WIP', 'Offered', 'Vacant']);

export const deriveNameStatus = (name) => {
    const k = String(name || '').trim().toLowerCase();
    return NAME_STATUS_VALUES.has(k) ? k : null;
};

export const normalizeCurrentStatus = (raw) => {
    const v = String(raw || '').trim();
    if (!v) return '';
    const match = [...VALID_STATUSES].find(s => s.toLowerCase() === v.toLowerCase());
    return match || v; // pass through unrecognized values; UI shows neutral chip
};

export const normalizeRow = (row) => {
    const get = (key) => {
        const val = row[key];
        if (val === undefined || val === null) return '';
        return typeof val === 'string' ? val.trim() : String(val).trim();
    };
    const name = get('Employee name');
    return {
        eid: get("Employee's Position Code"),
        name,
        _nameStatus: deriveNameStatus(name),
        managerEid: get("Line Manager's Position Code"),
        managerName: get('Line Manager Name'),
        jobTitle: formatJobTitle(get('Position Text') || get('Job Title')),
        level: get('Level'),
        employeeClass: get('Employee Class'),
        function1: get('Function 1'),
        functionPlant: get('Function/Plant'),
        location: get('Location Name'),
        asset: get('Asset'),
        cluster: get('Cluster'),
        gender: get('Gender'),
        dob: parseExcelDate(row['Date of Birth']),
        hrManagerName: get('HR Manager Name'),
        hrManagerEid: get('HR Manager EID'),
        mgmtBoardEid: get('Management Board EID'),
        dateOfJoining: parseExcelDate(row['Date of Joining']),
        dateInRole: parseExcelDate(row['Date in Role']),
        datePromoted: parseExcelDate(row['Date Promoted']),
        managerSince: parseExcelDate(row['Manager Since']),
        email: get('Email'),
        photoUrl: get('Photo URL'),
        matrixEids: splitSemicolonList(get('Matrix Manager EID(s)')),
        cohortTags: splitSemicolonList(get('Cohort Tags')),
        currentStatus: normalizeCurrentStatus(get('Current Status/Tag') || get('Current Status')),
    };
};

export const sha256Hex = async (text) => {
    const buf = new TextEncoder().encode(text);
    const hashBuf = await crypto.subtle.digest('SHA-256', buf);
    return Array.from(new Uint8Array(hashBuf)).map(b => b.toString(16).padStart(2, '0')).join('');
};

// --- Employee graph construction + insights ---
// Returns { data, employeeMap, ceoId } instead of setting React state.
export const processEmployeeData = (rawData) => {
    const empMap = {};
    const directReportsMap = {};
    const matrixReportsMap = {};

    // 1. Normalize each row + derive lookup keys
    rawData.forEach(rawRow => {
      const norm = normalizeRow(rawRow);
      if (!norm.eid || !norm.name) return; // require EID + name
      const id = norm.eid.toLowerCase();
      const emp = {
        ...norm,
        _id: id,
        _formattedName: formatDisplayFirstLast(norm.name),
        _formattedManagerName: formatDisplayFirstLast(norm.managerName),
        _initials: buildInitials(norm.name),
        _age: deriveAge(norm.dob),
      };
      const safeStr = (...parts) => parts.filter(Boolean).join(' ').toLowerCase();
      emp._searchString = safeStr(emp._formattedName, emp.eid, emp.jobTitle, emp.function1, emp.location, (emp.cohortTags || []).join(' '));
      empMap[id] = emp;
    });

    // 2. Build manager lookup from EID and resolve directs / matrix
    Object.values(empMap).forEach(emp => {
      const lmId = emp.managerEid ? emp.managerEid.toLowerCase() : '';
      if (lmId && empMap[lmId] && lmId !== emp._id) {
        emp._managerId = lmId;
        if (!directReportsMap[lmId]) directReportsMap[lmId] = [];
        directReportsMap[lmId].push(emp._id);
      }
      const matrixIds = (emp.matrixEids || [])
        .map(m => m.toLowerCase())
        .filter(m => m && m !== emp._id && empMap[m]);
      emp._matrixIds = matrixIds;
      matrixIds.forEach(mid => {
        if (!matrixReportsMap[mid]) matrixReportsMap[mid] = [];
        if (!matrixReportsMap[mid].includes(emp._id)) matrixReportsMap[mid].push(emp._id);
      });
    });

    // 3. Management Committee: own EID == own Management Board EID
    Object.values(empMap).forEach(emp => {
      const board = emp.mgmtBoardEid ? emp.mgmtBoardEid.toLowerCase() : '';
      emp._isMgmtCommittee = !!(board && board === emp._id);
    });

    // 4. Recursive insights
    const calculateInsights = (empId, visited = new Set()) => {
      if (visited.has(empId)) return empMap[empId]._insights;
      visited.add(empId);
      const directs = directReportsMap[empId] || [];
      const matrix = matrixReportsMap[empId] || [];

      let totalTeam = 0, directCount = 0, eaCount = 0;
      const genderCount = { male: 0, female: 0, other: 0 };
      const directGrades = {}, matrixGrades = {}, teamGrades = {};

      directs.forEach(childId => {
        const child = empMap[childId];
        if (!child) return;
        const childInsights = calculateInsights(childId, visited);
        const lvl = child.level || 'Unspecified';
        directGrades[lvl] = (directGrades[lvl] || 0) + 1;
        teamGrades[lvl] = (teamGrades[lvl] || 0) + 1;

        totalTeam += 1 + (childInsights ? childInsights.totalTeam : 0);
        if (childInsights) Object.entries(childInsights.teamGrades).forEach(([g, c]) => teamGrades[g] = (teamGrades[g] || 0) + c);

        if (!isEA(child)) {
          const gender = String(child.gender || '').toLowerCase();
          if (gender.startsWith('m')) genderCount.male++;
          else if (gender.startsWith('f')) genderCount.female++;
          else if (gender) genderCount.other++;
          directCount++;
        } else {
          eaCount++;
        }
      });
      matrix.forEach(childId => {
        const child = empMap[childId];
        if (child) {
          const lvl = child.level || 'Unspecified';
          matrixGrades[lvl] = (matrixGrades[lvl] || 0) + 1;
        }
      });

      const insights = { directCount, eaCount, matrixCount: matrix.length, totalTeam, directGrades, matrixGrades, teamGrades, genderCount };
      empMap[empId]._insights = insights;
      empMap[empId]._directs = directs;
      empMap[empId]._matrix = matrix;
      return insights;
    };
    Object.keys(empMap).forEach(id => calculateInsights(id));

    // 5. Peer benchmarks + share of manager team
    Object.values(empMap).forEach(emp => {
      const managerId = emp._managerId;
      if (managerId && empMap[managerId]) {
        const manager = empMap[managerId];
        const peers = (manager._directs || []).filter(id => id !== emp._id && !isEA(empMap[id]));
        const managerTeamSize = Math.max(1, manager._insights?.totalTeam || 1);
        const myBranchSize = 1 + (emp._insights?.totalTeam || 0);
        if (managerTeamSize > 0 && !isEA(emp)) {
          emp._insights.pctOfManagerTeam = Math.round((myBranchSize / managerTeamSize) * 100);
          emp._insights.managerValidDrCount = peers.length + 1;
        }
        if (peers.length > 0) {
          const peerDrs = peers.map(pId => empMap[pId]?._insights?.directCount || 0);
          emp._insights.peerMedianDirects = getMedian(peerDrs);
          emp._insights.peerMinDirects = Math.min(...peerDrs);
          emp._insights.peerMaxDirects = Math.max(...peerDrs);
        }
      }

      emp._tenureFormatted = emp.dateOfJoining ? formatDuration(emp.dateOfJoining) : '';
      emp._timeInRoleFormatted = emp.dateInRole ? formatDuration(emp.dateInRole) : '';
      emp._lastPromotionFormatted = emp.datePromoted ? formatDuration(emp.datePromoted) : '';
      emp._timeWithManagerFormatted = emp.managerSince ? formatDuration(emp.managerSince) : '';
    });

    // 6. Pick top node: largest team among roots (rows with empty Line Manager's Position Code), prefer MC
    const roots = Object.values(empMap).filter(e => !e._managerId);
    let topNode = null;
    if (roots.length > 0) {
      topNode = roots.slice().sort((a, b) => sortEmployees(a, b, null))[0];
    } else {
      topNode = Object.values(empMap).slice().sort((a, b) => sortEmployees(a, b, null))[0];
    }
    const computedCeoId = topNode ? topNode._id : null;

    const baseDataArr = Object.values(empMap).sort((a, b) => sortEmployees(a, b, computedCeoId));
    return { data: baseDataArr, employeeMap: empMap, ceoId: computedCeoId };
};

export const getCohortStats = (arr) => {
    if (arr.length === 0) return { count: 0 };
    const drs = arr.map(a => a._insights?.directCount || 0);
    const matrix = arr.map(a => a._insights?.matrixCount || 0);
    const teams = arr.map(a => a._insights?.totalTeam || 0);
    const totalReps = arr.map(a => (a._insights?.directCount || 0) + (a._insights?.matrixCount || 0));
    const nzMatrix = matrix.filter(m => m > 0);
    return {
        count: arr.length,
        drMin: Math.min(...drs), drMax: Math.max(...drs), drMedian: getMedian(drs),
        teamMin: Math.min(...teams), teamMax: Math.max(...teams), teamMedian: getMedian(teams),
        matrixMin: nzMatrix.length ? Math.min(...nzMatrix) : 0,
        matrixMax: nzMatrix.length ? Math.max(...nzMatrix) : 0,
        matrixMedian: getMedian(nzMatrix),
        matrixHasZeros: nzMatrix.length !== matrix.length,
        matrixNzCount: nzMatrix.length,
        totalRepMin: Math.min(...totalReps), totalRepMax: Math.max(...totalReps), totalRepMedian: getMedian(totalReps)
    };
};

// --- Print pagination (pure data) ---

// Smart column count for the subject-page side panes (0-12 only; above that we paginate).
export const sideColumns = (n) => (n <= 4 ? 1 : 2);

// Direct-reports grid column count. With no matrix pane the DR side gets the
// whole right half of the canvas, so widen toward ~4 rows (up to 3 cols) to
// fill the space instead of overflowing vertically into an empty right margin.
// With a matrix pane present the DR side is narrower, so cap at 2 cols.
export const drColumns = (n, hasMatrix) => {
    if (hasMatrix) return n <= 4 ? 1 : 2;
    if (n <= 2) return n || 1;   // 1-2 reports → 1-2 cols
    if (n <= 8) return 2;        // 3-8 reports → 2 cols (≤4 rows)
    return 3;                    // 9-12 reports → 3 cols (≤4 rows)
};

// Pure function: split one subject's reports across a 'subject' page plus
// 'continuation' pages. Returns a list of page descriptors:
//   { kind, subject, drs, matrix, drStart, drTotal }
export const planSubjectPages = (subject, employeeMap, ceoId) => {
    const drs = (subject._directs || []).map(id => employeeMap[id]).filter(Boolean).sort((a, b) => sortEmployees(a, b, ceoId));
    const matrix = (subject._matrix || []).map(id => employeeMap[id]).filter(Boolean).sort((a, b) => sortEmployees(a, b, ceoId));

    // When there's no matrix the DR pane gets the whole canvas and fits more;
    // when matrix is present the side panes share the canvas, so cap tighter.
    const firstPageDrCap = matrix.length > 0 ? NORMAL_PAGE_CAP : NORMAL_PAGE_CAP + 4; // 12

    // TODO: matrix overflow continuation is out of scope. If matrix.length
    // exceeds NORMAL_PAGE_CAP the surplus is clipped by .print-page overflow.
    // Extremely rare in current data; add matrix continuation pages if a real
    // customer hits it.

    const firstChunk = drs.slice(0, firstPageDrCap);
    const overflow = drs.slice(firstPageDrCap);

    const pages = [{
        kind: 'subject',
        subject,
        drs: firstChunk,
        matrix,
        drStart: 0,
        drTotal: drs.length,
    }];

    for (let i = 0; i < overflow.length; i += CONTINUATION_PAGE_CAP) {
        pages.push({
            kind: 'continuation',
            subject,
            drs: overflow.slice(i, i + CONTINUATION_PAGE_CAP),
            matrix: [],
            drStart: firstPageDrCap + i,
            drTotal: drs.length,
        });
    }
    return pages;
};

// Depth-first subject collection: a manager's page is immediately followed by
// the pages of their own reporting managers, so the PDF reads top-down one
// sub-tree at a time. Only people who actually have reports become subjects.
export const collectPrintSubjects = (emp, depth, employeeMap, ceoId) => {
    const subjects = [emp];
    if (depth <= 0) return subjects;
    const drs = (emp._directs || []).map(id => employeeMap[id]).filter(Boolean).sort((a, b) => sortEmployees(a, b, ceoId));
    drs.forEach(dr => {
        if ((dr._insights?.directCount || 0) > 0 || (dr._insights?.matrixCount || 0) > 0) {
            subjects.push(...collectPrintSubjects(dr, depth - 1, employeeMap, ceoId));
        }
    });
    return subjects;
};
