// Filtering, sorting, and cohort/benchmark metric computation.
// Pure functions over the employee data array — ported verbatim from the
// former React useMemo bodies.
import { FILTER_FIELD_MAP, MULTI_SELECT_FIELDS, NUMERIC_FIELDS } from './constants.js';
import { getMedian, getCohortStats, isEA } from './data.js';

export const computeAllUniqueByField = (data) => {
    const out = {};
    Object.entries(FILTER_FIELD_MAP).forEach(([label, key]) => {
        out[label] = [...new Set(data.map(emp => emp[key]).filter(Boolean))].sort((a, b) => a.localeCompare(b));
    });
    const allCohorts = new Set();
    data.forEach(emp => (emp.cohortTags || []).forEach(t => t && allCohorts.add(t)));
    out['Cohort Tag'] = [...allCohorts].sort((a, b) => a.localeCompare(b));
    out['Mgmt Committee'] = ['Yes', 'No'];
    return out;
};

export const computeAvailableFilterFields = (allUniqueByField) =>
    MULTI_SELECT_FIELDS.filter(f => (allUniqueByField[f] || []).length > 0);

export const searchEmployees = (data, searchQuery) => {
    if (!searchQuery) return [];
    const query = searchQuery.toLowerCase();
    // Using precomputed formatted names handles queries faster
    return data.filter(emp => emp._searchString.includes(query)).slice(0, 5);
};

// Decoupled filtering logic
export const applyFilters = (data, filterConditions, filterMatchMode) => {
    if (filterConditions.length === 0) return data;
    return data.filter(emp => {
        const results = filterConditions.map(cond => {
            if (NUMERIC_FIELDS.includes(cond.field)) {
                if (cond.value === '' || cond.value === null) return false;
                const numVal = Number(cond.value);
                if (isNaN(numVal)) return false;
                let empVal = 0;
                if (cond.field === 'Team Size') empVal = emp._insights?.totalTeam || 0;
                else if (cond.field === 'DR Size') empVal = emp._insights?.directCount || 0;
                else if (cond.field === 'Total Reportees') empVal = (emp._insights?.directCount || 0) + (emp._insights?.matrixCount || 0) + (emp._insights?.eaCount || 0);
                if (cond.operator === '=') return empVal === numVal;
                if (cond.operator === '>') return empVal > numVal;
                if (cond.operator === '<') return empVal < numVal;
                return false;
            }
            if (!Array.isArray(cond.value) || cond.value.length === 0) return false;
            if (cond.field === 'Cohort Tag') {
                return (emp.cohortTags || []).some(t => cond.value.includes(t));
            }
            if (cond.field === 'Mgmt Committee') {
                const mc = emp._isMgmtCommittee ? 'Yes' : 'No';
                return cond.value.includes(mc);
            }
            const key = FILTER_FIELD_MAP[cond.field];
            if (!key) return false;
            return cond.value.includes(emp[key] || '');
        });
        if (results.length === 0) return false;
        return filterMatchMode === 'and' ? results.every(r => r) : results.some(r => r);
    });
};

// Tabular specific sorted view
export const sortTabular = (baseFilteredData, sortConfigs) => {
    let filtered = [...baseFilteredData];
    if (sortConfigs.length > 0) {
        filtered.sort((a, b) => {
            for (let config of sortConfigs) {
                let valA, valB;
                switch (config.field) {
                    case 'Employee': valA = a._formattedName; valB = b._formattedName; break;
                    case 'Level': valA = a.level || ''; valB = b.level || ''; break;
                    case 'Status': valA = a.currentStatus || ''; valB = b.currentStatus || ''; break;
                    case 'JobTitle': valA = a.jobTitle || ''; valB = b.jobTitle || ''; break;
                    case 'Function1': valA = a.function1 || ''; valB = b.function1 || ''; break;
                    case 'Location': valA = a.location || ''; valB = b.location || ''; break;
                    case 'DRSize': valA = a._insights?.directCount || 0; valB = b._insights?.directCount || 0; break;
                    case 'MatrixSize': valA = a._insights?.matrixCount || 0; valB = b._insights?.matrixCount || 0; break;
                    case 'TeamSize': valA = a._insights?.totalTeam || 0; valB = b._insights?.totalTeam || 0; break;
                    case 'Manager': valA = a._formattedManagerName; valB = b._formattedManagerName; break;
                    default: valA = ''; valB = '';
                }
                if (valA === valB) continue;
                let cmp = (typeof valA === 'string' && typeof valB === 'string') ? valA.localeCompare(valB) : (valA > valB ? 1 : -1);
                return config.dir === 'asc' ? cmp : -cmp;
            }
            return 0;
        });
    }
    return filtered;
};

export const computeCohortMetrics = (baseFilteredData, ceoId) => {
    const mc = baseFilteredData.filter(e => e._isMgmtCommittee && e._id !== ceoId);
    const tagBuckets = {};
    baseFilteredData.forEach(emp => {
        (emp.cohortTags || []).forEach(t => {
            if (!t) return;
            if (!tagBuckets[t]) tagBuckets[t] = [];
            tagBuckets[t].push(emp);
        });
    });
    const out = { 'Mgmt Committee': getCohortStats(mc) };
    Object.entries(tagBuckets).forEach(([tag, arr]) => {
        out[tag] = getCohortStats(arr);
    });
    return out;
};

export const computeDynamicGlobalMetrics = (baseFilteredData, cohortMetrics) => {
    const buckets = {};
    baseFilteredData.forEach(emp => {
        if (isEA(emp)) return;
        const lvl = emp.level;
        if (!lvl) return;
        if (!buckets[lvl]) buckets[lvl] = { drs: [], teams: [] };
        buckets[lvl].drs.push(emp._insights?.directCount || 0);
        buckets[lvl].teams.push(emp._insights?.totalTeam || 0);
    });
    const levelMetrics = {};
    Object.entries(buckets).forEach(([lvl, b]) => {
        levelMetrics[lvl] = {
            drMin: Math.min(...b.drs), drMax: Math.max(...b.drs), drMedian: getMedian(b.drs),
            teamMin: Math.min(...b.teams), teamMax: Math.max(...b.teams), teamMedian: getMedian(b.teams)
        };
    });
    return {
        level: levelMetrics,
        mgmtCommittee: cohortMetrics['Mgmt Committee']
    };
};

export const computeHeatmapStats = (baseFilteredData) => {
    const buckets = {};
    baseFilteredData.forEach(emp => {
        if (emp._insights?.directCount > 0) {
            const key = emp.function1 || emp.location;
            if (!key) return;
            if (!buckets[key]) buckets[key] = [];
            buckets[key].push(emp._insights.directCount);
        }
    });
    return Object.entries(buckets)
        .map(([d, drs]) => ({ dept: d, medianDr: getMedian(drs), count: drs.length }))
        .sort((a, b) => b.medianDr - a.medianDr)
        .slice(0, 10);
};

export const defaultsForField = (field) => {
    if (NUMERIC_FIELDS.includes(field)) return { operator: '=', value: '' };
    return { operator: 'in', value: [] };
};
