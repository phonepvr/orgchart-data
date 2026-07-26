// Central application state + derived data + region render dispatch.
//
// Pattern: renderers mutate nothing; actions (main.js) mutate `state` then
// call render(...regions). Each region renderer rebuilds its container's
// HTML from state. At ≤ a few thousand rows a full region rebuild is
// single-digit milliseconds, so there is no fine-grained DOM diffing.
import { sortEmployees } from './data.js';
import {
    computeAllUniqueByField, computeAvailableFilterFields, searchEmployees,
    applyFilters, sortTabular, computeCohortMetrics, computeDynamicGlobalMetrics,
    computeHeatmapStats,
} from './filters.js';
import { NUMERIC_FIELDS } from './constants.js';

export const state = {
    unlocked: false,
    appTab: 'org', // 'org', 'table', 'compare'
    data: [],
    employeeMap: {},
    activeEmployeeId: null,
    ceoId: null,

    // Filtering & Search
    searchQuery: '',
    isSearchOpen: false,
    showFilterPanel: false,
    filterMatchMode: 'and',
    filterConditions: [],
    openDropdown: null,
    sortConfigs: [{ field: 'TeamSize', dir: 'desc' }],
    activeCohortScale: null,

    loading: false,
    error: null,
    warnings: [],
    viewMode: 'direct',
    isSidebarOpen: true,
    // Resets every fresh upload (per handleFileUpload).
    showDataVerifyBanner: true,

    compareList: { blue: [], green: [], purple: [], orange: [], red: [] },
    contextMenu: null,
    printNodeId: null,

    // Compare view UI state (was local component state in React)
    compareActiveColor: 'blue',
    compareIndOpen: false,
    compareOrgOpen: true,
};

// Derived data, recomputed once per render pass. Kept in a module-level
// object so overlay code (tooltips) can read the latest values between renders.
export const derived = {};

export function refreshDerived() {
    const allUniqueByField = computeAllUniqueByField(state.data);
    const availableFilterFields = computeAvailableFilterFields(allUniqueByField);
    const baseFilteredData = applyFilters(state.data, state.filterConditions, state.filterMatchMode);
    const cohortMetrics = computeCohortMetrics(baseFilteredData, state.ceoId);

    derived.allUniqueByField = allUniqueByField;
    derived.availableFilterFields = availableFilterFields;
    derived.filterFieldsOrder = [...availableFilterFields, ...NUMERIC_FIELDS];
    derived.filteredSearch = searchEmployees(state.data, state.searchQuery);
    derived.baseFilteredData = baseFilteredData;
    derived.tabularSortedData = sortTabular(baseFilteredData, state.sortConfigs);
    derived.cohortMetrics = cohortMetrics;
    derived.dynamicGlobalMetrics = computeDynamicGlobalMetrics(baseFilteredData, cohortMetrics);
    derived.heatmapStats = computeHeatmapStats(baseFilteredData);

    const activeEmployee = state.employeeMap[state.activeEmployeeId];
    derived.activeEmployee = activeEmployee;
    derived.manager = activeEmployee?._managerId ? state.employeeMap[activeEmployee._managerId] : null;
    const visible = (ids) => (ids || []).map(id => state.employeeMap[id]).filter(Boolean)
        .filter(emp => state.filterConditions.length === 0 || baseFilteredData.find(f => f._id === emp._id))
        .sort((a, b) => sortEmployees(a, b, state.ceoId));
    derived.directReports = visible(activeEmployee?._directs);
    derived.matrixReports = visible(activeEmployee?._matrix);
    return derived;
}

// Region renderers are registered by main.js (avoids circular imports).
export const renderers = {};

// render('header', 'content', ...) or render('app') for all app regions.
// Valid regions: screen, header, banner, sidebar, content, overlays, print
const APP_REGIONS = ['header', 'banner', 'sidebar', 'content', 'overlays'];

export function render(...regions) {
    refreshDerived();
    const list = regions.includes('app') ? APP_REGIONS.concat(regions.filter(r => r !== 'app')) : regions;
    [...new Set(list)].forEach(r => { if (renderers[r]) renderers[r](); });
}
