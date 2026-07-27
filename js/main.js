(function () {
'use strict';
const OS = window.OrgSense = window.OrgSense || {};
const { state, renderers, render, refreshDerived, sha256Hex, validateHeaders, processEmployeeData, defaultsForField, ACCESS_HASH, lockScreenHTML, uploadScreenHTML, headerInnerHTML, searchResultsHTML, bannerHTML, sidebarClass, sidebarInnerHTML, pillsBarHTML, orgViewHTML, tableViewHTML, compareViewHTML, printLayoutHTML, showInfoTooltip, scheduleHideInfoTooltip, cancelHideInfoTooltip, showGradeTooltip, scheduleHideGradeTooltip, cancelHideGradeTooltip, clearTooltips, renderContextMenu, avatarErrorFallback } = OS;
// Boot + event wiring. All interactions are delegated document-level
// listeners resolving data-action / data-input / data-change attributes, so
// region re-renders never need listener bookkeeping.
/* global XLSX */

const root = document.getElementById('root');
let currentScreen = null; // 'lock' | 'upload' | 'app'
let nextCondId = 1;
let resetTableScroll = false;

// --- Screen shells ---
const appShellHTML = () =>
    `<div class="min-h-screen w-full flex flex-col font-sans text-slate-800 bg-slate-100 overflow-hidden">` +
    `<div id="app-main" class="flex-col h-screen w-full overflow-hidden flex print:hidden">` +
    `<header id="header-region" class="bg-white border-b border-graphite-100 px-6 py-4 flex items-center justify-between shadow-sm z-30 flex-shrink-0"></header>` +
    `<div id="banner-region"></div>` +
    `<main class="flex-1 overflow-hidden flex flex-row w-full relative">` +
    `<aside id="sidebar-region" class="hidden"></aside>` +
    `<div class="flex-1 flex flex-col relative bg-slate-50 min-h-0 overflow-hidden" id="chart-container"></div>` +
    `</main></div>` +
    `<div id="print-region" class="hidden print:block w-full"></div>` +
    `</div>`;

function renderScreen() {
    const target = !state.unlocked ? 'lock' : (state.data.length === 0 ? 'upload' : 'app');
    if (target !== currentScreen) {
        currentScreen = target;
        clearTooltips();
        if (target === 'lock') {
            root.innerHTML = lockScreenHTML();
            wireLockScreen();
        } else if (target === 'upload') {
            root.innerHTML = uploadScreenHTML();
        } else {
            root.innerHTML = appShellHTML();
            render('app');
        }
    } else if (target === 'upload') {
        root.innerHTML = uploadScreenHTML();
    } else if (target === 'app') {
        render('app');
    }
}

// --- Region renderers (registered into state.js dispatch) ---
renderers.header = () => {
    document.getElementById('header-region').innerHTML = headerInnerHTML();
};
renderers.banner = () => {
    document.getElementById('banner-region').innerHTML = bannerHTML();
};
renderers.sidebar = () => {
    const el = document.getElementById('sidebar-region');
    // Preserve scroll positions and a focused filter input across the rebuild.
    const scrollEl = el.querySelector('#sidebar-scroll');
    const scrollTop = scrollEl ? scrollEl.scrollTop : 0;
    const ddEl = el.querySelector('[data-dropdown-list]');
    const ddScrollTop = ddEl ? ddEl.scrollTop : 0;
    const active = document.activeElement;
    const focusCondId = (active && el.contains(active) && active.dataset.input === 'filter-value')
        ? active.dataset.condId : null;

    el.className = sidebarClass();
    el.innerHTML = sidebarInnerHTML();

    const newScroll = el.querySelector('#sidebar-scroll');
    if (newScroll) newScroll.scrollTop = scrollTop;
    const newDd = el.querySelector('[data-dropdown-list]');
    if (newDd) newDd.scrollTop = ddScrollTop;
    if (focusCondId) {
        const input = el.querySelector(`[data-input="filter-value"][data-cond-id="${focusCondId}"]`);
        if (input) input.focus();
    }
};
renderers.content = () => {
    const el = document.getElementById('chart-container');
    const orgEl = el.querySelector('#org-scroll');
    const orgScroll = orgEl ? orgEl.scrollTop : 0;
    const tblEl = el.querySelector('#table-scroll');
    const tblScroll = tblEl ? tblEl.scrollTop : 0;

    el.innerHTML = pillsBarHTML() + tableViewHTML() + orgViewHTML() +
        `<div class="w-full h-full flex-col overflow-hidden p-0 bg-slate-50 min-h-0 ${state.appTab === 'compare' ? 'flex' : 'hidden'}">${compareViewHTML()}</div>`;

    const newOrg = el.querySelector('#org-scroll');
    if (newOrg) newOrg.scrollTop = orgScroll;
    const newTbl = el.querySelector('#table-scroll');
    if (newTbl) newTbl.scrollTop = resetTableScroll ? 0 : tblScroll;
    resetTableScroll = false;
};
renderers.overlays = () => renderContextMenu();
renderers.print = () => {
    const appMain = document.getElementById('app-main');
    const printRegion = document.getElementById('print-region');
    if (!appMain || !printRegion) return;
    if (state.printNodeId) {
        appMain.classList.remove('flex');
        appMain.classList.add('hidden');
        printRegion.innerHTML = printLayoutHTML(state.printNodeId, state.employeeMap, state.ceoId);
    } else {
        appMain.classList.remove('hidden');
        appMain.classList.add('flex');
        printRegion.innerHTML = '';
    }
};

// --- Shared behaviors ---
function handleEmployeeSelect(id) {
    state.activeEmployeeId = id;
    state.appTab = 'org';
    state.viewMode = 'direct';
    render('app');
}

function updateFilterCondition(id, key, val) {
    state.filterConditions = state.filterConditions.map(c => {
        if (c.id !== Number(id)) return c;
        if (key === 'field') {
            const d = defaultsForField(val);
            return { ...c, field: val, operator: d.operator, value: d.value };
        }
        return { ...c, [key]: val };
    });
    state.appTab = 'table';
    resetTableScroll = true;
}

function removeFilterCondition(id) {
    state.filterConditions = state.filterConditions.filter(c => c.id !== Number(id));
    resetTableScroll = true;
}

async function handleFileUpload(file) {
    state.loading = true; state.error = null; state.warnings = []; state.showDataVerifyBanner = true;
    renderScreen();
    try {
        const buffer = await file.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
        const sheetName = workbook.SheetNames.find(n => n.toLowerCase() === 'employees') || workbook.SheetNames[0];
        const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
        if (rawData.length === 0) throw new Error("Uploaded Excel file is empty.");

        const validation = validateHeaders(rawData);
        if (!validation.ok) {
            throw new Error(`Missing required column${validation.missingRequired.length > 1 ? 's' : ''}: ${validation.missingRequired.join(', ')}`);
        }
        // Filter out the template's "Required/Recommended/Optional" category row
        const labelMarkers = new Set(['required', 'recommended', 'optional']);
        const cleanedData = rawData.filter(row => {
            const eid = String(row["Employee's Position Code"] || '').trim().toLowerCase();
            return eid && !labelMarkers.has(eid);
        });
        if (cleanedData.length === 0) throw new Error("No employee rows found after parsing.");
        const w = [];
        if (validation.missingRecommended.length > 0) {
            w.push(`Missing recommended column${validation.missingRecommended.length > 1 ? 's' : ''}: ${validation.missingRecommended.join(', ')}. Some UI elements will be hidden.`);
        }
        state.warnings = w;
        const { data, employeeMap, ceoId } = processEmployeeData(cleanedData);
        state.data = data;
        state.employeeMap = employeeMap;
        if (ceoId) {
            state.activeEmployeeId = ceoId;
            state.ceoId = ceoId;
        }
    } catch (err) {
        state.error = err.message || "Failed to process file.";
    } finally {
        state.loading = false;
        renderScreen();
    }
}

// Post-action lifecycle: replicates the React useEffect scroll behaviors.
function afterAction(prev) {
    if (state.appTab === 'org' && state.activeEmployeeId &&
        (prev.appTab !== 'org' || prev.activeEmployeeId !== state.activeEmployeeId)) {
        setTimeout(() => {
            const activeEl = document.getElementById('active-employee-card');
            if (activeEl) activeEl.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
        }, 100);
    }
    if (state.appTab === 'table' && prev.appTab !== 'table' && state.activeEmployeeId) {
        setTimeout(() => {
            const activeEl = document.getElementById(`table-row-${state.activeEmployeeId}`);
            const container = document.getElementById('table-scroll');
            if (activeEl && container) {
                container.scrollTo({ top: Math.max(0, activeEl.offsetTop - 45), behavior: 'smooth' });
            }
        }, 100);
    }
}

// --- Click actions ---
const clickActions = {
    noop() {},
    tab(el) {
        state.appTab = el.dataset.tab;
        if (el.dataset.tab === 'org' || el.dataset.tab === 'table') state.activeCohortScale = null;
        if (el.dataset.tab === 'org' && state.activeEmployeeId) state.viewMode = 'direct';
        render('app');
    },
    'go-top'() { handleEmployeeSelect(state.ceoId); },
    'search-select'(el) {
        state.searchQuery = '';
        state.isSearchOpen = false;
        handleEmployeeSelect(el.dataset.id);
    },
    'card-click'(el) { handleEmployeeSelect(el.dataset.id); },
    'select-direct'(el) {
        state.activeEmployeeId = el.dataset.id;
        state.viewMode = 'direct';
        render('app');
    },
    'select-matrix'(el) {
        // Matches the React behavior: selecting matrix on a non-active card
        // changes the active employee, and the tab effect resets to 'direct'.
        if (el.dataset.id === state.activeEmployeeId) {
            state.viewMode = 'matrix';
        } else {
            state.activeEmployeeId = el.dataset.id;
            state.viewMode = 'direct';
        }
        render('app');
    },
    'dismiss-banner'() { state.showDataVerifyBanner = false; render('banner'); },
    'toggle-sidebar'() { state.isSidebarOpen = !state.isSidebarOpen; render('sidebar'); },
    'toggle-filter-panel'() { state.showFilterPanel = !state.showFilterPanel; render('sidebar'); },
    'match-mode'(el) { state.filterMatchMode = el.dataset.mode; resetTableScroll = true; render('app'); },
    'add-filter'() {
        const order = [...(refreshDerived().filterFieldsOrder)];
        if (order.length === 0) return;
        const nextField = order[state.filterConditions.length % order.length];
        const { operator, value } = defaultsForField(nextField);
        state.filterConditions = [...state.filterConditions, { id: nextCondId++, field: nextField, operator, value }];
        state.appTab = 'table';
        if (!state.isSidebarOpen) state.isSidebarOpen = true;
        resetTableScroll = true;
        render('app');
    },
    'filter-dropdown-toggle'(el) {
        const id = Number(el.dataset.condId);
        state.openDropdown = state.openDropdown === id ? null : id;
        render('sidebar');
    },
    'filter-remove'(el) { removeFilterCondition(el.dataset.condId); render('app'); },
    'cohort-tile'(el) {
        const type = el.dataset.key;
        const newFilters = state.filterConditions.filter(f => f.field !== 'Cohort Tag' && f.field !== 'Mgmt Committee');
        if (type === 'Mgmt Committee') {
            newFilters.push({ id: nextCondId++, field: 'Mgmt Committee', operator: 'in', value: ['Yes'] });
        } else {
            newFilters.push({ id: nextCondId++, field: 'Cohort Tag', operator: 'in', value: [type] });
        }
        state.filterConditions = newFilters;
        state.activeCohortScale = type;
        state.appTab = 'table';
        resetTableScroll = true;
        render('app');
    },
    'cohort-scale-close'() { state.activeCohortScale = null; render('sidebar'); },
    'pill-remove'(el) {
        const condId = Number(el.dataset.condId);
        if (el.dataset.type === 'single') {
            removeFilterCondition(condId);
        } else {
            const cond = state.filterConditions.find(c => c.id === condId);
            if (cond) {
                const newVals = cond.value.filter(v => v !== el.dataset.val);
                if (newVals.length === 0) removeFilterCondition(condId);
                else updateFilterCondition(condId, 'value', newVals);
            }
        }
        render('app');
    },
    'clear-filters'() {
        state.filterConditions = [];
        state.activeCohortScale = null;
        resetTableScroll = true;
        render('app');
    },
    sort(el) {
        const field = el.dataset.field;
        const existingIdx = state.sortConfigs.findIndex(c => c.field === field);
        if (existingIdx === -1) {
            state.sortConfigs = [...state.sortConfigs, { field, dir: 'asc' }];
        } else {
            const newConfigs = [...state.sortConfigs];
            if (newConfigs[existingIdx].dir === 'asc') newConfigs[existingIdx] = { field, dir: 'desc' };
            else newConfigs.splice(existingIdx, 1);
            state.sortConfigs = newConfigs;
        }
        resetTableScroll = true;
        render('content');
    },
    'table-row'(el) { handleEmployeeSelect(el.dataset.id); },
    'compare-tab'(el) { state.compareActiveColor = el.dataset.color; render('content'); },
    'compare-accordion'(el) {
        if (el.dataset.which === 'ind') state.compareIndOpen = !state.compareIndOpen;
        else state.compareOrgOpen = !state.compareOrgOpen;
        render('content');
    },
    'compare-color'(el) {
        const { color, empId } = el.dataset;
        const group = state.compareList[color] || [];
        if (group.length >= 4 && !group.includes(empId)) {
            alert('Maximum 4 employees per group.');
        } else if (!group.includes(empId)) {
            state.compareList = { ...state.compareList, [color]: [...group, empId] };
        }
        state.contextMenu = null;
        render('content', 'overlays');
    },
    'print-structure'(el) {
        state.printNodeId = el.dataset.empId;
        state.contextMenu = null;
        render('overlays', 'print');
        setTimeout(() => { window.print(); }, 500);
    },
};

// --- Wire delegated listeners ---
function wireGlobalListeners() {
    document.addEventListener('click', (e) => {
        const el = e.target.closest('[data-action]');
        if (!el) return;
        const handler = clickActions[el.dataset.action];
        if (!handler) return;
        const prev = { appTab: state.appTab, activeEmployeeId: state.activeEmployeeId };
        handler(el);
        afterAction(prev);
    });

    document.addEventListener('change', (e) => {
        const t = e.target;
        if (t.id === 'file-upload') {
            if (t.files && t.files[0]) handleFileUpload(t.files[0]);
            return;
        }
        const kind = t.dataset.change;
        if (!kind) return;
        const prev = { appTab: state.appTab, activeEmployeeId: state.activeEmployeeId };
        if (kind === 'filter-field') {
            updateFilterCondition(t.dataset.condId, 'field', t.value);
            render('app');
        } else if (kind === 'filter-operator') {
            updateFilterCondition(t.dataset.condId, 'operator', t.value);
            render('app');
        } else if (kind === 'filter-check') {
            const cond = state.filterConditions.find(c => c.id === Number(t.dataset.condId));
            if (cond) {
                const cur = Array.isArray(cond.value) ? cond.value : [];
                const newVals = t.checked ? [...cur, t.dataset.val] : cur.filter(v => v !== t.dataset.val);
                updateFilterCondition(t.dataset.condId, 'value', newVals);
                render('app');
            }
        }
        afterAction(prev);
    });

    document.addEventListener('input', (e) => {
        const t = e.target;
        const kind = t.dataset.input;
        if (kind === 'search') {
            state.searchQuery = t.value;
            state.isSearchOpen = true;
            refreshDerived();
            const sr = document.getElementById('search-results');
            if (sr) sr.innerHTML = searchResultsHTML();
        } else if (kind === 'filter-value') {
            const prev = { appTab: state.appTab, activeEmployeeId: state.activeEmployeeId };
            updateFilterCondition(t.dataset.condId, 'value', t.value);
            // Sidebar re-render preserves this input's focus (renderers.sidebar).
            render('app');
            afterAction(prev);
        }
    });

    document.addEventListener('focusin', (e) => {
        if (e.target.dataset && e.target.dataset.input === 'search') {
            state.isSearchOpen = true;
            refreshDerived();
            const sr = document.getElementById('search-results');
            if (sr) sr.innerHTML = searchResultsHTML();
        }
    });

    // Global outside-click handling (was the React handleClickOutside effect).
    document.addEventListener('mousedown', (e) => {
        const sw = document.getElementById('search-wrapper');
        if (sw && !sw.contains(e.target) && state.isSearchOpen) {
            state.isSearchOpen = false;
            const sr = document.getElementById('search-results');
            if (sr) sr.innerHTML = '';
        }
        if (!e.target.closest('.filter-dropdown-wrapper') && state.openDropdown !== null) {
            state.openDropdown = null;
            render('sidebar');
        }
        if (!e.target.closest('.context-menu') && state.contextMenu) {
            state.contextMenu = null;
            render('overlays');
        }
    });

    document.addEventListener('contextmenu', (e) => {
        const t = e.target.closest('[data-ctx]');
        if (t) {
            e.preventDefault();
            state.contextMenu = { x: e.clientX, y: e.clientY, empId: t.dataset.ctx };
            render('overlays');
        }
    });

    // Hover tooltips + privacy chip popup.
    let shownTip = null;
    document.addEventListener('mouseover', (e) => {
        const info = e.target.closest('[data-tip="info"]');
        if (info) {
            const key = `info:${info.dataset.id}`;
            if (shownTip !== key) { showInfoTooltip(info.dataset.id, info); shownTip = key; }
            else cancelHideInfoTooltip();
            return;
        }
        const grade = e.target.closest('[data-tip="grade"]');
        if (grade) {
            const key = `grade:${grade.dataset.type}:${grade.dataset.id}`;
            if (shownTip !== key) { showGradeTooltip(grade.dataset.id, grade.dataset.type, grade); shownTip = key; }
            else cancelHideGradeTooltip();
            return;
        }
        if (e.target.closest('#info-tooltip-layer')) { cancelHideInfoTooltip(); return; }
        if (e.target.closest('#grade-tooltip-layer')) { cancelHideGradeTooltip(); return; }
        const priv = e.target.closest('[data-hover="privacy"]');
        if (priv) {
            const pop = priv.querySelector('[data-privacy-pop]');
            if (pop) pop.classList.remove('hidden');
        }
    });
    document.addEventListener('mouseout', (e) => {
        const to = e.relatedTarget;
        const info = e.target.closest('[data-tip="info"]');
        if (info && !(to && info.contains(to))) { scheduleHideInfoTooltip(); shownTip = null; }
        const grade = e.target.closest('[data-tip="grade"]');
        if (grade && !(to && grade.contains(to))) { scheduleHideGradeTooltip(); shownTip = null; }
        const il = e.target.closest('#info-tooltip-layer');
        if (il && !(to && il.contains(to))) scheduleHideInfoTooltip();
        const gl = e.target.closest('#grade-tooltip-layer');
        if (gl && !(to && gl.contains(to))) scheduleHideGradeTooltip();
        const priv = e.target.closest('[data-hover="privacy"]');
        if (priv && !(to && priv.contains(to))) {
            const pop = priv.querySelector('[data-privacy-pop]');
            if (pop) pop.classList.add('hidden');
        }
    });

    // Upload drag & drop (class toggles instead of re-render, to avoid
    // dragleave churn from replacing the element under the cursor).
    const dragOn = (dz) => { dz.classList.add('border-red-600', 'bg-red-50'); dz.classList.remove('border-graphite-200'); };
    const dragOff = (dz) => { dz.classList.remove('border-red-600', 'bg-red-50'); dz.classList.add('border-graphite-200'); };
    document.addEventListener('dragover', (e) => {
        const dz = e.target.closest('#dropzone');
        if (dz) { e.preventDefault(); dragOn(dz); }
    });
    document.addEventListener('dragleave', (e) => {
        const dz = e.target.closest('#dropzone');
        if (dz && !(e.relatedTarget && dz.contains(e.relatedTarget))) dragOff(dz);
    });
    document.addEventListener('drop', (e) => {
        const dz = e.target.closest('#dropzone');
        if (dz) {
            e.preventDefault();
            dragOff(dz);
            if (e.dataTransfer.files && e.dataTransfer.files[0]) handleFileUpload(e.dataTransfer.files[0]);
        }
    });

    // Avatar photo load errors -> initials/placeholder fallback.
    document.addEventListener('error', (e) => {
        if (e.target.tagName === 'IMG') {
            const container = e.target.closest('[data-avatar]');
            if (container) avatarErrorFallback(container, state.employeeMap);
        }
    }, true);

    // Print lifecycle.
    window.addEventListener('afterprint', () => {
        if (state.printNodeId) {
            state.printNodeId = null;
            render('print');
        }
    });
}

// --- Lock screen wiring (local DOM state, no region re-render) ---
function wireLockScreen() {
    const form = document.getElementById('lock-form');
    const pwd = document.getElementById('lock-pwd');
    const btn = document.getElementById('lock-submit');
    const err = document.getElementById('lock-err');
    const setBtn = (busy) => {
        const enabled = !busy && !!pwd.value;
        btn.disabled = !enabled;
        btn.className = `w-full py-3 rounded-brand font-sans font-semibold text-white tracking-wide transition-colors duration-brand-fast inline-flex items-center justify-center gap-2 ${enabled ? 'bg-red-600 hover:bg-red-700 cursor-pointer' : 'bg-graphite-300 cursor-not-allowed'}`;
        btn.innerHTML = busy ? 'Verifying…' : 'Unlock <span aria-hidden="true">→</span>';
        pwd.disabled = busy;
    };
    pwd.addEventListener('input', () => { err.classList.add('hidden'); setBtn(false); });
    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!pwd.value) return;
        setBtn(true);
        err.classList.add('hidden');
        try {
            const hash = await sha256Hex(pwd.value);
            if (hash === ACCESS_HASH) {
                state.unlocked = true;
                renderScreen();
                return;
            }
            err.textContent = 'Incorrect password.';
            err.classList.remove('hidden');
            pwd.value = '';
        } catch (ex) {
            err.textContent = 'Password check failed: ' + ex.message;
            err.classList.remove('hidden');
        }
        setBtn(false);
        pwd.focus();
    });
    pwd.focus();
}

// --- Boot ---
wireGlobalListeners();
renderScreen();
// Signals js/diag.js that startup succeeded (renderScreen just replaced the
// static boot-fallback markup inside #root).
window.__orgSenseBooted = true;

})();
