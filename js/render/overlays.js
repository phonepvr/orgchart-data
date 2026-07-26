// Overlays: hover tooltips (spotlight + grade summaries) and the right-click
// context menu. These live in fixed body-level layers (#info-tooltip-layer,
// #grade-tooltip-layer, #context-menu-layer) that region re-renders never
// touch — the vanilla equivalent of the former React portals.
import { esc } from '../util.js';
import { COMPARE_COLORS } from '../constants.js';
import { icon } from '../icons.js';
import { spotlightTooltipHTML, gradeTooltipHTML } from './spotlight.js';
import { state, derived } from '../state.js';

const infoLayer = () => document.getElementById('info-tooltip-layer');
const gradeLayer = () => document.getElementById('grade-tooltip-layer');
const menuLayer = () => document.getElementById('context-menu-layer');

let hideTimeout = null;
let hideGradeTimeout = null;

const styleString = (style) => Object.entries(style)
    .map(([k, v]) => `${k.replace(/[A-Z]/g, c => '-' + c.toLowerCase())}: ${v}`)
    .join('; ');

// --- Spotlight (info) tooltip ---
export const showInfoTooltip = (empId, triggerEl) => {
    clearTimeout(hideTimeout);
    const employee = state.employeeMap[empId];
    if (!employee) return;
    const insights = employee._insights || { genderCount: { male: 0, female: 0, other: 0 } };
    const isIndividualContributor = insights.directCount === 0 && (insights.eaCount || 0) === 0 && insights.matrixCount === 0;

    const cardRect = triggerEl.closest('.group').getBoundingClientRect();
    const tooltipWidth = 360;

    const style = { overflowY: 'auto' };

    // Determine horizontal side based on space, preferring the right side
    if (cardRect.right + tooltipWidth + 20 > window.innerWidth) {
        style.left = `${Math.max(10, cardRect.left - tooltipWidth - 10)}px`;
    } else {
        style.left = `${cardRect.right + 10}px`;
    }

    if (isIndividualContributor) {
        const isTopHalf = cardRect.top < window.innerHeight / 2;
        if (isTopHalf) {
            style.top = `${cardRect.top}px`;
            style.maxHeight = `${window.innerHeight - cardRect.top - 20}px`;
        } else {
            style.bottom = `${Math.max(10, window.innerHeight - cardRect.bottom)}px`;
            style.maxHeight = `${cardRect.bottom - 90}px`; // Accounting for header height approx 80px
        }
    } else {
        // Managers: Anchor to top below header
        style.top = `80px`;
        style.maxHeight = `calc(100vh - 100px)`;
    }

    infoLayer().innerHTML =
        `<div style="${styleString(style)}" class="fixed w-[360px] bg-white rounded-xl shadow-[0_0_40px_rgba(0,0,0,0.2)] border border-slate-200 p-0 text-sm overflow-hidden flex flex-col animate-scale-in z-[99999]">` +
        spotlightTooltipHTML(employee, state.ceoId, derived.dynamicGlobalMetrics) +
        `</div>`;
};

export const scheduleHideInfoTooltip = () => {
    hideTimeout = setTimeout(() => { infoLayer().innerHTML = ''; }, 200);
};
export const cancelHideInfoTooltip = () => clearTimeout(hideTimeout);

// --- Grade summary tooltip ---
export const showGradeTooltip = (empId, type, triggerEl) => {
    clearTimeout(hideGradeTimeout);
    const employee = state.employeeMap[empId];
    if (!employee) return;

    const pillRect = triggerEl.getBoundingClientRect();
    const cardRect = triggerEl.closest('.group').getBoundingClientRect();
    const tooltipWidth = 192;

    const style = {};

    let h = (type === 'team') ? 'right' : 'left';

    if (h === 'left' && cardRect.left - tooltipWidth - 10 < 0) h = 'right';
    if (h === 'right' && cardRect.right + tooltipWidth + 10 > window.innerWidth) h = 'left';

    if (h === 'right') {
        style.left = `${cardRect.right + 10}px`;
    } else {
        style.left = `${cardRect.left - tooltipWidth - 10}px`;
    }

    const isTopHalf = cardRect.top < window.innerHeight / 2;
    if (isTopHalf) {
        style.top = `${pillRect.top - 5}px`;
    } else {
        style.bottom = `${Math.max(20, window.innerHeight - cardRect.bottom)}px`;
    }

    gradeLayer().innerHTML =
        `<div style="${styleString(style)}" class="fixed w-48 bg-white rounded-lg shadow-[0_0_20px_rgba(0,0,0,0.15)] border border-slate-200 text-sm overflow-hidden animate-scale-in z-[99999]">` +
        gradeTooltipHTML(employee, type) +
        `</div>`;
};

export const scheduleHideGradeTooltip = () => {
    hideGradeTimeout = setTimeout(() => { gradeLayer().innerHTML = ''; }, 200);
};
export const cancelHideGradeTooltip = () => clearTimeout(hideGradeTimeout);

export const clearTooltips = () => {
    infoLayer().innerHTML = '';
    gradeLayer().innerHTML = '';
};

// --- Right-click context menu (Add to Compare / Print Structure) ---
export const renderContextMenu = () => {
    const cm = state.contextMenu;
    if (!cm) { menuLayer().innerHTML = ''; return; }
    menuLayer().innerHTML =
        `<div class="fixed bg-white border border-slate-200 shadow-xl rounded-xl p-3 z-[999999] animate-scale-in context-menu" style="top: ${cm.y}px; left: ${cm.x}px">` +
        `<div class="text-xs font-bold mb-3 text-slate-500 uppercase tracking-wider">Add to Compare</div>` +
        `<div class="flex gap-2.5 mb-3">` +
        COMPARE_COLORS.map(c =>
            `<button data-action="compare-color" data-color="${c.id}" data-emp-id="${esc(cm.empId)}" class="w-6 h-6 rounded-md ${c.bg} shadow-sm hover:scale-110 transition-transform hover:ring-2 hover:ring-offset-1 hover:ring-${c.id}-400"></button>`
        ).join('') +
        `</div>` +
        `<div class="w-full h-px bg-slate-100 my-2"></div>` +
        `<button data-action="print-structure" data-emp-id="${esc(cm.empId)}" class="w-full flex items-center justify-center gap-2 text-[11px] font-bold text-slate-600 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 py-1.5 rounded transition-colors">` +
        `${icon('Printer', { size: 12 })} Print Structure</button>` +
        `</div>`;
};
