(function () {
'use strict';
const OS = window.OrgSense = window.OrgSense || {};
const { esc, STATUS_STYLES, NAME_STATUS_TINT, formatNum, icon, avatarHTML, statusChipHTML, nameStatusChipHTML, state, derived } = OS;
// Org chart view: employee cards, manager/active/reports layout.

// One employee card. Interactions are wired via data attributes handled by
// the delegated listeners in main.js:
//   data-action="card-click"     — navigate to this employee (non-active only)
//   data-action="select-direct"/"select-matrix" — expand reports of that kind
//   data-tip="info"/"grade"      — hover tooltips (overlays.js)
//   data-ctx                     — right-click context menu
const employeeCardHTML = (employee, { isActive = false, isMatrixNode = false, viewMode = 'direct' } = {}) => {
    const insights = employee._insights || { genderCount: { male: 0, female: 0, other: 0 } };
    const isIndividualContributor = insights.directCount === 0 && (insights.eaCount || 0) === 0 && insights.matrixCount === 0;

    const nameTint = employee._nameStatus ? NAME_STATUS_TINT[employee._nameStatus] : null;
    const statusStyle = STATUS_STYLES[employee.currentStatus];
    const baseBg = nameTint ? nameTint.card : 'bg-white';
    let cardClasses = `relative w-64 min-w-[16rem] mx-auto ${baseBg} rounded-xl shadow-md border p-4 transition-all duration-brand-base flex flex-col group `;
    if (isActive) cardClasses += "border-blue-500 ring-4 ring-blue-100 shadow-xl scale-105 cursor-default z-10";
    else if (isMatrixNode) cardClasses += "border-purple-300 border-dashed hover:border-purple-500 hover:shadow-lg cursor-pointer";
    else if (!nameTint) cardClasses += "border-slate-200 hover:border-blue-400 hover:shadow-lg cursor-pointer";
    else cardClasses += " hover:shadow-lg cursor-pointer";
    const cardStyle = statusStyle ? ` style="border-left: 4px solid ${statusStyle.rule}"` : '';
    const clickAttrs = !isActive ? ` data-action="card-click" data-id="${esc(employee._id)}"` : '';

    const mcRibbon = employee._isMgmtCommittee
        ? `<div class="absolute top-0 right-10 z-20 drop-shadow-sm" title="Management Committee">` +
          `<div class="bg-amber-100 text-amber-700 text-xs font-bold px-2 pt-1 pb-2.5 border-x border-b border-amber-200" style="clip-path: polygon(0 0, 100% 0, 100% 100%, 50% 80%, 0 100%)">MC</div></div>`
        : '';

    const chips = (employee.currentStatus || employee._nameStatus)
        ? `<div class="flex flex-wrap items-center gap-1.5 mb-3">` +
          (employee.currentStatus ? statusChipHTML(employee.currentStatus) : '') +
          (employee._nameStatus ? nameStatusChipHTML(employee._nameStatus) : '') + `</div>`
        : '';

    let metaRows = '';
    if (employee.function1 || employee.level) {
        metaRows += `<div class="flex items-center justify-between">` +
            (employee.function1
                ? `<div class="flex items-center space-x-1 truncate pr-2">${icon('Building2', { size: 12, cls: 'flex-shrink-0' })} <span class="truncate">${esc(employee.function1)}</span></div>`
                : `<span></span>`) +
            (employee.level
                ? `<div class="flex items-center space-x-1 text-slate-500 font-bold whitespace-nowrap" title="${esc(employee.level)}">${icon('Award', { size: 12, cls: 'flex-shrink-0' })} <span>${esc(employee.level)}</span></div>`
                : '') +
            `</div>`;
    }
    if (employee.location) {
        metaRows += `<div class="flex items-center space-x-1 truncate">${icon('MapPin', { size: 12, cls: 'flex-shrink-0' })} <span class="truncate">${esc(employee.location)}</span></div>`;
    }

    const id = esc(employee._id);
    const directPillCls = isActive && viewMode === 'direct' ? 'bg-blue-100 text-blue-800 ring-1 ring-blue-300' : 'hover:bg-blue-50 text-slate-600';
    const matrixPillCls = isActive && viewMode === 'matrix' ? 'bg-purple-100 text-purple-700 ring-1 ring-purple-300' : 'hover:bg-purple-50 text-purple-600';

    let footer;
    if (isIndividualContributor) {
        footer = `<div class="mt-3 flex justify-center items-center text-[10px] font-semibold pt-2 border-t text-slate-400 italic">Individual Contributor</div>`;
    } else if ((insights.directCount === 0 && (insights.eaCount || 0) === 0) && insights.matrixCount > 0) {
        footer = `<div class="mt-3 flex justify-end items-center text-[10px] font-semibold pt-2 border-t text-slate-600">` +
            `<div data-action="select-matrix" data-id="${id}" data-tip="grade" data-type="matrix" class="flex items-center px-1 py-0.5 rounded transition-colors cursor-pointer ${matrixPillCls}">` +
            `<span>${formatNum(insights.matrixCount)} Matrix</span></div></div>`;
    } else {
        footer = `<div class="mt-3 flex justify-between items-center text-[10px] font-semibold pt-2 border-t text-slate-600">` +
            `<div data-action="select-direct" data-id="${id}" data-tip="grade" data-type="direct" class="flex items-center px-1 py-0.5 rounded transition-colors cursor-pointer ${directPillCls}">` +
            icon('User', { size: 12, cls: `mr-1 ${isActive && viewMode === 'direct' ? 'text-blue-600' : 'text-blue-500'}` }) + ` ` +
            `${insights.eaCount > 0 ? `${insights.directCount} + EA` : `${formatNum(insights.directCount)} Direct`}</div>` +
            (insights.matrixCount > 0
                ? `<div data-action="select-matrix" data-id="${id}" data-tip="grade" data-type="matrix" class="flex items-center px-1 py-0.5 rounded transition-colors cursor-pointer ${matrixPillCls}">` +
                  `<span>${formatNum(insights.matrixCount)} Matrix</span></div>`
                : '') +
            `<div data-tip="grade" data-type="team" data-id="${id}" class="flex items-center cursor-help px-1 py-0.5 hover:bg-slate-100 rounded">` +
            icon('Users', { size: 12, cls: 'mr-1 text-slate-500' }) + ` ${formatNum(insights.totalTeam)} Team</div></div>`;
    }

    const avatarBg = isActive ? 'bg-blue-600' : isMatrixNode ? 'bg-purple-500' : 'bg-graphite-700';

    return `<div${isActive ? ' id="active-employee-card"' : ''} class="relative flex justify-center w-full ${isActive ? 'z-10' : 'z-0'}">` +
        `<div class="${cardClasses}"${cardStyle}${clickAttrs} data-ctx="${id}">` +
        mcRibbon +
        `<div class="absolute top-3 right-3 text-slate-400 hover:text-blue-600 z-20 cursor-help" data-tip="info" data-id="${id}" data-action="noop">${icon('Info', { size: 18 })}</div>` +
        `<div class="flex items-center space-x-3 mb-3 pr-6">` +
        avatarHTML(employee, 48, { bgClass: avatarBg }) +
        `<div class="flex-1 min-w-0">` +
        `<h3 class="font-bold text-graphite-900 truncate text-sm" title="${esc(employee.name)}">${esc(employee._formattedName) || (nameTint ? nameTint.label : '')}</h3>` +
        `<p class="text-xs text-graphite-500 line-clamp-2 mt-0.5 leading-snug min-h-[2.25em]" title="${esc(employee.jobTitle)}">${esc(employee.jobTitle || '')}</p>` +
        `</div></div>` +
        chips +
        `<div class="text-xs text-slate-600 bg-slate-50 p-2 rounded-md flex flex-col gap-1.5">${metaRows}</div>` +
        footer +
        `</div></div>`;
};

// Renders one expanded reports section (Direct or Matrix) under the active
// employee. Self-hides when the employee has zero reports of that kind, so
// Direct and Matrix can be stacked and each only appears when present.
const reportsSectionHTML = (reports, mode) => {
    const activeEmployee = derived.activeEmployee;
    const isMatrix = mode === 'matrix';
    const totalUnfilteredReports = isMatrix ? (activeEmployee?._matrix || []).length : (activeEmployee?._directs || []).length;
    if (totalUnfilteredReports === 0) return '';

    const hasFilteredReports = state.filterConditions.length > 0 && totalUnfilteredReports > reports.length;
    const isCompletelyFiltered = totalUnfilteredReports > 0 && reports.length === 0;

    let pillClasses = `text-[10px] font-bold uppercase tracking-wider px-4 py-1.5 rounded-full shadow-sm border flex items-center gap-2 `;
    if (hasFilteredReports || isCompletelyFiltered) {
        pillClasses += `bg-slate-100 text-slate-500 border-slate-200`;
    } else {
        pillClasses += isMatrix ? 'bg-purple-50 text-purple-700 border-purple-200' : 'bg-white text-slate-600 border-slate-200';
    }

    const cards = reports.length > 0
        ? `<div class="flex justify-center flex-wrap gap-6 w-full px-4">` +
          reports.map(emp =>
              `<div class="flex flex-col items-center relative w-full sm:w-auto">` +
              employeeCardHTML(emp, { isMatrixNode: isMatrix }) + `</div>`
          ).join('') + `</div>`
        : `<div class="text-sm text-slate-400 italic bg-white px-6 py-4 rounded-xl border border-slate-200 shadow-sm mt-2">` +
          `All ${isMatrix ? 'matrix' : 'direct'} reports for this employee have been hidden by the current filters.</div>`;

    return `<div class="flex flex-col items-center animate-fade-in-up w-full mt-2">` +
        `<div class="h-6 w-px ${isMatrix ? 'bg-purple-400' : 'bg-slate-300'}"></div>` +
        `<div class="flex flex-col items-center gap-1.5 mb-6">` +
        `<div class="${pillClasses}">` +
        `<span>${isMatrix ? 'Matrix Reports' : 'Direct Reports'} (${reports.length}${(hasFilteredReports || isCompletelyFiltered) ? ` / ${totalUnfilteredReports}` : ''})</span>` +
        ((hasFilteredReports || isCompletelyFiltered)
            ? `<div class="w-px h-3 bg-slate-300"></div><span class="text-slate-400 flex items-center gap-1">${icon('Filter', { size: 10 })} Filters Applied</span>`
            : '') +
        `</div></div>` +
        cards + `</div>`;
};

// The org-chart tab content (manager → active card → reports sections).
const orgViewHTML = () => {
    const { manager, activeEmployee, directReports, matrixReports } = derived;
    return `<div id="org-scroll" class="w-full mx-auto flex-col items-center pb-32 p-4 sm:p-8 overflow-y-auto ${state.appTab === 'org' ? 'flex' : 'hidden'}">` +
        (manager
            ? `<div class="flex flex-col items-center animate-fade-in-down w-full">` +
              employeeCardHTML(manager) +
              `<div class="h-10 w-px bg-slate-300 my-2"></div></div>`
            : '') +
        (activeEmployee
            ? `<div class="relative flex justify-center items-center my-4 animate-scale-in z-10 w-full max-w-sm">` +
              employeeCardHTML(activeEmployee, { isActive: true, viewMode: state.viewMode }) + `</div>`
            : '') +
        reportsSectionHTML(directReports, 'direct') +
        reportsSectionHTML(matrixReports, 'matrix') +
        `</div>`;
};

Object.assign(OS, { employeeCardHTML, orgViewHTML });
})();
