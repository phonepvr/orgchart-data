(function () {
'use strict';
const OS = window.OrgSense = window.OrgSense || {};
const { esc, COMPARE_COLORS, formatNum, sortEmployees, icon, state } = OS;
// Compare view: color slots, individual context, organizational context.

const coloredGradeListHTML = (gradesObj, textClass) => {
    if (!gradesObj || Object.keys(gradesObj).length === 0) {
        return `<div class="text-[10px] text-slate-400 italic">None</div>`;
    }
    const sorted = Object.entries(gradesObj).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
    return `<div class="flex flex-col gap-y-1 text-[10px]">${sorted.map(([g, c]) =>
        `<div class="flex justify-between items-center border-b border-slate-100 pb-0.5">` +
        `<span class="text-slate-600 font-medium truncate pr-1">${esc(g)}</span>` +
        `<span class="font-bold ${textClass}">${c}</span></div>`
    ).join('')}</div>`;
};

const compareReporteeTileHTML = (employee, isMatrix) =>
    `<div class="p-3 border bg-white rounded-xl flex flex-col text-slate-800 break-inside-avoid w-full shadow-sm ${isMatrix ? 'border-2 border-dashed border-purple-300 bg-purple-50/50' : 'border-solid border-slate-200'}">` +
    `<div class="flex justify-between items-start gap-1 mb-1">` +
    `<div class="font-bold text-xs leading-tight truncate pr-1">${esc(employee._formattedName)}</div>` +
    (employee.level ? `<div class="text-[10px] font-bold px-1.5 py-0.5 rounded border whitespace-nowrap flex-shrink-0 ${isMatrix ? 'bg-white border-purple-200 text-purple-700' : 'bg-slate-50 border-slate-300 text-slate-700'}">${esc(employee.level)}</div>` : '') +
    `</div>` +
    `<div class="text-[11px] text-slate-500 truncate font-medium">${esc(employee.jobTitle || '')}</div></div>`;

const headerCardHTML = (emp, activeColorObj) => {
    const insights = emp._insights || {};
    const isIC = (insights.directCount || 0) === 0 && (insights.eaCount || 0) === 0 && (insights.matrixCount || 0) === 0;

    let footer;
    if (isIC) {
        footer = `<div class="mt-3 flex justify-center items-center text-[10px] font-semibold pt-2 border-t border-slate-100 text-slate-400 italic">Individual Contributor</div>`;
    } else if ((insights.directCount || 0) === 0 && (insights.eaCount || 0) === 0 && (insights.matrixCount || 0) > 0) {
        footer = `<div class="mt-3 flex justify-end items-center text-[10px] font-semibold pt-2 border-t border-slate-100 text-slate-600">` +
            `<div class="flex items-center px-1.5 py-0.5 rounded bg-purple-50 text-purple-700"><span>${formatNum(insights.matrixCount)} Matrix</span></div></div>`;
    } else {
        footer = `<div class="mt-3 flex justify-between items-center text-[10px] font-semibold pt-2 border-t border-slate-100 text-slate-600">` +
            `<div class="flex items-center px-1.5 py-0.5 rounded bg-blue-50 text-slate-700">` +
            icon('User', { size: 12, cls: 'mr-1 text-blue-500' }) + ` ` +
            `${insights.eaCount > 0 ? `${insights.directCount} + EA` : `${formatNum(insights.directCount)} Direct`}</div>` +
            (insights.matrixCount > 0
                ? `<div class="flex items-center px-1.5 py-0.5 rounded bg-purple-50 text-purple-700"><span>${formatNum(insights.matrixCount)} Matrix</span></div>`
                : '') +
            `<div class="flex items-center px-1.5 py-0.5 rounded bg-slate-100 text-slate-700">` +
            icon('Users', { size: 12, cls: 'mr-1 text-slate-500' }) + ` ${formatNum(insights.totalTeam)} Team</div></div>`;
    }

    return `<div class="w-[320px] bg-white rounded-xl shadow-md border-t-4 ${activeColorObj.border} border-x border-b border-slate-200 p-4 relative">` +
        `<div class="flex justify-between items-start gap-2 mb-1">` +
        `<div class="font-bold text-slate-800 text-lg leading-tight truncate">${esc(emp._formattedName)}</div>` +
        (emp.level ? `<span class="bg-slate-100 text-slate-600 px-2 py-0.5 rounded text-xs font-bold border border-slate-200 shadow-sm flex-shrink-0">${esc(emp.level)}</span>` : '') +
        `</div>` +
        `<div class="text-sm text-slate-600 font-medium mb-1.5 truncate">${esc(emp.jobTitle)}</div>` +
        (emp.location ? `<div class="text-[10px] text-slate-500 flex items-center gap-1 font-medium">${icon('MapPin', { size: 10 })}${esc(emp.location)}</div>` : '') +
        (emp._isMgmtCommittee ? `<span class="absolute -top-1.5 -right-1 bg-amber-100 text-amber-700 text-[9px] font-bold px-1.5 py-0.5 rounded shadow-sm border border-amber-200">MC</span>` : '') +
        footer + `</div>`;
};

const indContextCardHTML = (emp) => {
    const cell = (label, value) => `<div><span class="text-slate-400 block">${label}</span>${value}</div>`;
    return `<div class="w-[320px] bg-white p-5 rounded-xl border border-slate-200 shadow-sm flex flex-col">` +
        `<div class="grid grid-cols-2 gap-4 text-xs">` +
        cell('Total Tenure', `<span class="font-bold text-slate-700">${emp._tenureFormatted ? `<span title="Tenure">${esc(emp._tenureFormatted)}</span>` : '<span class="text-slate-400">-</span>'}</span>`) +
        cell('Time in Role', `<span class="font-bold text-slate-700">${esc(emp._timeInRoleFormatted || '-')}</span>`) +
        cell('Since Promoted', `<span class="font-bold ${emp._lastPromotionFormatted ? 'text-green-700' : 'text-slate-400'}">${esc(emp._lastPromotionFormatted || '-')}</span>`) +
        cell('With Manager', `<span class="font-bold ${emp._timeWithManagerFormatted ? 'text-indigo-700' : 'text-slate-400'}">${esc(emp._timeWithManagerFormatted || '-')}</span>`) +
        (emp._age != null ? cell('Age', `<span class="font-bold text-slate-700">${emp._age}</span>`) : '') +
        (emp.gender ? cell('Gender', `<span class="font-bold text-slate-700">${esc(emp.gender)}</span>`) : '') +
        `</div>` +
        ((emp.cohortTags && emp.cohortTags.length > 0)
            ? `<div class="mt-4 flex flex-wrap gap-1.5">${emp.cohortTags.map(t => `<span class="text-[10px] font-bold bg-blue-50 text-blue-700 border border-blue-100 rounded-full px-2 py-0.5">${esc(t)}</span>`).join('')}</div>`
            : '') +
        `</div>`;
};

const orgContextColumnHTML = (emp, employeeMap, ceoId) => {
    const pageDrs = (emp._directs || []).map(id => employeeMap[id]).filter(Boolean).sort((a, b) => sortEmployees(a, b, ceoId));
    const pageMatrix = (emp._matrix || []).map(id => employeeMap[id]).filter(Boolean).sort((a, b) => sortEmployees(a, b, ceoId));

    return `<div class="w-[320px] flex flex-col gap-5">` +
        `<div class="flex gap-4 w-full bg-white p-4 rounded-xl border border-slate-200 shadow-sm items-start">` +
        `<div class="flex-1">` +
        `<div class="text-[10px] font-bold text-blue-600 uppercase tracking-wider mb-2 border-b border-blue-100 pb-1">Direct Summary</div>` +
        coloredGradeListHTML(emp._insights?.directGrades, 'text-blue-700') + `</div>` +
        `<div class="flex-1">` +
        `<div class="text-[10px] font-bold text-purple-600 uppercase tracking-wider mb-2 border-b border-purple-100 pb-1">Matrix Summary</div>` +
        coloredGradeListHTML(emp._insights?.matrixGrades, 'text-purple-700') + `</div>` +
        `</div>` +
        (pageDrs.length > 0
            ? `<div class="flex flex-col gap-2.5 mt-2">` +
              `<div class="text-[11px] font-bold text-slate-500 uppercase tracking-wider border-b border-slate-200 pb-1 px-1">Direct Reports (${pageDrs.length})</div>` +
              pageDrs.map(dr => compareReporteeTileHTML(dr, false)).join('') + `</div>`
            : '') +
        (pageMatrix.length > 0
            ? `<div class="flex flex-col gap-2.5 mt-2">` +
              `<div class="text-[11px] font-bold text-purple-500 uppercase tracking-wider border-b border-purple-200 pb-1 px-1">Matrix Reports (${pageMatrix.length})</div>` +
              pageMatrix.map(mr => compareReporteeTileHTML(mr, true)).join('') + `</div>`
            : '') +
        ((pageDrs.length === 0 && pageMatrix.length === 0)
            ? `<div class="text-center text-sm text-slate-400 italic mt-4 bg-white p-4 rounded-xl border border-slate-200 shadow-sm">Individual Contributor</div>`
            : '') +
        `</div>`;
};

const compareViewHTML = () => {
    const { compareList, employeeMap, ceoId } = state;

    // Was a useEffect: fall back to the first populated color slot.
    if (compareList[state.compareActiveColor].length === 0) {
        const firstPopulated = COMPARE_COLORS.find(c => compareList[c.id].length > 0);
        if (firstPopulated) state.compareActiveColor = firstPopulated.id;
    }
    const activeColor = state.compareActiveColor;
    const emps = compareList[activeColor].map(id => employeeMap[id]).filter(Boolean);
    const activeColorObj = COMPARE_COLORS.find(c => c.id === activeColor);

    if (Object.values(compareList).every(arr => arr.length === 0)) {
        return `<div class="flex flex-col items-center justify-center h-full text-slate-500 space-y-4 w-full bg-white rounded-xl shadow-sm border border-slate-200">` +
            icon('Users', { size: 48, cls: 'text-slate-300' }) +
            `<p>No employees added to compare yet.</p>` +
            `<p class="text-sm italic">Right-click any employee card in the Org Chart to add them.</p></div>`;
    }

    const tabs = COMPARE_COLORS.map(c =>
        `<button data-action="compare-tab" data-color="${c.id}" class="w-6 h-6 rounded-md border-2 ${activeColor === c.id ? c.border : 'border-transparent'} ${c.bg} opacity-80 hover:opacity-100 relative transition-all">` +
        (compareList[c.id].length > 0 ? `<span class="absolute -top-1.5 -right-1 bg-white text-[9px] font-bold rounded-full w-4 h-4 flex items-center justify-center text-slate-800 shadow-md border border-slate-100">${compareList[c.id].length}</span>` : '') +
        `</button>`
    ).join('');

    const indAccordion =
        `<div class="flex items-center justify-between w-full bg-slate-200/50 border border-slate-300 px-4 py-2.5 rounded-lg cursor-pointer hover:bg-slate-200 transition-colors" data-action="compare-accordion" data-which="ind">` +
        `<span class="font-bold text-slate-700 text-xs uppercase tracking-wider">Individual Context</span>` +
        icon(state.compareIndOpen ? 'ChevronDown' : 'ChevronRight', { size: 16, cls: 'text-slate-500' }) + `</div>` +
        (state.compareIndOpen
            ? `<div class="flex gap-6 mt-4 mb-6">${emps.map(indContextCardHTML).join('')}</div>`
            : '');

    const orgAccordion =
        `<div class="flex items-center justify-between w-full bg-slate-200/50 border border-slate-300 px-4 py-2.5 rounded-lg cursor-pointer hover:bg-slate-200 transition-colors mb-4 mt-2" data-action="compare-accordion" data-which="org">` +
        `<span class="font-bold text-slate-700 text-xs uppercase tracking-wider">Organizational Context</span>` +
        icon(state.compareOrgOpen ? 'ChevronDown' : 'ChevronRight', { size: 16, cls: 'text-slate-500' }) + `</div>` +
        (state.compareOrgOpen
            ? `<div class="flex gap-6 pb-20">${emps.map(e => orgContextColumnHTML(e, employeeMap, ceoId)).join('')}</div>`
            : '');

    return `<div class="w-full h-full flex flex-col overflow-hidden print:hidden bg-white rounded-xl shadow-sm border border-slate-200 min-h-0">` +
        `<div class="flex justify-between items-center px-6 py-4 border-b border-slate-100 flex-shrink-0">` +
        `<div class="flex gap-2 bg-slate-50 p-1 rounded-lg border border-slate-200">${tabs}</div></div>` +
        `<div class="flex-1 overflow-auto w-full bg-slate-50" style="scrollbar-width: thin">` +
        `<div class="flex flex-col p-6 pt-0 w-max mx-auto min-h-full">` +
        `<div class="sticky top-0 z-30 flex gap-6 pb-4 pt-6 bg-slate-50 border-b border-slate-200/50 mb-4 shadow-[0_4px_6px_-1px_rgb(248,250,252)]">` +
        emps.map(e => headerCardHTML(e, activeColorObj)).join('') + `</div>` +
        indAccordion + orgAccordion +
        `</div></div></div>`;
};

Object.assign(OS, { compareViewHTML });
})();
