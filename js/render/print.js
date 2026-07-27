(function () {
'use strict';
const OS = window.OrgSense = window.OrgSense || {};
const { esc, STATUS_STYLES, NAME_STATUS_TINT, GRID_COLS, PRINT_SUBJECT_DEPTH, sideColumns, drColumns, planSubjectPages, collectPrintSubjects } = OS;
// Print layout: A4-landscape position maps (subject + continuation pages).

const printTileHTML = (employee, { isMatrix = false, isLineManager = false } = {}) => {
    const matrixCount = employee._insights?.matrixCount || 0;
    const directCount = employee._insights?.directCount || 0;
    const eaCount = employee._insights?.eaCount || 0;
    const hasAny = !isLineManager && (matrixCount > 0 || directCount > 0 || eaCount > 0);
    const widthClass = isLineManager ? 'w-[220px] max-w-[220px]' : 'w-[160px] max-w-[160px]';

    const statusStyle = STATUS_STYLES[employee.currentStatus];
    const nameTint = employee._nameStatus ? NAME_STATUS_TINT[employee._nameStatus] : null;
    const tintStyle = nameTint ? nameTint.printTile : 'bg-white border-graphite-300';
    const ruleColor = statusStyle ? statusStyle.rule : '#2E3647';
    const borderStyle = isMatrix ? 'border-2 border-dashed' : 'border';

    return `<div class="p-2 ${borderStyle} ${tintStyle} rounded-brand flex flex-col text-graphite-900 break-inside-avoid shadow-sm ${widthClass}" style="border-left: 4px solid ${ruleColor}">` +
        `<div class="flex justify-between items-start gap-1 mb-0.5">` +
        `<div class="font-display font-medium text-[11px] leading-tight truncate pr-1 min-w-0 flex-1">${esc(employee._formattedName) || (nameTint ? nameTint.label : '')}</div>` +
        (employee.level ? `<div class="text-[9px] font-mono font-semibold px-1 rounded-brand border border-graphite-300 whitespace-nowrap flex-shrink-0 bg-white">${esc(employee.level.split(' - ')[0])}</div>` : '') +
        `</div>` +
        `<div class="text-[9px] font-sans text-graphite-600 line-clamp-2 leading-snug" title="${esc(employee.jobTitle)}">${esc(employee.jobTitle || '')}</div>` +
        ((employee.function1 || employee.location)
            ? `<div class="text-[8px] font-sans text-graphite-500 truncate mt-0.5">${esc([employee.function1, employee.location].filter(Boolean).join(' · '))}</div>`
            : '') +
        ((employee.currentStatus || employee._nameStatus || employee._isMgmtCommittee)
            ? `<div class="flex flex-wrap items-center gap-1 mt-1">` +
              (employee.currentStatus ? `<span class="text-[8px] font-sans font-bold uppercase tracking-wider px-1 py-0.5 rounded-brand border" style="color: ${ruleColor}; border-color: ${ruleColor}; background-color: ${ruleColor}14">${esc(employee.currentStatus)}</span>` : '') +
              (employee._nameStatus ? `<span class="text-[8px] font-sans font-bold uppercase tracking-wider px-1 py-0.5 rounded-brand border border-graphite-400 bg-white text-graphite-700">${nameTint.label}</span>` : '') +
              (employee._isMgmtCommittee ? `<span class="text-[8px] font-sans font-bold uppercase tracking-wider px-1 py-0.5 rounded-brand border border-red-300 bg-red-50 text-red-700">MC</span>` : '') +
              `</div>`
            : '') +
        (hasAny
            ? `<div class="flex justify-between items-center text-[9px] font-sans font-bold mt-1.5 pt-1 border-t border-graphite-200">` +
              (matrixCount > 0 ? `<span class="text-graphite-600">Matrix: ${matrixCount}</span>` : `<span></span>`) +
              `<span class="text-graphite-600">${eaCount > 0 ? `${directCount} + EA` : `Direct: ${directCount}`}</span></div>`
            : '') +
        `</div>`;
};

const printGradeListHTML = (gradesObj) => {
    if (!gradesObj) return '';
    const entries = Object.entries(gradesObj);
    if (entries.length === 0) return '';
    const sorted = entries.sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
    return `<div class="flex flex-col gap-y-0.5 text-[10px] font-sans">${sorted.map(([g, c]) =>
        `<div class="flex justify-between items-center border-b border-graphite-100 pb-0.5">` +
        `<span class="text-graphite-600 font-medium truncate pr-1">${esc(g)}</span>` +
        `<span class="font-bold text-graphite-900 font-mono">${c}</span></div>`
    ).join('')}</div>`;
};

const printLegendHTML = () =>
    `<div class="flex flex-wrap gap-x-3 gap-y-1 items-center text-[8px] font-sans">` +
    Object.entries(STATUS_STYLES).map(([k, v]) =>
        `<span class="inline-flex items-center gap-1 text-graphite-700">` +
        `<span class="w-2 h-2 rounded-full" style="background-color: ${v.rule}"></span>` +
        `<span class="uppercase tracking-wider font-semibold">${k}</span></span>`
    ).join('') +
    `<span class="inline-flex items-center gap-1 text-graphite-700">` +
    `<span class="w-3 h-2 rounded-brand bg-[#EBEDF1] border border-[#B4BBC8]"></span>` +
    `<span class="uppercase tracking-wider font-semibold">Approved</span></span>` +
    `<span class="inline-flex items-center gap-1 text-graphite-700">` +
    `<span class="w-3 h-2 rounded-brand bg-[#E5EFF8] border border-[#5FA1D6]"></span>` +
    `<span class="uppercase tracking-wider font-semibold">Unapproved</span></span>` +
    `<span class="inline-flex items-center gap-1 text-red-700">` +
    `<span class="w-3 h-2 rounded-brand border border-red-300 bg-red-50"></span>` +
    `<span class="uppercase tracking-wider font-semibold">MC</span></span>` +
    `</div>`;

const printLayoutHTML = (rootId, employeeMap, ceoId) => {
    const rootEmp = employeeMap[rootId];
    if (!rootEmp) return '';

    const subjectsToPrint = collectPrintSubjects(rootEmp, PRINT_SUBJECT_DEPTH, employeeMap, ceoId);
    const pages = subjectsToPrint.flatMap(s => planSubjectPages(s, employeeMap, ceoId));

    const printedAt = new Date().toLocaleString('en-GB', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });

    const pagesHTML = pages.map((page, index) => {
        const emp = page.subject;
        const isContinuation = page.kind === 'continuation';
        const manager = emp._managerId ? employeeMap[emp._managerId] : null;
        const pageDrs = page.drs;
        const pageMatrix = page.matrix || [];

        const hasDrs = pageDrs.length > 0;
        const hasMatrix = pageMatrix.length > 0;

        let matrixWidthClass = "flex-1";
        let drWidthClass = "flex-1";
        if (hasMatrix && hasDrs) {
            const diff = pageDrs.length - pageMatrix.length;
            if (diff >= 4) { drWidthClass = "flex-[2]"; matrixWidthClass = "flex-1"; }
            else if (diff <= -4) { matrixWidthClass = "flex-[2]"; drWidthClass = "flex-1"; }
        }

        const centralStatus = STATUS_STYLES[emp.currentStatus];
        const centralTint = emp._nameStatus ? NAME_STATUS_TINT[emp._nameStatus] : null;
        const centralBg = centralTint ? centralTint.printTile : 'bg-white border-graphite-900';
        const centralRule = centralStatus ? centralStatus.rule : '#0E1219';

        const isLast = index === pages.length - 1;

        const subHeader = isContinuation
            ? `<div class="w-full px-8 pt-3 pb-2 border-b border-graphite-100 flex items-end justify-between gap-4 flex-shrink-0">` +
              `<div class="min-w-0">` +
              `<div class="font-sans uppercase tracking-[0.18em] text-[9px] text-red-600 font-semibold">Direct Reports · continued</div>` +
              `<div class="font-display text-lg text-graphite-900 leading-tight truncate">${esc([emp._formattedName || (centralTint ? centralTint.label : ''), emp.jobTitle].filter(Boolean).join(' — '))}</div>` +
              `<div class="font-sans text-[10px] text-graphite-500 mt-0.5">Showing reports ${page.drStart + 1}–${page.drStart + pageDrs.length} of ${page.drTotal}</div>` +
              `</div>` + printLegendHTML() + `</div>`
            : `<div class="w-full px-8 pt-5 pb-2 border-b border-graphite-100 flex items-end justify-between gap-4 flex-shrink-0">` +
              `<div class="min-w-0">` +
              `<div class="font-sans uppercase tracking-[0.18em] text-[9px] text-red-600 font-semibold">Position structure</div>` +
              `<div class="font-display text-2xl text-graphite-900 leading-tight truncate">${esc(emp._formattedName) || (centralTint ? centralTint.label : '')}</div>` +
              `<div class="font-sans text-[11px] text-graphite-500 mt-0.5">${esc([emp.jobTitle, emp.function1, emp.location].filter(Boolean).join(' · '))}</div>` +
              (page.drTotal > pageDrs.length
                  ? `<div class="font-mono text-[9px] text-graphite-400 mt-0.5">Direct Reports: showing first ${pageDrs.length} of ${page.drTotal} · continues over</div>`
                  : '') +
              `</div>` + printLegendHTML() + `</div>`;

        const canvas = isContinuation
            ? `<div class="flex-1 py-6 px-8 flex justify-center items-start">` +
              `<div class="grid grid-cols-4 gap-3 justify-center">` +
              pageDrs.map(d => `<div style="page-break-inside: avoid">${printTileHTML(d)}</div>`).join('') +
              `</div></div>`
            : `<div class="flex-1 py-8 px-8 flex justify-center items-start">` +
              `<div class="flex gap-8 w-full items-start max-w-7xl justify-center">` +
              (hasMatrix
                  ? `<div class="${matrixWidthClass} pt-16 flex flex-col items-center">` +
                    `<div class="text-[10px] font-mono font-semibold uppercase tracking-[0.18em] text-graphite-500 mb-4 border-b border-graphite-200 pb-1 w-full text-center max-w-[160px]">MATRIX REPORTS</div>` +
                    `<div class="grid gap-3 justify-center ${sideColumns(pageMatrix.length) === 2 ? 'grid-cols-2' : 'grid-cols-1'}">` +
                    pageMatrix.map(m => `<div style="page-break-inside: avoid">${printTileHTML(m, { isMatrix: true })}</div>`).join('') +
                    `</div></div>`
                  : '') +
              `<div class="w-[260px] flex-shrink-0 flex flex-col items-center">` +
              (manager
                  ? `<div class="text-[8px] font-mono font-semibold uppercase text-graphite-400 mb-1 tracking-[0.18em]">Line Manager</div>` +
                    printTileHTML(manager, { isLineManager: true }) +
                    `<div class="w-px h-6 bg-graphite-300 my-1"></div>`
                  : '') +
              `<div class="w-full ${centralBg} rounded-brand p-3 mb-6 shadow-md border" style="border-left: 5px solid ${centralRule}">` +
              `<div class="flex justify-between items-start gap-1 mb-1">` +
              `<div class="font-display font-medium text-lg leading-tight truncate text-graphite-900 min-w-0 flex-1">${esc(emp._formattedName) || (centralTint ? centralTint.label : '')}</div>` +
              (emp.level ? `<div class="text-[10px] font-mono font-semibold px-1.5 py-0.5 border border-graphite-400 rounded-brand whitespace-nowrap flex-shrink-0 bg-white">${esc(emp.level.split(' - ')[0])}</div>` : '') +
              `</div>` +
              `<div class="text-[11px] font-sans text-graphite-700 font-medium mb-1.5 truncate">${esc(emp.jobTitle)}</div>` +
              ((emp.function1 || emp.location)
                  ? `<div class="text-[9px] font-sans text-graphite-500 mb-2">${esc([emp.function1, emp.location].filter(Boolean).join(' · '))}</div>`
                  : '') +
              `<div class="flex flex-wrap gap-1.5 mb-2">` +
              (emp.currentStatus && centralStatus ? `<span class="text-[9px] font-sans font-bold uppercase tracking-wider px-1.5 py-0.5 rounded-brand border" style="color: ${centralStatus.rule}; border-color: ${centralStatus.rule}; background-color: ${centralStatus.rule}14">${esc(emp.currentStatus)}</span>` : '') +
              (centralTint ? `<span class="text-[9px] font-sans font-bold uppercase tracking-wider px-1.5 py-0.5 rounded-brand border border-graphite-400 bg-white text-graphite-700">${centralTint.label}</span>` : '') +
              (emp._isMgmtCommittee ? `<span class="text-[9px] font-sans font-bold uppercase tracking-wider px-1.5 py-0.5 rounded-brand border border-red-300 bg-red-50 text-red-700">Management Committee</span>` : '') +
              `</div>` +
              ((emp._insights?.matrixCount > 0 || emp._insights?.directCount > 0 || emp._insights?.eaCount > 0)
                  ? `<div class="flex justify-between mt-2 pt-2 border-t border-graphite-300 text-[10px] font-sans font-bold">` +
                    (emp._insights?.matrixCount > 0 ? `<span class="text-graphite-900">Matrix: ${emp._insights?.matrixCount}</span>` : `<span></span>`) +
                    `<span class="text-graphite-900">${emp._insights?.eaCount > 0 ? `${emp._insights.directCount} + EA` : `Direct: ${emp._insights?.directCount || 0}`}</span></div>`
                  : '') +
              `</div>` +
              `<div class="w-full flex justify-center gap-6 px-2">` +
              (hasMatrix
                  ? `<div class="flex-1">` +
                    `<div class="text-[9px] font-mono font-semibold uppercase tracking-[0.18em] border-b border-graphite-300 pb-0.5 mb-2 text-graphite-500">Matrix</div>` +
                    printGradeListHTML(emp._insights.matrixGrades) + `</div>`
                  : '') +
              (hasDrs
                  ? `<div class="flex-1">` +
                    `<div class="text-[9px] font-mono font-semibold uppercase tracking-[0.18em] border-b border-graphite-300 pb-0.5 mb-2 text-graphite-500">Direct</div>` +
                    printGradeListHTML(emp._insights.directGrades) + `</div>`
                  : '') +
              `</div>` +
              `</div>` +
              (hasDrs
                  ? `<div class="${drWidthClass} pt-16 flex flex-col items-center">` +
                    `<div class="text-[10px] font-mono font-semibold uppercase tracking-[0.18em] text-graphite-500 mb-4 border-b border-graphite-200 pb-1 w-full text-center max-w-[160px]">DIRECT REPORTS</div>` +
                    `<div class="grid gap-3 justify-center ${GRID_COLS[drColumns(pageDrs.length, hasMatrix)]}">` +
                    pageDrs.map(d => `<div style="page-break-inside: avoid">${printTileHTML(d)}</div>`).join('') +
                    `</div></div>`
                  : '') +
              `</div></div>`;

        return `<div class="print-page flex flex-col box-border bg-white" style="page-break-after: ${isLast ? 'auto' : 'always'}">` +
            `<div class="print-brand-stripe w-full" aria-hidden="true"></div>` +
            `<div class="w-full bg-red-500 text-white px-8 py-2.5 flex items-center justify-between flex-shrink-0">` +
            `<div class="font-display font-bold text-base leading-none">` +
            `<span>AM</span><span class="opacity-80 px-0.5">/</span><span>NS</span>` +
            `<span class="font-sans uppercase tracking-[0.18em] text-[10px] ml-3 opacity-90">Org Sense · Position Map</span></div>` +
            `<div class="font-mono text-[10px] opacity-90">Printed ${printedAt}</div></div>` +
            subHeader + canvas +
            `<div class="w-full px-8 pt-2 pb-1 border-t border-graphite-100 flex items-center justify-between flex-shrink-0 text-graphite-500 text-[9px] font-sans">` +
            `<span class="font-bold uppercase tracking-[0.12em] text-graphite-900">Smarter Steels<span class="text-red-500">.</span> Brighter Futures<span class="text-red-500">.</span> <span class="text-graphite-400 font-normal normal-case tracking-normal">· AM/NS Org Sense</span></span>` +
            `<span class="font-mono">Page ${index + 1} of ${pages.length}</span></div>` +
            `<div class="w-full px-8 pb-2 text-[8px] font-sans italic text-graphite-500 text-center flex-shrink-0">` +
            `Data accuracy reminder &mdash; reflects the spreadsheet uploaded into Org Sense. Verify with the source before relying on this chart.</div>` +
            `</div>`;
    }).join('');

    return `<div class="w-full bg-white print:bg-white text-graphite-900 p-0 m-0 font-sans">${pagesHTML}</div>`;
};

Object.assign(OS, { printLayoutHTML });
})();
