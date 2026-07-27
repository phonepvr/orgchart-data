(function () {
'use strict';
const OS = window.OrgSense = window.OrgSense || {};
const { esc, STATUS_STYLES, formatNum, icon, statusChipHTML, gradesListHTML } = OS;
// Spotlight (info tooltip) + benchmark scales. Shared by the hover tooltip
// (overlays) and the sidebar cohort scales.

const metricScaleHTML = ({ label, min, max, median, value, hideCurrent = false }) => {
    const safeMax = hideCurrent ? max : Math.max(max, value, 1);
    const safeMin = hideCurrent ? min : Math.min(min, value);
    const range = safeMax - safeMin;
    const getPos = (v) => range === 0 ? 50 : ((v - safeMin) / range) * 100;

    const isValMin = !hideCurrent && value === min;
    const isValMed = !hideCurrent && value === median;
    const isValMax = !hideCurrent && value === max;

    const baseCircle = "absolute top-1/2 rounded-full shadow-sm transform -translate-x-1/2 -translate-y-1/2";
    const blueHollow = `${baseCircle} h-3.5 w-3.5 border-[2px] border-blue-500 bg-white z-10 cursor-help`;
    const orangeHollow = `${baseCircle} h-4 w-4 border-[2px] border-orange-500 bg-white z-20 cursor-help`;
    const blueWithOrangeFill = `${baseCircle} h-4 w-4 border-[2px] border-blue-500 bg-orange-500 z-20 cursor-help`;

    let markers = '';
    if (!isValMin) markers += `<div class="${blueHollow}" style="left: ${getPos(min)}%" title="Min: ${min}"></div>`;
    if (!isValMed && median !== min && median !== max) markers += `<div class="${blueHollow}" style="left: ${getPos(median)}%" title="Median: ${median}"></div>`;
    if (!isValMax && max !== min) markers += `<div class="${blueHollow}" style="left: ${getPos(max)}%" title="Max: ${max}"></div>`;
    if (!hideCurrent && (isValMin || isValMed || isValMax)) {
        markers += `<div class="${blueWithOrangeFill}" style="left: ${getPos(value)}%" title="Current: ${value} (Overlaps with Benchmark)"></div>`;
    } else if (!hideCurrent) {
        markers += `<div class="${orangeHollow}" style="left: ${getPos(value)}%" title="Current: ${value}"></div>`;
    }

    let labels = `<span title="Min: ${min}" class="absolute transform -translate-x-1/2 cursor-help ${isValMin ? 'text-orange-600 z-10' : 'text-blue-500'}" style="left: ${getPos(min)}%">${min}</span>`;
    if (median !== min && median !== max) {
        labels += `<span title="Median: ${median}" class="absolute transform -translate-x-1/2 cursor-help ${isValMed ? 'text-orange-600 z-10' : 'text-blue-500'}" style="left: ${getPos(median)}%">${median}</span>`;
    }
    if (max !== min) {
        labels += `<span title="Max: ${max}" class="absolute transform -translate-x-1/2 cursor-help ${isValMax ? 'text-orange-600 z-10' : 'text-blue-500'}" style="left: ${getPos(max)}%">${max}</span>`;
    }
    if (!hideCurrent && !(isValMin || isValMed || isValMax)) {
        labels += `<span title="Current: ${value}" class="absolute text-orange-600 z-10 bg-white/90 px-0.5 rounded transform -translate-x-1/2 cursor-help" style="left: ${getPos(value)}%">${value}</span>`;
    }

    return `<div class="mb-5 mt-2">` +
        `<div class="mb-1.5 text-sm text-slate-700 font-semibold leading-none">${esc(label)}</div>` +
        `<div class="relative w-full h-4 mt-2 mb-1.5">` +
        `<div class="absolute left-3 right-3 top-1/2 transform -translate-y-1/2 h-1.5 bg-blue-100 rounded-full overflow-hidden">` +
        (!hideCurrent ? `<div class="absolute top-0 bottom-0 left-0 bg-orange-400" style="width: ${getPos(value)}%"></div>` : '') +
        `</div>` +
        `<div class="absolute left-3 right-3 top-0 bottom-0">${markers}</div>` +
        `</div>` +
        `<div class="relative w-full h-4 mt-1 text-xs font-bold">` +
        `<div class="absolute left-3 right-3 top-0 bottom-0">${labels}</div>` +
        `</div></div>`;
};

const benchmarkBoxHTML = ({ title, rightElement = '', borderColor = 'border-slate-200', titleColor = 'text-slate-500', bgClass = '' }, children) =>
    `<div class="relative border ${borderColor} rounded-xl p-4 pt-5 mb-6 mt-4 ${bgClass}">` +
    `<div class="absolute -top-2.5 left-3 bg-white px-2 text-xs font-bold ${titleColor} uppercase tracking-wider">${esc(title)}</div>` +
    (rightElement ? `<div class="absolute -top-3 right-3 bg-white px-1">${rightElement}</div>` : '') +
    children + `</div>`;

const tenureHTML = (employee) => employee._tenureFormatted
    ? `<span title="Tenure">${esc(employee._tenureFormatted)}</span>`
    : `<span class="text-slate-400">-</span>`;

// The full Spotlight tooltip body (was the showTooltip portal in EmployeeCard).
const spotlightTooltipHTML = (employee, ceoId, globalMetrics) => {
    const insights = employee._insights || { genderCount: { male: 0, female: 0, other: 0 } };
    const isIndividualContributor = insights.directCount === 0 && (insights.eaCount || 0) === 0 && insights.matrixCount === 0;
    const totalGender = insights.genderCount.male + insights.genderCount.female + insights.genderCount.other;
    const malePct = totalGender > 0 ? Math.round((insights.genderCount.male / totalGender) * 100) : 0;
    const femalePct = totalGender > 0 ? Math.round((insights.genderCount.female / totalGender) * 100) : 0;
    const isTopNode = employee._id === ceoId;

    const tile = (label, value) =>
        `<div class="bg-slate-50 p-2.5 rounded-lg border border-slate-100 flex flex-col items-center text-center">` +
        `<span class="text-slate-400 font-medium mb-1 text-xs">${label}</span>${value}</div>`;

    let indTiles = '';
    if (employee._tenureFormatted) indTiles += tile('Total Tenure', `<span class="font-bold text-slate-700 flex items-center">${icon('CalendarDays', { size: 14, cls: 'mr-1.5' })} ${tenureHTML(employee)}</span>`);
    if (employee._timeInRoleFormatted) indTiles += tile('Time in Role', `<span class="font-bold text-slate-700 flex items-center">${icon('Clock', { size: 14, cls: 'mr-1.5' })} ${esc(employee._timeInRoleFormatted)}</span>`);
    if (employee._lastPromotionFormatted) indTiles += tile('Since Promoted', `<span class="font-bold text-green-700 flex items-center">${icon('Clock', { size: 14, cls: 'mr-1.5' })} ${esc(employee._lastPromotionFormatted)}</span>`);
    if (employee._timeWithManagerFormatted) indTiles += tile('With Manager', `<span class="font-bold text-indigo-700 flex items-center">${icon('Users', { size: 14, cls: 'mr-1.5' })} ${esc(employee._timeWithManagerFormatted)}</span>`);
    if (employee._age != null) indTiles += tile('Age', `<span class="font-bold text-slate-700">${employee._age}</span>`);
    if (employee.gender) indTiles += tile('Gender', `<span class="font-bold text-slate-700">${esc(employee.gender)}</span>`);
    if (employee.currentStatus) {
        indTiles += `<div class="bg-slate-50 p-2.5 rounded-lg border border-slate-100 flex flex-col items-center text-center col-span-2 sm:col-span-1">` +
            `<span class="text-slate-400 font-medium mb-1 text-xs">Position Status</span>${statusChipHTML(employee.currentStatus)}</div>`;
    }

    let extraRows = '';
    if (employee.email || employee.hrManagerName || (employee.cohortTags && employee.cohortTags.length > 0) || employee.employeeClass || employee.functionPlant || employee.asset || employee.cluster) {
        extraRows = `<div class="mt-4 pt-4 border-t border-slate-100 space-y-1.5 text-xs">` +
            (employee.email ? `<div class="flex items-center gap-2">${icon('Mail', { size: 12, cls: 'text-slate-400 flex-shrink-0' })}<a href="mailto:${esc(employee.email)}" class="text-blue-600 hover:underline truncate">${esc(employee.email)}</a></div>` : '') +
            (employee.hrManagerName ? `<div class="flex items-center gap-2">${icon('User', { size: 12, cls: 'text-slate-400 flex-shrink-0' })}<span class="text-slate-500">HR Manager:</span><span class="font-semibold text-slate-700 truncate">${esc(employee.hrManagerName)}</span></div>` : '') +
            (employee.employeeClass ? `<div class="flex items-center gap-2"><span class="text-slate-500">Class:</span><span class="font-semibold text-slate-700">${esc(employee.employeeClass)}</span></div>` : '') +
            (employee.functionPlant ? `<div class="flex items-center gap-2"><span class="text-slate-500">Function/Plant:</span><span class="font-semibold text-slate-700 truncate">${esc(employee.functionPlant)}</span></div>` : '') +
            (employee.asset ? `<div class="flex items-center gap-2"><span class="text-slate-500">Asset:</span><span class="font-semibold text-slate-700 truncate">${esc(employee.asset)}</span></div>` : '') +
            (employee.cluster ? `<div class="flex items-center gap-2"><span class="text-slate-500">Cluster:</span><span class="font-semibold text-slate-700 truncate">${esc(employee.cluster)}</span></div>` : '') +
            ((employee.cohortTags && employee.cohortTags.length > 0)
                ? `<div class="flex flex-wrap gap-1.5 pt-1">${employee.cohortTags.map(t => `<span class="text-[10px] font-bold bg-blue-50 text-blue-700 border border-blue-100 rounded-full px-2 py-0.5">${esc(t)}</span>`).join('')}</div>` : '') +
            `</div>`;
    }

    let orgContext;
    if (isIndividualContributor) {
        orgContext = `<div class="pt-4 text-center border-t border-slate-100 mt-4">` +
            `<p class="text-base font-semibold text-slate-600">Individual Contributor</p>` +
            `<p class="text-sm text-slate-400 mt-1">No reports.</p></div>`;
    } else {
        let benches = '';
        if (!isTopNode && employee._isMgmtCommittee && globalMetrics.mgmtCommittee && globalMetrics.mgmtCommittee.count > 0) {
            benches += benchmarkBoxHTML(
                { title: 'MC Benchmark', borderColor: 'border-amber-200', titleColor: 'text-amber-600', bgClass: 'bg-amber-50/20' },
                metricScaleHTML({ label: 'Direct Reports', min: globalMetrics.mgmtCommittee.drMin, max: globalMetrics.mgmtCommittee.drMax, median: globalMetrics.mgmtCommittee.drMedian, value: insights.directCount }) +
                metricScaleHTML({ label: 'Total Team Size', min: globalMetrics.mgmtCommittee.teamMin, max: globalMetrics.mgmtCommittee.teamMax, median: globalMetrics.mgmtCommittee.teamMedian, value: insights.totalTeam })
            );
        }
        if (!isTopNode && !employee._isMgmtCommittee) {
            if (insights.peerMedianDirects !== undefined) {
                let share = '';
                if (insights.pctOfManagerTeam !== undefined && insights.managerValidDrCount > 0) {
                    const expected = 100 / insights.managerValidDrCount;
                    let shareColor = "text-slate-700";
                    if (insights.pctOfManagerTeam <= expected * 0.92) shareColor = "text-red-600";
                    else if (insights.pctOfManagerTeam >= expected * 1.18) shareColor = "text-green-600";
                    else shareColor = "text-blue-600";
                    share = `<div class="mt-4 p-3 rounded-lg border border-slate-200 bg-slate-50 flex justify-between items-center">` +
                        `<span class="text-sm font-bold text-slate-600">Share of Manager's Team</span>` +
                        `<span class="font-bold text-xl leading-tight ${shareColor}">${insights.pctOfManagerTeam}%</span></div>`;
                }
                benches += benchmarkBoxHTML({ title: 'Peer Benchmark' },
                    metricScaleHTML({ label: 'Direct Reports', min: insights.peerMinDirects, max: insights.peerMaxDirects, median: insights.peerMedianDirects, value: insights.directCount }) + share);
            }
            if (employee.level && globalMetrics.level && globalMetrics.level[employee.level]) {
                const lm = globalMetrics.level[employee.level];
                benches += benchmarkBoxHTML({
                    title: 'Level Benchmark',
                    rightElement: `<div class="flex items-center space-x-1.5 text-slate-500 font-bold bg-slate-100 px-2 py-0.5 rounded text-xs">${icon('Award', { size: 12, cls: 'flex-shrink-0' })} <span>${esc(employee.level)}</span></div>`,
                },
                    metricScaleHTML({ label: 'Direct Reports', min: lm.drMin, max: lm.drMax, median: lm.drMedian, value: insights.directCount }) +
                    metricScaleHTML({ label: 'Total Team Size', min: lm.teamMin, max: lm.teamMax, median: lm.teamMedian, value: insights.totalTeam }));
            }
        }
        let diversity = '';
        if (insights.directCount > 0 && totalGender > 0) {
            diversity = `<div class="mt-5 px-1">` +
                `<h4 class="text-xs font-bold text-slate-500 uppercase tracking-wider mb-3">Team Diversity (DR)</h4>` +
                `<div class="w-full bg-slate-200 h-2.5 rounded-full overflow-hidden flex mt-2 shadow-inner">` +
                (malePct > 0 ? `<div style="width: ${malePct}%" class="bg-blue-500 h-full"></div>` : '') +
                (femalePct > 0 ? `<div style="width: ${femalePct}%" class="bg-pink-500 h-full"></div>` : '') +
                `</div>` +
                `<div class="flex justify-between text-sm mt-2 text-slate-600 font-medium">` +
                `<span>Male: <span class="font-bold text-slate-800">${malePct}%</span></span>` +
                `<span>Female: <span class="font-bold text-slate-800">${femalePct}%</span></span>` +
                `</div></div>`;
        }
        orgContext = `<div class="mt-4">` +
            `<h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-3 pb-1 border-b border-slate-100">Organizational Context</h4>` +
            benches + diversity + `</div>`;
    }

    return `<div class="bg-slate-800 text-white px-5 py-4 border-b flex items-center flex-shrink-0">` +
        `${icon('Info', { size: 18, cls: 'mr-2' })}<span class="font-bold text-base">Spotlight</span></div>` +
        `<div class="p-5 space-y-6 overflow-y-auto flex-1" style="scrollbar-width: thin">` +
        `<div>` +
        `<h4 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-3 pb-1 border-b border-slate-100">Individual Context</h4>` +
        `<div class="grid grid-cols-2 gap-3 text-sm mb-4">${indTiles}</div>` +
        extraRows +
        `</div>` +
        orgContext +
        `</div>`;
};

// Grade tooltip body (was the gradeTooltip portal in EmployeeCard).
const gradeTooltipHTML = (employee, type) => {
    const insights = employee._insights || {};
    let popupHeaderClass = "px-3 py-2 border-b text-xs font-bold uppercase tracking-wider flex justify-between ";
    if (type === 'direct') popupHeaderClass += "bg-blue-100 text-blue-800 border-blue-200";
    else if (type === 'matrix') popupHeaderClass += "bg-purple-100 text-purple-800 border-purple-200";
    else popupHeaderClass += "bg-slate-100 text-slate-700 border-slate-200";
    const title = type === 'direct' ? 'DR Summary' : type === 'matrix' ? 'Matrix Summary' : 'Team Summary';
    const grades = type === 'direct' ? insights.directGrades : type === 'matrix' ? insights.matrixGrades : insights.teamGrades;
    return `<div class="${popupHeaderClass}"><span>${title}</span></div>` +
        `<div class="p-2 max-h-64 overflow-y-auto" style="scrollbar-width: thin">${gradesListHTML(grades)}</div>`;
};

Object.assign(OS, { metricScaleHTML, benchmarkBoxHTML, spotlightTooltipHTML, gradeTooltipHTML });
})();
