// App shell: lock screen, upload screen, header (tabs + search), data-verify
// banner, sidebar (filters + cohort summaries), active-filter pills.
import { esc } from '../util.js';
import { NAME_STATUS_TINT, NUMERIC_FIELDS } from '../constants.js';
import { formatNum } from '../data.js';
import { icon } from '../icons.js';
import { amnsMarkHTML, brandStrokeHTML, statusChipHTML } from './bits.js';
import { metricScaleHTML } from './spotlight.js';
import { state, derived } from '../state.js';

// --- Lock screen (busy/error handled by direct DOM updates in main.js) ---
export const lockScreenHTML = () =>
    `<div class="h-screen w-full flex bg-white text-graphite-900 overflow-hidden">` +
    `<div class="hidden md:flex md:w-1/2 lg:w-3/5 bg-graphite-900 text-white flex-col justify-between p-12 lg:p-16 relative overflow-hidden">` +
    brandStrokeHTML('absolute top-0 right-0 h-32 w-64 opacity-90') +
    `<div class="relative z-10">${amnsMarkHTML('md', 'reverse')}</div>` +
    `<div class="relative z-10 max-w-xl">` +
    `<p class="font-sans font-semibold text-red-400 text-[11px] uppercase tracking-[0.18em] mb-5">#SmarterSteelsBrighterFutures</p>` +
    `<h1 class="font-sans font-bold uppercase tracking-[0.08em] text-5xl lg:text-6xl leading-[1.05] mb-4">` +
    `Smarter Steels<span class="text-red-500">.</span><br>Brighter Futures<span class="text-red-500">.</span></h1>` +
    `<p class="font-sans italic text-graphite-300 text-2xl lg:text-3xl leading-snug">Reimagineering Bharat.</p></div>` +
    `<div class="relative z-10 flex items-end justify-between text-graphite-300">` +
    `<p class="font-sans text-xs leading-relaxed max-w-xs">Banaunga main, banega Bharat.<br>` +
    `<span class="text-graphite-400">JV between ArcelorMittal &amp; Nippon Steel — 9 MTPA across India.</span></p>` +
    `<span class="font-mono text-[10px] text-graphite-400 hidden lg:inline">Org Sense · v1.0</span></div>` +
    `<div aria-hidden="true" class="absolute -bottom-32 -right-32 w-96 h-96 rounded-full bg-red-600 opacity-20 blur-3xl"></div>` +
    `</div>` +
    `<div class="flex-1 flex flex-col items-center justify-center p-8 lg:p-16 bg-graphite-50">` +
    `<div class="md:hidden mb-10">${amnsMarkHTML('md')}</div>` +
    `<div class="w-full max-w-sm">` +
    `<p class="font-sans font-semibold text-red-600 text-[11px] uppercase tracking-[0.18em] mb-3">Restricted access</p>` +
    `<h2 class="font-display text-3xl text-graphite-900 mb-3 leading-tight">Sign in to Org Sense.</h2>` +
    `<p class="font-sans text-graphite-500 text-[15px] leading-relaxed mb-8">For authorised AM/NS personnel only. Enter the access password issued by the People &amp; Culture team to continue.</p>` +
    `<form id="lock-form" class="flex flex-col gap-3">` +
    `<input type="password" id="lock-pwd" autofocus autocomplete="off"` +
    ` class="w-full px-4 py-3 bg-white border border-graphite-200 rounded-brand focus:outline-none focus:ring-2 focus:ring-red-600/20 focus:border-red-600 font-mono tracking-wider text-graphite-900 transition-colors duration-brand-fast"` +
    ` placeholder="Access password">` +
    `<p id="lock-err" class="text-red-700 text-sm font-medium hidden" role="alert"></p>` +
    `<button type="submit" id="lock-submit" disabled` +
    ` class="w-full py-3 rounded-brand font-sans font-semibold text-white tracking-wide transition-colors duration-brand-fast inline-flex items-center justify-center gap-2 bg-graphite-300 cursor-not-allowed">` +
    `Unlock <span aria-hidden="true">→</span></button>` +
    `</form>` +
    `<div class="mt-10 pt-6 border-t border-graphite-200">` +
    `<p class="font-sans text-[11px] text-graphite-500 leading-relaxed">` +
    `<span class="font-semibold text-graphite-700">Privacy.</span> Your spreadsheet is parsed in this browser tab and lives in memory only — no upload, no localStorage, no analytics, no third-party requests. Refresh or close the tab and the data is gone; you&rsquo;ll need to re-upload to continue.</p>` +
    `</div></div></div></div>`;

// --- Upload screen ---
export const uploadScreenHTML = () => {
    const templateHref = './orglens_sample_template.xlsx';
    const { loading, warnings, error } = state;
    return `<div class="h-screen w-full bg-graphite-50 flex flex-col">` +
        `<header class="px-8 py-5 border-b border-graphite-100 bg-white flex items-center justify-between flex-shrink-0">` +
        amnsMarkHTML('sm') +
        `<div class="flex items-center gap-4">` +
        brandStrokeHTML('h-5 w-16 hidden md:inline-block') +
        `<span class="font-mono text-[10px] text-graphite-400 hidden md:inline">Org Sense · v1.0</span></div></header>` +
        `<main class="flex-1 overflow-y-auto flex items-center justify-center p-6">` +
        `<div class="w-full max-w-2xl">` +
        `<p class="font-sans font-semibold text-red-600 text-[11px] uppercase tracking-[0.18em] mb-3">Step 01 · Upload</p>` +
        `<h1 class="font-display font-bold text-4xl md:text-5xl text-graphite-900 leading-[1.1] mb-3">` +
        `Upload your <span class="text-red-500">organisation file.</span></h1>` +
        `<p class="font-sans text-graphite-500 text-[15px] leading-relaxed mb-8 max-w-xl">` +
        `Drop in the AM/NS sample template populated with your employee data, or any Excel file with the same headers. Parsing happens in this browser — nothing is uploaded, and a refresh clears everything.</p>` +
        `<div id="dropzone" class="bg-white p-10 border border-graphite-200 rounded-brand transition-colors duration-brand-fast">` +
        `<div class="flex flex-col items-start gap-5">` +
        `<div class="w-12 h-12 bg-red-50 text-red-600 rounded-brand flex items-center justify-center border border-red-100">${icon('Upload', { size: 22 })}</div>` +
        `<div>` +
        `<h2 class="font-display text-xl text-graphite-900 mb-1">Drag &amp; drop, or browse</h2>` +
        `<p class="font-sans text-sm text-graphite-500">Accepts .xlsx and .xls files.</p></div>` +
        `<input type="file" accept=".xlsx, .xls" class="hidden" id="file-upload"${loading ? ' disabled' : ''}>` +
        `<div class="flex flex-wrap items-center gap-3">` +
        `<label for="file-upload" class="px-5 py-2.5 text-white font-sans font-semibold rounded-brand transition-colors duration-brand-fast inline-flex items-center gap-2 text-sm ${loading ? 'bg-graphite-300 cursor-not-allowed' : 'bg-red-600 hover:bg-red-700 cursor-pointer'}">` +
        `${loading ? 'Processing…' : 'Select Excel file <span aria-hidden="true">→</span>'}</label>` +
        `<a href="${templateHref}" download="orglens_sample_template.xlsx" class="px-5 py-2.5 font-sans font-semibold rounded-brand transition-colors duration-brand-fast inline-flex items-center gap-2 text-sm bg-graphite-900 hover:bg-graphite-800 text-white">` +
        `Download sample template</a></div>` +
        `<div class="w-full pt-4 border-t border-graphite-100">` +
        `<p class="font-sans text-[11px] text-graphite-500 leading-relaxed">` +
        `<span class="font-mono uppercase tracking-wider text-graphite-700">Required:</span> ` +
        `<span class="font-mono">Employee&#39;s Position Code</span>, ` +
        `<span class="font-mono">Employee name</span>, ` +
        `<span class="font-mono">Line Manager&#39;s Position Code</span>. Every other column is optional and missing fields are simply hidden in the UI.</p></div>` +
        (warnings && warnings.length > 0 ? warnings.map(w => `<p class="text-ember text-xs font-sans">${esc(w)}</p>`).join('') : '') +
        (error ? `<p class="text-red-700 text-sm font-sans font-semibold">${esc(error)}</p>` : '') +
        `</div></div>` +
        `<p class="font-sans text-[11px] text-graphite-500 mt-6 leading-relaxed max-w-xl">` +
        `<span class="font-semibold text-graphite-700">Privacy.</span> Your spreadsheet is parsed in this browser tab and lives in memory only — no upload, no localStorage, no analytics, no third-party requests. Refresh or close the tab and the data is gone; you&rsquo;ll need to re-upload to continue.</p>` +
        `</div></main></div>`;
};

// --- Privacy chip (popup visibility toggled by hover delegation) ---
const privacyChipHTML = () =>
    `<div class="relative" data-hover="privacy">` +
    `<span class="inline-flex items-center gap-1.5 px-2 py-1 rounded-brand border border-graphite-200 bg-graphite-50 text-[10px] font-mono uppercase tracking-wider text-graphite-700 cursor-help">` +
    `<span class="w-1.5 h-1.5 rounded-full bg-leaf" aria-hidden="true"></span>Local only</span>` +
    `<div data-privacy-pop class="hidden absolute left-0 top-full mt-2 w-72 bg-white border border-graphite-200 rounded-brand shadow-lg p-3 z-50 text-[11px] font-sans text-graphite-700 leading-relaxed">` +
    `<p class="font-semibold text-graphite-900 mb-1.5">Your data stays in this tab.</p>` +
    `<ul class="space-y-1 list-disc pl-4">` +
    `<li>File parsed in your browser via SheetJS.</li>` +
    `<li>Lives in memory only &mdash; no storage, no cookies.</li>` +
    `<li>No analytics, telemetry, or third-party scripts.</li>` +
    `<li>Refresh or close the tab and you&rsquo;ll need to re-upload.</li>` +
    `</ul></div></div>`;

// --- Header (app bar with tabs + search) ---
export const searchResultsHTML = () => {
    if (!(state.isSearchOpen && state.searchQuery)) return '';
    const results = derived.filteredSearch;
    return `<div class="absolute top-full right-0 mt-2 w-80 bg-white rounded-brand shadow-xl border border-graphite-100 overflow-hidden z-50">` +
        (results.length > 0 ? results.map(emp => {
            const rowCls = emp._nameStatus === 'approved' ? 'bg-graphite-50 hover:bg-graphite-100'
                : emp._nameStatus === 'unapproved' ? 'bg-signal/5 hover:bg-signal/10' : 'hover:bg-graphite-50';
            return `<button data-action="search-select" data-id="${esc(emp._id)}" class="w-full text-left px-4 py-3 border-b border-graphite-100 last:border-0 flex items-start gap-2 transition-colors duration-brand-fast ${rowCls}">` +
                (emp._nameStatus ? `<span class="mt-1.5 w-2 h-2 rounded-full flex-shrink-0 ${emp._nameStatus === 'approved' ? 'bg-graphite-400' : 'bg-signal'}" aria-hidden="true"></span>` : '') +
                `<span class="flex flex-col min-w-0 flex-1">` +
                `<span class="font-sans font-semibold text-graphite-900 truncate">${esc(emp._formattedName) || (emp._nameStatus ? NAME_STATUS_TINT[emp._nameStatus].label : '')}</span>` +
                `<span class="text-xs text-graphite-500 truncate">${esc([emp.jobTitle, emp.function1 || emp.location].filter(Boolean).join(' • '))}</span>` +
                `</span>` +
                (emp.currentStatus ? `<span class="flex-shrink-0">${statusChipHTML(emp.currentStatus, 'xs')}</span>` : '') +
                `</button>`;
        }).join('') : `<div class="px-4 py-3 text-graphite-500 text-sm">No employees found.</div>`) +
        `</div>`;
};

export const headerInnerHTML = () => {
    const tabBtn = (tab, label, iconHtml = '') =>
        `<button data-action="tab" data-tab="${tab}" class="px-4 py-1.5 rounded-brand text-sm font-sans font-semibold transition-all flex items-center gap-1.5 ${state.appTab === tab ? 'bg-white text-red-600 shadow-sm' : 'text-graphite-500 hover:text-graphite-800'}">${iconHtml}${label}</button>`;

    const searchArea = (state.appTab === 'org' || state.appTab === 'table')
        ? `<div class="relative w-64 hidden md:block" id="search-wrapper">` +
          icon('Search', { size: 18, cls: 'absolute left-3 top-1/2 transform -translate-y-1/2 text-graphite-400' }) +
          `<input type="text" placeholder="Search employee..." data-input="search"` +
          ` class="w-full pl-10 pr-4 py-1.5 border border-graphite-200 rounded-brand focus:outline-none focus:ring-2 focus:ring-red-600/20 focus:border-red-600 bg-graphite-50 text-sm transition-colors duration-brand-fast"` +
          ` value="${esc(state.searchQuery)}">` +
          `<div id="search-results">${searchResultsHTML()}</div>` +
          `</div>` +
          (state.ceoId && state.filterConditions.length === 0
              ? `<button data-action="go-top" class="px-4 py-1.5 bg-graphite-900 hover:bg-graphite-800 text-white rounded-brand font-sans font-semibold transition-colors duration-brand-fast text-sm whitespace-nowrap">Go to Top</button>`
              : '')
        : '';

    return `<div class="flex items-center gap-3 w-1/3">${amnsMarkHTML('sm')}${privacyChipHTML()}</div>` +
        `<div class="flex bg-graphite-50 p-1 rounded-brand border border-graphite-100 w-fit mx-auto justify-center">` +
        tabBtn('org', 'Structure') + tabBtn('table', 'Table') + tabBtn('compare', 'Compare', icon('BarChart2', { size: 14 }) + ' ') +
        `</div>` +
        `<div class="flex items-center justify-end space-x-4 w-1/3">${searchArea}</div>`;
};

// --- Data verify banner ---
export const bannerHTML = () => {
    if (!state.showDataVerifyBanner || state.data.length === 0) return '';
    const rowCount = state.data.length;
    return `<div class="bg-ember/10 border-b border-ember/40 px-6 py-2 flex items-center justify-between gap-4 print:hidden">` +
        `<div class="flex items-center gap-2 text-ember text-[12px] font-sans leading-snug">` +
        icon('AlertTriangle', { size: 14, cls: 'flex-shrink-0' }) +
        `<span><span class="font-semibold">Verify your data before printing or sharing.</span> ` +
        `This chart reflects the spreadsheet you uploaded (${rowCount.toLocaleString()} rows). ` +
        `Cross-check the source for accuracy &mdash; the output is only as correct as the input.</span></div>` +
        `<button data-action="dismiss-banner" class="text-ember/70 hover:text-ember transition-colors flex-shrink-0" aria-label="Dismiss data verification reminder">${icon('X', { size: 14 })}</button>` +
        `</div>`;
};

// --- Sidebar (filters + cohort summaries + benchmark scales + heatmap) ---
const filterConditionRowHTML = (cond) => {
    const { availableFilterFields, allUniqueByField } = derived;
    const fieldOptions = availableFilterFields.map(f =>
        `<option value="${esc(f)}"${cond.field === f ? ' selected' : ''}>${esc(f)}</option>`).join('') +
        `<option value="DR Size"${cond.field === 'DR Size' ? ' selected' : ''}>Direct Reports</option>` +
        `<option value="Total Reportees"${cond.field === 'Total Reportees' ? ' selected' : ''}>Total Reportees</option>` +
        `<option value="Team Size"${cond.field === 'Team Size' ? ' selected' : ''}>Team Size</option>`;

    const isNumeric = NUMERIC_FIELDS.includes(cond.field);
    const operatorSelect = isNumeric
        ? `<select data-change="filter-operator" data-cond-id="${cond.id}" class="bg-white border border-slate-200 rounded px-1 text-xs font-medium text-slate-700 focus:outline-none h-6">` +
          ['=', '>', '<'].map(op => `<option value="${esc(op)}"${cond.operator === op ? ' selected' : ''}>${esc(op)}</option>`).join('') +
          `</select>`
        : '';

    let valueArea;
    if (isNumeric) {
        valueArea = `<input type="number" data-input="filter-value" data-cond-id="${cond.id}" class="w-full border border-slate-200 rounded px-2 text-xs font-medium text-slate-700 focus:outline-none focus:border-blue-400 h-6" placeholder="0" value="${esc(cond.value)}">`;
    } else {
        const selectedCount = Array.isArray(cond.value) ? cond.value.length : 0;
        const dropdown = state.openDropdown === cond.id
            ? `<div data-dropdown-list class="absolute top-full left-0 mt-1 w-full max-h-48 overflow-y-auto bg-white border border-slate-200 shadow-xl rounded-md z-50 p-1 flex flex-col" style="scrollbar-width: thin">` +
              (derived.allUniqueByField[cond.field] || []).map(item =>
                  `<label class="flex items-center gap-2 text-xs p-1.5 hover:bg-slate-50 rounded cursor-pointer border border-transparent transition-colors">` +
                  `<input type="checkbox" data-change="filter-check" data-cond-id="${cond.id}" data-val="${esc(item)}" class="rounded text-blue-600 focus:ring-blue-500 w-3 h-3 m-0"${Array.isArray(cond.value) && cond.value.includes(item) ? ' checked' : ''}>` +
                  `<span class="truncate text-slate-700" title="${esc(item)}">${esc(item)}</span></label>`
              ).join('') + `</div>`
            : '';
        valueArea = `<div class="relative">` +
            `<button data-action="filter-dropdown-toggle" data-cond-id="${cond.id}" class="w-full border border-slate-200 rounded px-2 text-xs font-medium bg-white text-left flex justify-between items-center focus:border-blue-400 h-6 truncate">` +
            `<span class="truncate text-slate-700">${selectedCount > 0 ? `${selectedCount} Selected` : 'Select...'}</span>${icon('ChevronDown', { size: 12, cls: 'text-slate-400 flex-shrink-0 ml-1' })}` +
            `</button>${dropdown}</div>`;
    }

    return `<div class="flex items-center gap-1.5 p-1 rounded hover:bg-slate-50 border border-transparent hover:border-slate-200 group transition-colors">` +
        `<select data-change="filter-field" data-cond-id="${cond.id}" class="bg-transparent text-xs font-bold text-slate-700 focus:outline-none cursor-pointer p-0 border-none w-[110px] flex-shrink-0 appearance-none">${fieldOptions}</select>` +
        operatorSelect +
        `<div class="flex-1 min-w-0 relative">${valueArea}</div>` +
        `<button data-action="filter-remove" data-cond-id="${cond.id}" class="text-slate-300 hover:text-red-500 p-1 rounded transition-colors opacity-0 group-hover:opacity-100 flex-shrink-0">${icon('X', { size: 14 })}</button>` +
        `</div>`;
};

const cohortSummariesHTML = () => {
    const { cohortMetrics, heatmapStats } = derived;
    const populated = Object.keys(cohortMetrics).filter(k => cohortMetrics[k] && cohortMetrics[k].count > 0);

    const tiles = populated.map(k => {
        const s = cohortMetrics[k];
        const isMC = k === 'Mgmt Committee';
        const bgClass = isMC ? 'bg-amber-50 border-amber-200 hover:border-amber-400' : 'bg-white border-slate-200 hover:border-slate-400';
        const titleClass = isMC ? 'text-amber-800' : 'text-slate-700';
        return `<button data-action="cohort-tile" data-key="${esc(k)}" class="rounded-xl p-3 shadow-sm transition-all text-left border group ${bgClass} ${state.activeCohortScale === k ? 'ring-2 ring-blue-500 ring-offset-1' : ''}">` +
            `<div class="flex justify-between items-end mb-2.5 border-b border-slate-200/50 pb-2">` +
            `<span class="font-bold text-sm flex items-center gap-1.5 ${titleClass}">${isMC ? '' : icon('Award', { size: 14 })}${esc(k)}</span>` +
            `<span class="text-xs text-slate-500 bg-white/60 px-1.5 py-0.5 rounded font-bold">${s.count}</span></div>` +
            `<div class="grid grid-cols-3 gap-1 text-center divide-x divide-slate-200/50">` +
            `<div title="Median"><div class="text-[9px] text-slate-500 mb-0.5 font-bold uppercase">Direct</div><div class="font-bold text-blue-600">${formatNum(s.drMedian)}</div></div>` +
            `<div title="${s.matrixHasZeros ? `Median for ${s.matrixNzCount} employees` : 'Median'}"><div class="text-[9px] text-slate-500 mb-0.5 font-bold uppercase">Matrix</div><div class="font-bold text-purple-600">${formatNum(s.matrixMedian)}${s.matrixHasZeros && s.matrixNzCount > 0 ? '*' : ''}</div></div>` +
            `<div title="Median"><div class="text-[9px] text-slate-500 mb-0.5 font-bold uppercase">Total</div><div class="font-bold text-orange-600">${formatNum(s.totalRepMedian)}</div></div>` +
            `</div></button>`;
    }).join('');

    const empty = populated.length === 0
        ? `<div class="text-xs text-slate-400 italic px-2">No cohorts available. Provide Management Board EID or Cohort Tags.</div>`
        : '';

    let scales = '';
    const k = state.activeCohortScale;
    if (k && cohortMetrics[k] && cohortMetrics[k].count > 0) {
        const cm = cohortMetrics[k];
        const heatmap = heatmapStats.length > 0
            ? `<div class="mt-6 pt-5 border-t border-slate-100">` +
              `<h3 class="text-xs font-bold text-slate-500 uppercase tracking-wider mb-4 flex items-center">${icon('BarChart2', { size: 14, cls: 'mr-1.5' })} Median Span of Control</h3>` +
              `<div class="flex flex-col gap-2">` +
              heatmapStats.map(hs => {
                  const maxVal = Math.max(...heatmapStats.map(d => d.medianDr));
                  const intensity = maxVal > 0 ? (hs.medianDr / maxVal) : 0;
                  let colorClass = 'bg-slate-50 border-slate-200 text-slate-700';
                  if (intensity > 0.7) colorClass = 'bg-blue-500 border-blue-600 text-white';
                  else if (intensity > 0.3) colorClass = 'bg-blue-100 border-blue-200 text-blue-900';
                  return `<div class="border rounded-lg px-3 py-2 text-xs font-semibold flex items-center justify-between shadow-sm ${colorClass}">` +
                      `<span class="truncate pr-2">${esc(hs.dept)}</span>` +
                      `<span class="bg-white/40 px-2 py-0.5 rounded shadow-sm text-sm">${formatNum(hs.medianDr)}</span></div>`;
              }).join('') +
              `</div></div>`
            : '';

        scales = `<div class="mt-6 pt-5 border-t border-slate-200 animate-fade-in-down">` +
            `<div class="flex justify-between items-center mb-4">` +
            `<h4 class="text-sm font-bold text-slate-700 flex items-center gap-2">` +
            `${esc(k)} Benchmark <span class="text-xs font-bold bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full">${cm.count}</span></h4>` +
            `<button data-action="cohort-scale-close" class="text-slate-400 hover:text-slate-600 bg-slate-50 border border-slate-100 shadow-sm p-1 rounded-md">${icon('X', { size: 16 })}</button>` +
            `</div>` +
            `<div class="flex flex-col gap-2 mb-6">` +
            metricScaleHTML({ label: 'Direct Reports', min: cm.drMin, max: cm.drMax, median: cm.drMedian, value: 0, hideCurrent: true }) +
            `<div class="relative">` +
            metricScaleHTML({ label: 'Matrix Reports', min: cm.matrixMin, max: cm.matrixMax, median: cm.matrixMedian, value: 0, hideCurrent: true }) +
            (cm.matrixHasZeros && cm.matrixNzCount > 0
                ? `<p class="text-[9px] text-slate-400 italic absolute -bottom-2">* ${cm.matrixNzCount} employees in this cohort have matrix reports</p>`
                : '') +
            `</div>` +
            metricScaleHTML({ label: 'Total Reportees', min: cm.totalRepMin, max: cm.totalRepMax, median: cm.totalRepMedian, value: 0, hideCurrent: true }) +
            metricScaleHTML({ label: 'Team Size', min: cm.teamMin, max: cm.teamMax, median: cm.teamMedian, value: 0, hideCurrent: true }) +
            `</div>` + heatmap + `</div>`;
    }

    return `<div class="p-5">` +
        `<h3 class="text-xs font-bold text-slate-400 uppercase tracking-wider mb-4">Cohort Summaries</h3>` +
        `<div class="flex flex-col gap-3">${tiles}${empty}</div>` +
        scales + `</div>`;
};

export const sidebarClass = () => {
    if (!(state.appTab === 'org' || state.appTab === 'table')) return 'hidden';
    return `${state.isSidebarOpen ? 'w-72 md:w-80' : 'w-12'} bg-white border-r border-slate-200 flex-shrink-0 flex flex-col relative transition-all duration-300 z-50 shadow-[2px_0_10px_rgba(0,0,0,0.05)] hidden sm:flex`;
};

export const sidebarInnerHTML = () => {
    if (!(state.appTab === 'org' || state.appTab === 'table')) return '';

    const filterSection =
        `<div class="p-5 border-b border-slate-100 bg-white filter-dropdown-wrapper">` +
        `<button data-action="toggle-filter-panel" class="flex justify-between items-center w-full text-left focus:outline-none">` +
        `<h3 class="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-2">` +
        `${icon('Filter', { size: 14 })} Filters ${state.filterConditions.length > 0 ? `(${state.filterConditions.length})` : ''}</h3>` +
        icon(state.showFilterPanel ? 'ChevronDown' : 'ChevronRight', { size: 14, cls: 'text-slate-400' }) +
        `</button>` +
        (state.showFilterPanel
            ? `<div class="mt-4 space-y-3 animate-fade-in-down">` +
              `<div class="flex bg-slate-100 rounded p-0.5 border border-slate-200">` +
              `<button data-action="match-mode" data-mode="and" class="flex-1 px-2 py-1 text-[10px] uppercase font-bold rounded transition-colors ${state.filterMatchMode === 'and' ? 'bg-white shadow-sm text-blue-600' : 'text-slate-500 hover:text-slate-700'}">Match All</button>` +
              `<button data-action="match-mode" data-mode="or" class="flex-1 px-2 py-1 text-[10px] uppercase font-bold rounded transition-colors ${state.filterMatchMode === 'or' ? 'bg-white shadow-sm text-purple-600' : 'text-slate-500 hover:text-slate-700'}">Match Any</button>` +
              `</div>` +
              `<div class="space-y-1.5">${state.filterConditions.map(filterConditionRowHTML).join('')}</div>` +
              `<button data-action="add-filter" class="w-full flex justify-center items-center text-xs font-semibold text-blue-600 hover:text-blue-700 bg-blue-50 hover:bg-blue-100 py-1.5 rounded transition-colors border border-blue-100">` +
              `${icon('Plus', { size: 14, cls: 'mr-1' })} Add Rule</button>` +
              `</div>`
            : '') +
        `</div>`;

    const body = state.isSidebarOpen
        ? `<div id="sidebar-scroll" class="h-full overflow-y-auto flex flex-col" style="scrollbar-width: thin">` +
          filterSection + cohortSummariesHTML() + `</div>`
        : `<div class="flex flex-col items-center h-full pt-10 text-slate-400 font-bold uppercase tracking-widest text-xs">` +
          `<span class="rotate-90 whitespace-nowrap mt-16">Dashboard</span></div>`;

    return `<button data-action="toggle-sidebar" class="absolute -right-3.5 top-6 bg-slate-50 border border-slate-300 shadow-md rounded-full p-1.5 z-[60] text-slate-500 hover:text-blue-600 focus:outline-none">` +
        icon(state.isSidebarOpen ? 'ChevronLeft' : 'ChevronRight', { size: 16 }) + `</button>` +
        `<div class="flex-1 overflow-hidden relative z-40">${body}</div>`;
};

// --- Active filter pills (sticky bar above the content) ---
export const pillsBarHTML = () => {
    if (!((state.appTab === 'org' || state.appTab === 'table') && state.filterConditions.length > 0)) return '';
    const pills = state.filterConditions.flatMap(cond => {
        const out = [];
        const displayField = cond.field === 'DR Size' ? 'Directs' : cond.field;
        if (NUMERIC_FIELDS.includes(cond.field)) {
            if (cond.value !== '' && cond.value !== null) {
                out.push({ condId: cond.id, type: 'single', display: `${displayField} ${cond.operator} ${cond.value}` });
            }
        } else if (Array.isArray(cond.value)) {
            cond.value.forEach(val => {
                out.push({ condId: cond.id, type: 'array', val, display: `${displayField}: ${val}` });
            });
        }
        return out;
    });
    return `<div class="w-full bg-slate-50/90 backdrop-blur-md border-b border-slate-200 px-4 py-2.5 sm:px-8 flex flex-wrap items-center gap-2 z-20 shadow-sm min-h-[44px]">` +
        `<span class="text-[11px] font-bold text-slate-500 uppercase tracking-wider mr-1">` +
        `${state.filterMatchMode === 'or' ? 'Matches any:' : 'Matches all:'}</span>` +
        pills.map(pill =>
            `<div class="flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-bold border shadow-sm transition-colors ${state.filterMatchMode === 'or' ? 'bg-purple-100 text-purple-800 border-purple-200 hover:bg-purple-200' : 'bg-blue-100 text-blue-800 border-blue-200 hover:bg-blue-200'}">` +
            `<span>${esc(pill.display)}</span>` +
            `<button data-action="pill-remove" data-cond-id="${pill.condId}" data-type="${pill.type}"${pill.type === 'array' ? ` data-val="${esc(pill.val)}"` : ''} class="opacity-50 hover:opacity-100 bg-white/50 rounded-full p-0.5">${icon('X', { size: 10 })}</button>` +
            `</div>`
        ).join('') +
        `<button data-action="clear-filters" class="text-[10px] font-bold text-slate-400 hover:text-red-600 uppercase tracking-wider ml-auto flex items-center gap-1 transition-colors">${icon('Trash2', { size: 12 })} Clear All</button>` +
        `</div>`;
};
