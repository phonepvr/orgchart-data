// Tabular view with sortable headers.
import { esc } from '../util.js';
import { NAME_STATUS_TINT } from '../constants.js';
import { formatNum } from '../data.js';
import { icon } from '../icons.js';
import { statusChipHTML, nameStatusChipHTML } from './bits.js';
import { state, derived } from '../state.js';

const sortableHeaderHTML = ({ label, field, align = 'left', width = '' }) => {
    const config = state.sortConfigs.find(c => c.field === field);
    const sortIcon = !config
        ? `<div class="w-4 inline-block"></div>`
        : config.dir === 'asc'
            ? icon('ArrowUp', { size: 14, cls: 'inline ml-1 text-blue-600' })
            : icon('ArrowDown', { size: 14, cls: 'inline ml-1 text-blue-600' });
    return `<th class="px-4 py-3 font-semibold cursor-pointer bg-slate-50 hover:bg-slate-200 select-none transition-colors ${width} ${align === 'center' ? 'text-center' : 'text-left'}" data-action="sort" data-field="${field}">` +
        `<div class="flex items-center ${align === 'center' ? 'justify-center' : 'justify-start'}"> ${label} ${sortIcon} </div></th>`;
};

const rowHTML = (emp) => {
    const rowCls = emp._nameStatus === 'approved' ? 'bg-graphite-50 hover:bg-graphite-100'
        : emp._nameStatus === 'unapproved' ? 'bg-signal/5 hover:bg-signal/10'
        : 'bg-white hover:bg-blue-50/50';
    return `<tr id="table-row-${esc(emp._id)}" class="cursor-pointer transition-colors duration-brand-fast ${rowCls}" data-action="table-row" data-id="${esc(emp._id)}">` +
        `<td class="px-4 py-3"><div class="font-bold text-graphite-900 flex items-center gap-1.5">` +
        `<span class="truncate max-w-[200px]">${esc(emp._formattedName) || (emp._nameStatus ? NAME_STATUS_TINT[emp._nameStatus].label : '')}</span>` +
        (emp._isMgmtCommittee ? `<span class="text-[8px] bg-amber-100 text-amber-700 px-1 rounded uppercase font-bold flex-shrink-0">MC</span>` : '') +
        (emp._nameStatus ? nameStatusChipHTML(emp._nameStatus) : '') +
        `</div></td>` +
        `<td class="px-4 py-3">${emp.currentStatus ? statusChipHTML(emp.currentStatus) : '<span class="text-graphite-300">-</span>'}</td>` +
        `<td class="px-4 py-3">${emp.level ? `<span class="bg-graphite-100 text-graphite-700 px-1.5 py-0.5 rounded-brand text-[10px] font-bold border border-graphite-200">${esc(emp.level)}</span>` : '<span class="text-graphite-300">-</span>'}</td>` +
        `<td class="px-4 py-3 text-graphite-700"><div class="truncate max-w-[200px]" title="${esc(emp.jobTitle)}">${esc(emp.jobTitle || '')}</div></td>` +
        `<td class="px-4 py-3 text-graphite-600"><div class="truncate max-w-[150px]" title="${esc(emp.function1)}">${esc(emp.function1 || '')}</div></td>` +
        `<td class="px-4 py-3 text-graphite-600"><div class="truncate max-w-[150px]" title="${esc(emp.location)}">${esc(emp.location || '')}</div></td>` +
        `<td class="px-4 py-3 text-center font-medium text-blue-700">${formatNum(emp._insights?.directCount)}</td>` +
        `<td class="px-4 py-3 text-center font-medium text-purple-600">${formatNum(emp._insights?.matrixCount)}</td>` +
        `<td class="px-4 py-3 text-center font-medium text-orange-600">${formatNum(emp._insights?.totalTeam)}</td>` +
        `<td class="px-4 py-3 text-slate-600"><div class="truncate max-w-[150px]" title="${esc(emp._formattedManagerName)}">${esc(emp._formattedManagerName || '-')}</div></td>` +
        `</tr>`;
};

export const tableViewHTML = () => {
    const rows = derived.tabularSortedData;
    const body = rows.length === 0
        ? `<div class="p-10 text-center text-slate-500">No employees match your current filter conditions.</div>`
        : `<div id="table-scroll" class="flex-1 overflow-auto" style="scrollbar-width: thin">` +
          `<table class="w-full text-left text-sm">` +
          `<thead class="text-slate-600 border-b border-slate-200 sticky top-0 z-10 bg-slate-50 shadow-sm"><tr>` +
          sortableHeaderHTML({ label: 'Employee', field: 'Employee' }) +
          sortableHeaderHTML({ label: 'Status', field: 'Status' }) +
          sortableHeaderHTML({ label: 'Level', field: 'Level' }) +
          sortableHeaderHTML({ label: 'Position Text', field: 'JobTitle' }) +
          sortableHeaderHTML({ label: 'Function 1', field: 'Function1' }) +
          sortableHeaderHTML({ label: 'Location', field: 'Location' }) +
          sortableHeaderHTML({ label: 'DR', field: 'DRSize', align: 'center' }) +
          sortableHeaderHTML({ label: 'Mat', field: 'MatrixSize', align: 'center' }) +
          sortableHeaderHTML({ label: 'Team', field: 'TeamSize', align: 'center' }) +
          sortableHeaderHTML({ label: 'Line Manager', field: 'Manager' }) +
          `</tr></thead>` +
          `<tbody class="divide-y divide-slate-100">${rows.map(rowHTML).join('')}</tbody>` +
          `</table></div>`;

    return `<div class="bg-white m-4 md:m-8 rounded-xl shadow-sm border border-slate-200 flex-1 flex-col overflow-hidden min-h-0 animate-fade-in-up ${state.appTab === 'table' ? 'flex' : 'hidden'}">` +
        `<div class="p-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">` +
        `<h2 class="text-lg font-bold text-slate-800 flex items-center gap-2">Filtered Results <span class="text-xs font-medium text-slate-500 bg-white border border-slate-200 px-2 py-0.5 rounded-full">${rows.length} records</span></h2>` +
        `</div>` + body + `</div>`;
};
