(function () {
'use strict';
const OS = window.OrgSense = window.OrgSense || {};
const { esc, STATUS_STYLES, NAME_STATUS_TINT, buildInitials, icon } = OS;
// Small shared HTML builders: chips, avatar, brand marks, grade lists.

const statusChipHTML = (status, size = 'sm') => {
    if (!status) return '';
    const sty = STATUS_STYLES[status];
    const cls = sty ? sty.chip : 'bg-graphite-100 text-graphite-700 border-graphite-200';
    const sizeCls = size === 'xs' ? 'text-[9px] px-1.5 py-0.5' : 'text-[10px] px-2 py-0.5';
    return `<span class="inline-flex items-center gap-1 ${sizeCls} font-sans font-bold uppercase tracking-wider rounded-brand border ${cls}">` +
        `<span class="w-1.5 h-1.5 rounded-full" style="background-color: ${sty ? sty.rule : '#5F6B80'}"></span>${esc(status)}</span>`;
};

const nameStatusChipHTML = (nameStatus) => {
    if (!nameStatus) return '';
    const tint = NAME_STATUS_TINT[nameStatus];
    return `<span class="inline-flex items-center text-[9px] font-sans font-bold uppercase tracking-wider rounded-brand border px-1.5 py-0.5 ${nameStatus === 'approved' ? 'bg-graphite-100 text-graphite-700 border-graphite-300' : 'bg-signal/10 text-signal border-signal/40'}">${tint.label}</span>`;
};

// Avatar with photo + initials fallback. On image load error a captured
// `error` listener (main.js) swaps the content for the fallback below.
const avatarHTML = (employee, size = 48, { ringClass = '', textClass = 'text-white', bgClass = 'bg-graphite-700' } = {}) => {
    const initials = employee._initials || buildInitials(employee.name);
    const showImg = !!employee.photoUrl;
    const isPlaceholder = !showImg && !!employee._nameStatus;
    const containerClasses = showImg
        ? 'bg-graphite-100'
        : (isPlaceholder ? 'bg-graphite-100 text-graphite-500 border border-graphite-300' : `${bgClass} ${textClass}`);
    const titleAttr = isPlaceholder
        ? ` title="${employee._nameStatus === 'approved' ? 'Approved seat – open' : 'Unapproved seat'}"` : '';
    const fallback = (!showImg && employee._nameStatus)
        ? icon('Armchair', { size: Math.round(size * 0.5), strokeWidth: 1.5 })
        : `<span>${esc(initials)}</span>`;
    const content = showImg
        ? `<img src="${esc(employee.photoUrl)}" alt="${esc(employee.name)}" class="w-full h-full object-cover" referrerpolicy="no-referrer" crossorigin="anonymous">`
        : fallback;
    return `<div data-avatar data-id="${esc(employee._id)}" data-size="${size}" data-bg="${bgClass}" data-text="${textClass}"` +
        ` class="rounded-full flex-shrink-0 flex items-center justify-center font-bold shadow-sm overflow-hidden ${ringClass} ${containerClasses}"` +
        ` style="width: ${size}px; height: ${size}px"${titleAttr}>${content}</div>`;
};

// The fallback applied when an avatar photo fails to load (mirrors the React
// onError -> errored state path).
const avatarErrorFallback = (container, employeeMap) => {
    const emp = employeeMap[container.dataset.id];
    const size = Number(container.dataset.size) || 48;
    const nameStatus = emp ? emp._nameStatus : null;
    const base = 'rounded-full flex-shrink-0 flex items-center justify-center font-bold shadow-sm overflow-hidden ';
    if (nameStatus) {
        container.className = base + 'bg-graphite-100 text-graphite-500 border border-graphite-300';
        container.title = nameStatus === 'approved' ? 'Approved seat – open' : 'Unapproved seat';
        container.innerHTML = icon('Armchair', { size: Math.round(size * 0.5), strokeWidth: 1.5 });
    } else {
        container.className = base + `${container.dataset.bg} ${container.dataset.text}`;
        container.innerHTML = `<span>${esc(emp ? emp._initials : '?')}</span>`;
    }
};

// Sorted grade rows for tooltips (was renderGradesList).
const gradesListHTML = (gradesObj) => {
    if (!gradesObj) return '<div class="p-2 text-slate-500 italic">No data</div>';
    const entries = Object.entries(gradesObj);
    if (entries.length === 0) return '<div class="p-2 text-slate-500 italic">No data</div>';
    const sorted = entries.sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
    return `<div class="flex flex-col space-y-1">${sorted.map(([g, c]) =>
        `<div class="flex justify-between items-center bg-slate-50 px-2 py-1 rounded">` +
        `<span class="text-slate-700 font-medium">${esc(g)}</span>` +
        `<span class="text-blue-600 font-bold bg-blue-100 px-2 rounded-full text-xs">${c}</span></div>`
    ).join('')}</div>`;
};

// AM/NS INDIA logo — two stacked lines, Smart Red on every background
// (red-on-white / red-on-black are the brand's permitted colorways).
// The `variant` parameter is kept so existing call sites don't change.
const amnsMarkHTML = (size = 'md', variant = 'light') => {
    const sizes = {
        sm: { mark: 'text-xl', gap: 'mt-0' },
        md: { mark: 'text-3xl', gap: 'mt-0' },
        lg: { mark: 'text-5xl', gap: 'mt-1' },
        xl: { mark: 'text-6xl', gap: 'mt-1' },
    }[size];
    return `<div class="flex flex-col">` +
        `<div class="font-display font-bold leading-none tracking-tight text-red-500 ${sizes.mark}">` +
        `<span>AM</span><span class="px-0.5">/</span><span>NS</span></div>` +
        `<div class="${sizes.gap} font-display font-bold leading-none tracking-tight text-red-500 ${sizes.mark}">` +
        `INDIA</div></div>`;
};

// Section 06 of the brand sheet — forward-diagonal accent. Smart Red by
// default; pure white / strong black are the only other permitted fills.
// Never reverse or break the angle.
const brandStrokeHTML = (className = '', tone = 'red') => {
    const fill = tone === 'white' ? '#FFFFFF' : tone === 'black' ? '#000000' : '#E52726';
    return `<svg viewBox="0 0 120 40" preserveAspectRatio="none" class="${className}" aria-hidden="true">` +
        `<polygon fill="${fill}" points="20,0 50,0 30,40 0,40"/>` +
        `<polygon fill="${fill}" points="80,0 110,0 90,40 60,40"/></svg>`;
};

Object.assign(OS, { statusChipHTML, nameStatusChipHTML, avatarHTML, avatarErrorFallback, gradesListHTML, amnsMarkHTML, brandStrokeHTML });
})();
