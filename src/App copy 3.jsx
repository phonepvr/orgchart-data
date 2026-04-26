import React, { useState, useMemo, useEffect, useRef } from 'react';
import { createPortal } from 'react-dom';
import { Upload, Search, Info, Users, User, MapPin, Building2, Clock, CalendarDays, Award, ChevronDown, ChevronRight, ChevronLeft, X, Filter, Plus, Trash2, ArrowUp, ArrowDown, BarChart2, Printer } from 'lucide-react';

// --- Format Helpers ---
const formatNum = (num) => (num === 0 || num === '0' || !num) ? '-' : num;

const formatJobTitle = (title) => {
    if (!title) return '';
    let t = String(title)
        .replace(/Ã¢Â€Â[”“-]/g, '-') 
        .replace(/Ã¢[^\w\s]*/g, '-') 
        .replace(/â€“|â€”/g, '-')
        .replace(/[\u2013\u2014]/g, '-')
        .replace(/\s*-\s*/g, ' - ')
        .replace(/\s+/g, ' ');

    return t
        .replace(/\bSenior Vice President\b/ig, 'SVP')
        .replace(/\bSenior VP\b/ig, 'SVP')
        .replace(/\bVice President\b/ig, 'VP')
        .replace(/\bGeneral Manager\b/ig, 'GM')
        .replace(/\bDeputy GM\b/ig, 'DGM')
        .replace(/\bSenior Manager\b/ig, 'Sr Manager')
        .replace(/Sr\.\s*Manager/ig, 'Sr Manager')
        .replace(/\bAssistant Manager\b/ig, 'Asst Manager')
        .trim();
};

const processGrade = (gradeStr) => {
    let grade = String(gradeStr || '').split('-')[0].trim() || 'Unspecified';
    const gradeUp = grade.toUpperCase();
    if (['CMT', 'MT', 'GET', 'OT'].includes(gradeUp) || gradeUp.includes('TRAINEE')) return 'Trainee (CMT / MT / GET / OT)';
    if (gradeUp === 'DEPUTY GENERAL MANAGER') return 'DGM';
    if (gradeUp === 'DEPUTY MANAGER') return 'DM';
    return grade;
};

const formatCardGrade = (gradeStr) => {
    let grade = String(gradeStr || '').split('-')[0].trim() || 'Unspecified';
    const gradeUp = grade.toUpperCase();
    if (gradeUp === 'DEPUTY GENERAL MANAGER') return 'DGM';
    if (gradeUp === 'DEPUTY MANAGER') return 'DM';
    return grade;
};

const gradeOrder = [
    'MD', 'ED', 'SVP', 'VP', 'GM', 'DGM', 
    'M3', 'M2', 'GMR', 'DM', 'M1', 
    'Trainee (CMT / MT / GET / OT)', 'Corporate', 'Executive'
];

const getRank = (grade) => {
    const idx = gradeOrder.findIndex(g => g.toLowerCase() === grade.toLowerCase());
    return idx === -1 ? 999 : idx; 
};

const getEmpSortRank = (emp, ceoId) => {
    if (!emp) return 999;
    if (emp._id === ceoId) return -1;
    const rank = getRank(emp._summaryGrade);
    if (rank === 0) return 1.5; 
    return rank;
};

const sortEmployees = (a, b, ceoId) => {
    const rankA = getEmpSortRank(a, ceoId);
    const rankB = getEmpSortRank(b, ceoId);
    if (rankA !== rankB) return rankA - rankB;
    
    const exA = a._isExCom ? 1 : 0;
    const exB = b._isExCom ? 1 : 0;
    if (exA !== exB) return exB - exA;
    
    const opA = a._isOpCom ? 1 : 0;
    const opB = b._isOpCom ? 1 : 0;
    if (opA !== opB) return opB - opA;
    
    return (b._insights?.totalTeam || 0) - (a._insights?.totalTeam || 0);
};

const getMedian = (arr) => {
    if (!arr || arr.length === 0) return 0;
    const s = [...arr].sort((a,b) => a - b);
    const mid = Math.floor(s.length / 2);
    return s.length % 2 !== 0 ? s[mid] : s[mid - 1];
};

const renderGradesList = (gradesObj) => {
    if (!gradesObj) return <div className="p-2 text-slate-500 italic">No data</div>;
    const entries = Object.entries(gradesObj);
    if(entries.length === 0) return <div className="p-2 text-slate-500 italic">No data</div>;
    const sorted = entries.sort((a, b) => {
        const rankA = getRank(a[0]);
        const rankB = getRank(b[0]);
        if (rankA !== rankB) return rankA - rankB;
        return b[1] - a[1]; 
    });
    return (
        <div className="flex flex-col space-y-1">
            {sorted.map(([g, c]) => (
                <div key={g} className="flex justify-between items-center bg-slate-50 px-2 py-1 rounded">
                    <span className="text-slate-700 font-medium">{g}</span>
                    <span className="text-blue-600 font-bold bg-blue-100 px-2 rounded-full text-xs">{c}</span>
                </div>
            ))}
        </div>
    );
};

const toProperCase = (str) => str ? str.replace(/\b\w+/g, txt => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase()) : '';

const formatDisplayFirstLast = (name) => {
    if (!name) return '';
    let cleanName = String(name).replace(/\./g, '').trim();
    const parts = cleanName.split(/\s+/);
    
    // If it's only two parts, keep both even if one is an initial or prefix
    if (parts.length === 2) {
        return toProperCase(parts.join(' '));
    }
    
    let startIdx = 0;
    const skipRegex = /^(?:[A-Za-z]|mohd|mohammad|mohamad|mohammed|mohamed|muhammad|muhamad|muhammed|muhamed|md)$/i;
    
    while (startIdx < parts.length - 1 && skipRegex.test(parts[startIdx])) {
        startIdx++;
    }
    const sigParts = parts.slice(startIdx);
    
    let display = sigParts.join(' ');
    if (sigParts.length >= 3) display = `${sigParts[0]} ${sigParts[sigParts.length - 1]}`;
    return toProperCase(display);
};

const parseExcelDate = (excelDate) => {
    if (excelDate === undefined || excelDate === null || excelDate === '') return null;
    if (typeof excelDate === 'number') return new Date(Math.round((excelDate - 25569) * 86400 * 1000));
    
    let dateStr = String(excelDate).trim();
    const parts = dateStr.split(/[-/]/);
    if (parts.length === 3) {
        let y, m, d;
        if (parts[2].length === 4) {
            d = parseInt(parts[0], 10);
            m = parseInt(parts[1], 10);
            y = parseInt(parts[2], 10);
        } else if (parts[0].length === 4) {
            y = parseInt(parts[0], 10);
            m = parseInt(parts[1], 10);
            d = parseInt(parts[2], 10);
        }
        if (y && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
            return new Date(y, m - 1, d);
        }
    }
    const fallbackDate = new Date(excelDate);
    return isNaN(fallbackDate.getTime()) ? null : fallbackDate;
};

const formatDateUI = (dateObj) => {
    if (!dateObj) return '-';
    return dateObj.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
};

const formatDuration = (start, end) => {
    if (!start) return '-';
    let s = start;
    if (s instanceof Date && s.getFullYear() >= 9999) s = new Date();
    let e = end || new Date();
    if (e instanceof Date && e.getFullYear() >= 9999) e = new Date();
    let months = (e - s) / (1000 * 60 * 60 * 24 * 30.4375);
    if (months <= 0) return '< 1 mo';
    if (months < 1) return '< 1 mo';
    if (months < 12) return Math.round(months) + ' mos';
    return (months / 12).toFixed(1) + ' yrs';
};

const isEA = (e) => {
    if (!e) return false;
    const title = String(e['Job Title'] || '').toLowerCase();
    return title.includes('executive assistant') || title.includes('executive secretary') || title.includes('confidential secretary');
};

const TenureDisplay = ({ employee }) => {
    const groupStr = employee._tenureFormatted;
    if (employee._isDiffTenure) {
        const orgStr = employee._orgTenureFormatted;
        let displayOrg = orgStr;
        
        if (orgStr.endsWith(' yrs') && groupStr.endsWith(' yrs')) {
            displayOrg = orgStr.replace(' yrs', '');
        } else if (orgStr.endsWith(' mos') && groupStr.endsWith(' mos')) {
            displayOrg = orgStr.replace(' mos', '');
        }
        
        return (
            <span className="flex items-center gap-0.5">
                <span title="Org Tenure">{displayOrg}</span>
                <span className="text-slate-400 px-0.5">/</span>
                <span title="Group Tenure">{groupStr}</span>
            </span>
        );
    }
    return <span title="Tenure">{employee._tenureFormatted}</span>;
};

// --- Shared Display Components ---
const MetricScale = ({ label, min, max, median, value, hideCurrent = false }) => {
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

    return (
        <div className="mb-5 mt-2">
            <div className="mb-1.5 text-sm text-slate-700 font-semibold leading-none">{label}</div>
            <div className="relative w-full h-4 mt-2 mb-1.5">
                <div className="absolute left-3 right-3 top-1/2 transform -translate-y-1/2 h-1.5 bg-blue-100 rounded-full overflow-hidden">
                    {!hideCurrent && <div className="absolute top-0 bottom-0 left-0 bg-orange-400" style={{ width: `${getPos(value)}%` }}></div>}
                </div>
                <div className="absolute left-3 right-3 top-0 bottom-0">
                    {!isValMin && <div className={blueHollow} style={{ left: `${getPos(min)}%` }} title={`Min: ${min}`}></div>}
                    {!isValMed && median !== min && median !== max && <div className={blueHollow} style={{ left: `${getPos(median)}%` }} title={`Median: ${median}`}></div>}
                    {!isValMax && max !== min && <div className={blueHollow} style={{ left: `${getPos(max)}%` }} title={`Max: ${max}`}></div>}
                    { !hideCurrent && (isValMin || isValMed || isValMax) ? (
                        <div className={blueWithOrangeFill} style={{ left: `${getPos(value)}%` }} title={`Current: ${value} (Overlaps with Benchmark)`}></div>
                    ) : (!hideCurrent && (
                        <div className={orangeHollow} style={{ left: `${getPos(value)}%` }} title={`Current: ${value}`}></div>
                    ))}
                </div>
            </div>
            <div className="relative w-full h-4 mt-1 text-xs font-bold">
                <div className="absolute left-3 right-3 top-0 bottom-0">
                    <span title={`Min: ${min}`} className={`absolute transform -translate-x-1/2 cursor-help ${isValMin ? 'text-orange-600 z-10' : 'text-blue-500'}`} style={{ left: `${getPos(min)}%` }}>{min}</span>
                    {median !== min && median !== max && (
                        <span title={`Median: ${median}`} className={`absolute transform -translate-x-1/2 cursor-help ${isValMed ? 'text-orange-600 z-10' : 'text-blue-500'}`} style={{ left: `${getPos(median)}%` }}>{median}</span>
                    )}
                    {max !== min && (
                        <span title={`Max: ${max}`} className={`absolute transform -translate-x-1/2 cursor-help ${isValMax ? 'text-orange-600 z-10' : 'text-blue-500'}`} style={{ left: `${getPos(max)}%` }}>{max}</span>
                    )}
                    {!hideCurrent && !(isValMin || isValMed || isValMax) && (
                        <span title={`Current: ${value}`} className="absolute text-orange-600 z-10 bg-white/90 px-0.5 rounded transform -translate-x-1/2 cursor-help" style={{ left: `${getPos(value)}%` }}>{value}</span>
                    )}
                </div>
            </div>
        </div>
    );
};

const BenchmarkBox = ({ title, rightElement, borderColor = 'border-slate-200', titleColor = 'text-slate-500', children, bgClass = '' }) => (
    <div className={`relative border ${borderColor} rounded-xl p-4 pt-5 mb-6 mt-4 ${bgClass}`}>
        <div className={`absolute -top-2.5 left-3 bg-white px-2 text-xs font-bold ${titleColor} uppercase tracking-wider`}>{title}</div>
        {rightElement && (<div className="absolute -top-3 right-3 bg-white px-1">{rightElement}</div>)}
        {children}
    </div>
);

const SortableHeader = ({ label, field, align = 'left', width = '', sortConfigs, handleSort }) => {
    const renderSortIcon = () => {
        const config = sortConfigs.find(c => c.field === field);
        if (!config) return <div className="w-4 inline-block"></div>;
        return config.dir === 'asc' ? <ArrowUp size={14} className="inline ml-1 text-blue-600"/> : <ArrowDown size={14} className="inline ml-1 text-blue-600"/>;
    };
    return (
        <th className={`px-4 py-3 font-semibold cursor-pointer bg-slate-50 hover:bg-slate-200 select-none transition-colors ${width} ${align === 'center' ? 'text-center' : 'text-left'}`} onClick={() => handleSort(field)}>
            <div className={`flex items-center ${align === 'center' ? 'justify-center' : 'justify-start'}`}> {label} {renderSortIcon()} </div>
        </th>
    );
};

const renderManagerRows = (nodes, depth = 0) => {
    if (!nodes || nodes.length === 0) return null;
    return nodes.map((node, i) => (
        <React.Fragment key={`${node.id || i}-${depth}`}>
            <tr className="border-b border-slate-100 hover:bg-slate-50 transition-colors">
                <td className="py-2 pl-6">
                    {node.children && node.children.length > 0 && (
                        <button className="p-1 rounded hover:bg-slate-200 text-slate-500 transition-colors">
                            <ChevronRight size={14} />
                        </button>
                    )}
                </td>
                <td className="px-6 py-2 font-medium text-slate-800" style={{ paddingLeft: `${depth * 1.5 + 1.5}rem` }}>{node.name || 'Unknown'}</td>
                <td className="px-6 py-2 text-slate-600 text-xs">{node.department || '-'}</td>
                <td className="px-6 py-2 text-center text-slate-700 font-bold">{node.teamSize || 0}</td>
                <td className="px-6 py-2 text-center text-orange-600 font-bold">{node.directAttrition || 0}</td>
                <td className="px-6 py-2 text-center text-red-600 font-bold">{node.totalAttrition || 0}</td>
            </tr>
            {renderManagerRows(node.children, depth + 1)}
        </React.Fragment>
    ));
};

// --- Print Layout Components ---
const PrintTile = ({ employee, isMatrix, isLineManager, targetLocation }) => {
    const showLocation = isLineManager && employee['Location Name'] && employee['Location Name'] !== targetLocation;
    const matrixCount = employee._insights?.matrixCount || 0;
    const directCount = employee._insights?.directCount || 0;
    const eaCount = employee._insights?.eaCount || 0;
    const hasAny = !isLineManager && (matrixCount > 0 || directCount > 0 || eaCount > 0);
    const widthClass = isLineManager ? 'w-[220px] max-w-[220px]' : 'w-[160px] max-w-[160px]';
    
    return (
        <div className={`p-2 border ${isMatrix ? 'border-2 border-dashed border-slate-400' : 'border border-solid border-slate-400'} bg-white rounded flex flex-col text-slate-800 break-inside-avoid shadow-sm ${widthClass}`}>
            <div className="flex justify-between items-start gap-1 mb-0.5">
                <div className="font-bold text-[11px] leading-tight truncate pr-1">{employee._formattedName}</div>
                <div className="text-[9px] font-bold px-1 rounded border border-slate-300 whitespace-nowrap flex-shrink-0 bg-slate-50">{employee._summaryGrade}</div>
            </div>
            
            <div className="text-[9px] text-slate-600 truncate">{employee['Job Title'] || 'No Title'}</div>
            
            {showLocation && (
                <div className="text-[8px] text-slate-500 mt-0.5">{employee['Location Name']}</div>
            )}
            
            {hasAny && (
                <div className="flex justify-between items-center text-[9px] font-bold mt-1.5 pt-1 border-t border-slate-200">
                    {matrixCount > 0 ? <span className="text-slate-600">Matrix: {matrixCount}</span> : <span></span>}
                    <span className="text-slate-600">{eaCount > 0 ? `${directCount} + EA` : `Direct: ${directCount}`}</span>
                </div>
            )}
        </div>
    );
};

const PrintGradeList = ({ gradesObj }) => {
    if (!gradesObj) return null;
    const entries = Object.entries(gradesObj);
    if(entries.length === 0) return null;
    const sorted = entries.sort((a, b) => {
        const rankA = getRank(a[0]);
        const rankB = getRank(b[0]);
        if (rankA !== rankB) return rankA - rankB;
        return b[1] - a[1]; 
    });
    return (
        <div className="flex flex-col gap-y-0.5 text-[10px]">
            {sorted.map(([g, c]) => (
                <div key={g} className="flex justify-between items-center border-b border-slate-100 pb-0.5">
                    <span className="text-slate-600 font-medium truncate pr-1">{g}</span>
                    <span className="font-bold text-slate-900">{c}</span>
                </div>
            ))}
        </div>
    );
};

const PrintLayout = ({ rootId, employeeMap, ceoId }) => {
    const rootEmp = employeeMap[rootId];
    if (!rootEmp) return null;

    const pages = [rootEmp];
    const rootDrs = (rootEmp._directs || []).map(id => employeeMap[id]).filter(Boolean).sort((a, b) => sortEmployees(a, b, ceoId));
    
    rootDrs.forEach(dr => {
        if ((dr._insights?.directCount || 0) > 0 || (dr._insights?.matrixCount || 0) > 0) {
            pages.push(dr);
        }
    });

    return (
        <div className="w-full bg-white print:bg-white text-black p-0 m-0">
            {pages.map((emp, index) => {
                const manager = emp._managerId ? employeeMap[emp._managerId] : null;
                const pageDrs = (emp._directs || []).map(id => employeeMap[id]).filter(Boolean).sort((a, b) => sortEmployees(a, b, ceoId));
                const pageMatrix = (emp._matrix || []).map(id => employeeMap[id]).filter(Boolean).sort((a, b) => sortEmployees(a, b, ceoId));

                const hasDrs = pageDrs.length > 0;
                const hasMatrix = pageMatrix.length > 0;

                // Layout ratio logic
                let matrixWidthClass = "flex-1";
                let drWidthClass = "flex-1";
                if (hasMatrix && hasDrs) {
                    const diff = pageDrs.length - pageMatrix.length;
                    if (diff >= 4) {
                        drWidthClass = "flex-[2]";
                        matrixWidthClass = "flex-1";
                    } else if (diff <= -4) {
                        matrixWidthClass = "flex-[2]";
                        drWidthClass = "flex-1";
                    }
                }

                return (
                    <div key={`print-${emp._id}-${index}`} className="w-full min-h-[100vh] py-10 px-8 flex justify-center items-start box-border" style={{ pageBreakAfter: index === pages.length - 1 ? 'auto' : 'always' }}>
                        <div className="flex gap-8 w-full items-start max-w-7xl justify-center">
                            
                            {/* LEFT: Matrix Reports */}
                            {hasMatrix && (
                                <div className={`${matrixWidthClass} pt-16 flex flex-col items-center`}>
                                    <div className="text-[10px] font-bold uppercase tracking-widest text-slate-500 mb-4 border-b border-slate-200 pb-1 w-full text-center max-w-[160px]">MATRIX REPORTS</div>
                                    <div className={`grid gap-3 justify-center ${pageMatrix.length > 4 ? 'grid-cols-2' : 'grid-cols-1'}`}>
                                        {pageMatrix.map(m => <PrintTile key={m._id} employee={m} isMatrix />)}
                                    </div>
                                </div>
                            )}

                            {/* MIDDLE: Context & Target */}
                            <div className="w-[240px] flex-shrink-0 flex flex-col items-center">
                                {manager && (
                                    <>
                                        <div className="text-[8px] font-bold uppercase text-slate-400 mb-1 tracking-widest">Line Manager</div>
                                        <PrintTile employee={manager} isLineManager targetLocation={emp['Location Name']} />
                                        <div className="w-px h-6 bg-slate-300 my-1"></div>
                                    </>
                                )}
                                
                                <div className="w-full border-2 border-slate-800 rounded p-3 bg-white mb-6 shadow-sm">
                                    <div className="flex justify-between items-start gap-1 mb-1">
                                        <div className="font-black text-base leading-tight truncate">{emp._formattedName}</div>
                                        <div className="text-[10px] font-bold px-1.5 py-0.5 border border-slate-400 rounded whitespace-nowrap bg-slate-50">{emp._summaryGrade}</div>
                                    </div>
                                    <div className="text-[11px] text-slate-700 font-medium mb-1.5 truncate">{emp['Job Title']}</div>
                                    <div className="text-[9px] text-slate-500 mb-2.5">{emp['Department (Label)']} • {emp['Location Name']}</div>
                                    
                                    {(emp._insights?.matrixCount > 0 || emp._insights?.directCount > 0 || emp._insights?.eaCount > 0) && (
                                        <div className="flex justify-between mt-2 pt-2 border-t border-slate-300 text-[10px] font-bold">
                                            {emp._insights?.matrixCount > 0 ? <span className="text-slate-800">Matrix: {emp._insights?.matrixCount}</span> : <span></span>}
                                            <span className="text-slate-800">{emp._insights?.eaCount > 0 ? `${emp._insights.directCount} + EA` : `Direct: ${emp._insights?.directCount || 0}`}</span>
                                        </div>
                                    )}
                                </div>

                                <div className="w-full flex justify-center gap-6 px-2">
                                    {hasMatrix && (
                                        <div className="flex-1">
                                            <div className="text-[9px] font-bold uppercase tracking-widest border-b border-slate-300 pb-0.5 mb-2 text-slate-500">Matrix</div>
                                            <PrintGradeList gradesObj={emp._insights.matrixGrades} />
                                        </div>
                                    )}
                                    {hasDrs && (
                                        <div className="flex-1">
                                            <div className="text-[9px] font-bold uppercase tracking-widest border-b border-slate-300 pb-0.5 mb-2 text-slate-500">Direct</div>
                                            <PrintGradeList gradesObj={emp._insights.directGrades} />
                                        </div>
                                    )}
                                </div>
                            </div>

                            {/* RIGHT: Direct Reports */}
                            {hasDrs && (
                                <div className={`${drWidthClass} pt-16 flex flex-col items-center`}>
                                    <div className="text-[10px] font-bold uppercase tracking-widest text-slate-500 mb-4 border-b border-slate-200 pb-1 w-full text-center max-w-[160px]">DIRECT REPORTS</div>
                                    <div className={`grid gap-3 justify-center ${pageDrs.length > 4 ? 'grid-cols-2' : 'grid-cols-1'}`}>
                                        {pageDrs.map(d => <PrintTile key={d._id} employee={d} />)}
                                    </div>
                                </div>
                            )}
                            
                        </div>
                    </div>
                );
            })}
        </div>
    );
};

// --- Compare View Components ---
const ColoredGradeList = ({ gradesObj, textClass }) => {
    if (!gradesObj || Object.keys(gradesObj).length === 0) {
        return <div className="text-[10px] text-slate-400 italic">None</div>;
    }
    const sorted = Object.entries(gradesObj).sort((a, b) => {
        const rankA = getRank(a[0]);
        const rankB = getRank(b[0]);
        if (rankA !== rankB) return rankA - rankB;
        return b[1] - a[1]; 
    });
    return (
        <div className="flex flex-col gap-y-1 text-[10px]">
            {sorted.map(([g, c]) => (
                <div key={g} className="flex justify-between items-center border-b border-slate-100 pb-0.5">
                    <span className="text-slate-600 font-medium truncate pr-1">{g}</span>
                    <span className={`font-bold ${textClass}`}>{c}</span>
                </div>
            ))}
        </div>
    );
};

const CompareReporteeTile = ({ employee, isMatrix }) => (
    <div className={`p-3 border bg-white rounded-xl flex flex-col text-slate-800 break-inside-avoid w-full shadow-sm ${isMatrix ? 'border-2 border-dashed border-purple-300 bg-purple-50/50' : 'border-solid border-slate-200'}`}>
        <div className="flex justify-between items-start gap-1 mb-1">
            <div className="font-bold text-xs leading-tight truncate pr-1">{employee._formattedName}</div>
            <div className={`text-[10px] font-bold px-1.5 py-0.5 rounded border whitespace-nowrap flex-shrink-0 ${isMatrix ? 'bg-white border-purple-200 text-purple-700' : 'bg-slate-50 border-slate-300 text-slate-700'}`}>{employee._summaryGrade}</div>
        </div>
        <div className="text-[11px] text-slate-500 truncate font-medium">{employee['Job Title'] || 'No Title'}</div>
    </div>
);

const CompareView = ({ compareList, employeeMap, ceoId }) => {
    const [activeColor, setActiveColor] = useState('blue');
    const [isIndContextOpen, setIsIndContextOpen] = useState(false);
    const [isOrgContextOpen, setIsOrgContextOpen] = useState(true);

    const colors = [
        { id: 'blue', bg: 'bg-blue-500', text: 'text-blue-700', border: 'border-blue-500', light: 'bg-blue-50' },
        { id: 'green', bg: 'bg-green-500', text: 'text-green-700', border: 'border-green-500', light: 'bg-green-50' },
        { id: 'red', bg: 'bg-red-500', text: 'text-red-700', border: 'border-red-500', light: 'bg-red-50' },
        { id: 'orange', bg: 'bg-orange-500', text: 'text-orange-700', border: 'border-orange-500', light: 'bg-orange-50' },
        { id: 'purple', bg: 'bg-purple-500', text: 'text-purple-700', border: 'border-purple-500', light: 'bg-purple-50' },
    ];
    
    useEffect(() => {
        if (compareList[activeColor].length === 0) {
            const firstPopulated = colors.find(c => compareList[c.id].length > 0);
            if (firstPopulated) setActiveColor(firstPopulated.id);
        }
    }, [compareList, activeColor]);

    const emps = compareList[activeColor].map(id => employeeMap[id]).filter(Boolean);
    const activeColorObj = colors.find(c => c.id === activeColor);

    if (Object.values(compareList).every(arr => arr.length === 0)) {
        return (
            <div className="flex flex-col items-center justify-center h-full text-slate-500 space-y-4 w-full bg-white rounded-xl shadow-sm border border-slate-200">
                <Users size={48} className="text-slate-300" />
                <p>No employees added to compare yet.</p>
                <p className="text-sm italic">Right-click any employee card in the Org Chart to add them.</p>
            </div>
        );
    }

    return (
        <div className="w-full h-full flex flex-col overflow-hidden print:hidden bg-white rounded-xl shadow-sm border border-slate-200 min-h-0">
            {/* Header Tabs */}
            <div className="flex justify-between items-center px-6 py-4 border-b border-slate-100 flex-shrink-0">
                <div className="flex gap-3 bg-slate-50 p-1.5 rounded-xl border border-slate-200">
                    {colors.map(c => (
                        <button key={c.id} onClick={() => setActiveColor(c.id)} className={`w-10 h-10 rounded-lg border-2 ${activeColor === c.id ? c.border : 'border-transparent'} ${c.bg} opacity-80 hover:opacity-100 relative transition-all`}>
                            {compareList[c.id].length > 0 && <span className="absolute -top-2 -right-2 bg-white text-xs font-bold rounded-full w-5 h-5 flex items-center justify-center text-slate-800 shadow-md border border-slate-100">{compareList[c.id].length}</span>}
                        </button>
                    ))}
                </div>
            </div>

            {/* Scrollable Compare Canvas */}
            <div className="flex-1 overflow-auto w-full bg-slate-50" style={{ scrollbarWidth: 'thin' }}>
                <div className="flex flex-col p-6 pt-0 w-max mx-auto min-h-full">
                     
                     {/* Sticky Header Row */}
                     <div className="sticky top-0 z-30 flex gap-6 pb-4 pt-6 bg-slate-50 border-b border-slate-200/50 mb-4 shadow-[0_4px_6px_-1px_rgb(248,250,252)]">
                         {emps.map(emp => (
                             <div key={`header-${emp._id}`} className={`w-[320px] bg-white rounded-xl shadow-md border-t-4 ${activeColorObj.border} border-x border-b border-slate-200 p-4 relative`}>
                                 <div className="flex justify-between items-start gap-2 mb-1">
                                     <div className="font-bold text-slate-800 text-lg leading-tight truncate">{emp._formattedName}</div>
                                     {emp._cardGrade && emp._cardGrade !== 'Unspecified' && <span className="bg-slate-100 text-slate-600 px-2 py-0.5 rounded text-xs font-bold border border-slate-200 shadow-sm flex-shrink-0">{emp._cardGrade}</span>}
                                 </div>
                                 <div className="text-sm text-slate-600 font-medium mb-1.5 truncate">{emp['Job Title'] || 'No Title'}</div>
                                 <div className="text-[10px] text-slate-500 flex items-center gap-1 font-medium"><MapPin size={10}/>{emp['Location Name']}</div>
                                 {emp._isExCom && <span className="absolute -top-1.5 -right-1 bg-amber-100 text-amber-700 text-[9px] font-bold px-1.5 py-0.5 rounded shadow-sm border border-amber-200">ExCom</span>}
                                 {emp._isOpCom && <span className="absolute -top-1.5 -right-1 bg-indigo-100 text-indigo-700 text-[9px] font-bold px-1.5 py-0.5 rounded shadow-sm border border-indigo-200">OpCom</span>}
                                 
                                 {(() => {
                                     const insights = emp._insights || {};
                                     const isIC = (insights.directCount || 0) === 0 && (insights.eaCount || 0) === 0 && (insights.matrixCount || 0) === 0;
                                     
                                     if (isIC) {
                                         return (
                                             <div className="mt-3 flex justify-center items-center text-[10px] font-semibold pt-2 border-t border-slate-100 text-slate-400 italic">
                                                 Individual Contributor
                                             </div>
                                         );
                                     }

                                     if ((insights.directCount || 0) === 0 && (insights.eaCount || 0) === 0 && (insights.matrixCount || 0) > 0) {
                                         return (
                                             <div className="mt-3 flex justify-end items-center text-[10px] font-semibold pt-2 border-t border-slate-100 text-slate-600">
                                                 <div className="flex items-center px-1.5 py-0.5 rounded bg-purple-50 text-purple-700">
                                                     <span>{formatNum(insights.matrixCount)} Matrix</span>
                                                 </div>
                                             </div>
                                         );
                                     }

                                     return (
                                         <div className="mt-3 flex justify-between items-center text-[10px] font-semibold pt-2 border-t border-slate-100 text-slate-600">
                                             <div className="flex items-center px-1.5 py-0.5 rounded bg-blue-50 text-slate-700">
                                                 <User size={12} className="mr-1 text-blue-500"/> 
                                                 {insights.eaCount > 0 ? `${insights.directCount} + EA` : `${formatNum(insights.directCount)} Direct`}
                                             </div>
                                             
                                             {insights.matrixCount > 0 && (
                                                 <div className="flex items-center px-1.5 py-0.5 rounded bg-purple-50 text-purple-700">
                                                     <span>{formatNum(insights.matrixCount)} Matrix</span>
                                                 </div>
                                             )}

                                             <div className="flex items-center px-1.5 py-0.5 rounded bg-slate-100 text-slate-700">
                                                 <Users size={12} className="mr-1 text-slate-500"/> {formatNum(insights.totalTeam)} Team
                                             </div>
                                         </div>
                                     );
                                 })()}
                             </div>
                         ))}
                     </div>

                     {/* Ind Context Accordion Header */}
                     <div className="flex items-center justify-between w-full bg-slate-200/50 border border-slate-300 px-4 py-2.5 rounded-lg cursor-pointer hover:bg-slate-200 transition-colors" onClick={() => setIsIndContextOpen(!isIndContextOpen)}>
                        <span className="font-bold text-slate-700 text-xs uppercase tracking-wider">Individual Context</span>
                        {isIndContextOpen ? <ChevronDown size={16} className="text-slate-500"/> : <ChevronRight size={16} className="text-slate-500"/>}
                     </div>

                     {/* Ind Context Content Row */}
                     {isIndContextOpen && (
                        <div className="flex gap-6 mt-4 mb-6">
                           {emps.map(emp => (
                               <div key={`ind-${emp._id}`} className="w-[320px] bg-white p-5 rounded-xl border border-slate-200 shadow-sm flex flex-col">
                                   <div className="grid grid-cols-2 gap-4 text-xs mb-2">
                                       <div><span className="text-slate-400 block">Total Tenure</span><span className="font-bold text-slate-700"><TenureDisplay employee={emp}/></span></div>
                                       <div><span className="text-slate-400 block">Time in Role</span><span className="font-bold text-slate-700">{emp._timeInRoleFormatted}</span></div>
                                       <div><span className="text-slate-400 block">Since Promoted</span><span className="font-bold text-green-700">{emp._lastPromotionFormatted}</span></div>
                                       <div><span className="text-slate-400 block">With Manager</span><span className="font-bold text-indigo-700">{emp._timeWithManagerFormatted}</span></div>
                                   </div>
                                   
                                   {emp._timeline && emp._timeline.length > 0 && (
                                       <div className="mt-5 border-t border-slate-100 pt-4">
                                           <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-4">Role History</div>
                                           <div className="relative ml-3 space-y-0 pb-1">
                                               {emp._timeline.map((item, i) => {
                                                   if (item.isGap) return (<div key={i} className="relative pl-7 py-3"><div className="absolute left-0 top-0 bottom-0 w-px border-l-2 border-dashed border-slate-300"></div><div className="text-xs italic text-slate-500 flex items-center bg-white relative z-10 py-1 -ml-3 pl-3">Data Unavailable <span className="mx-2">•</span> {item.duration} gap</div></div>);
                                                   let dotColor = 'bg-slate-400 ring-white';
                                                   if (item.type === 'promotion') dotColor = 'bg-green-500 ring-white';
                                                   else if (item.type === 'transfer') dotColor = 'bg-blue-500 ring-white';
                                                   else if (item.type === 'hire') dotColor = 'bg-purple-500 ring-white';
                                                   else if (item.type === 'company_transfer') dotColor = 'bg-orange-500 ring-white';
                                                   const showLocation = emp._hasLocationTransfer && item.location && item.location !== employee['Location Name'];
                                                   return (
                                                       <div key={i} className="relative pl-7 pb-5">
                                                           {i !== emp._timeline.length - 1 && <div className="absolute left-0 top-2 bottom-0 w-px bg-slate-200 -ml-[1px]"></div>}
                                                           <div className={`absolute -left-[6px] top-1.5 h-3 w-3 rounded-full ring-[3px] ${dotColor} z-10`}></div>
                                                           <div className="text-sm font-bold text-slate-800 leading-tight flex flex-wrap gap-2 items-center">
                                                               <span>{item.title || 'Unknown Title'}</span>
                                                               {item.grade && item.grade !== 'Unspecified' && <span className="text-slate-500 font-normal border border-slate-200 px-1.5 rounded text-[10px] bg-slate-50">{item.grade}</span>}
                                                           </div>
                                                           <div className="text-[11px] text-slate-500 mt-1.5 flex items-center gap-3 flex-wrap">
                                                               <span className="flex items-center"><Clock size={10} className="mr-1.5"/> {item.durationFormatted}</span>
                                                               {showLocation && <span className="flex items-center text-blue-600 font-medium"><MapPin size={10} className="mr-1.5"/> {item.location}</span>}
                                                           </div>
                                                       </div>
                                                   );
                                               })}
                                           </div>
                                       </div>
                                   )}
                               </div>
                           ))}
                        </div>
                     )}

                     {/* Org Context Accordion Header */}
                     <div className="flex items-center justify-between w-full bg-slate-200/50 border border-slate-300 px-4 py-2.5 rounded-lg cursor-pointer hover:bg-slate-200 transition-colors mb-4 mt-2" onClick={() => setIsOrgContextOpen(!isOrgContextOpen)}>
                        <span className="font-bold text-slate-700 text-xs uppercase tracking-wider">Organizational Context</span>
                        {isOrgContextOpen ? <ChevronDown size={16} className="text-slate-500"/> : <ChevronRight size={16} className="text-slate-500"/>}
                     </div>

                     {/* Org Context Content Row */}
                     {isOrgContextOpen && (
                        <div className="flex gap-6 pb-20">
                            {emps.map(emp => {
                                const pageDrs = (emp._directs || []).map(id => employeeMap[id]).filter(Boolean).sort((a, b) => sortEmployees(a, b, ceoId));
                                const pageMatrix = (emp._matrix || []).map(id => employeeMap[id]).filter(Boolean).sort((a, b) => sortEmployees(a, b, ceoId));

                                return (
                                    <div key={`org-${emp._id}`} className="w-[320px] flex flex-col gap-5">
                                       {/* Grade Summary Box */}
                                       <div className="flex gap-4 w-full bg-white p-4 rounded-xl border border-slate-200 shadow-sm items-start">
                                          <div className="flex-1">
                                              <div className="text-[10px] font-bold text-blue-600 uppercase tracking-wider mb-2 border-b border-blue-100 pb-1">Direct Summary</div>
                                              <ColoredGradeList gradesObj={emp._insights?.directGrades} textClass="text-blue-700" />
                                          </div>
                                          <div className="flex-1">
                                              <div className="text-[10px] font-bold text-purple-600 uppercase tracking-wider mb-2 border-b border-purple-100 pb-1">Matrix Summary</div>
                                              <ColoredGradeList gradesObj={emp._insights?.matrixGrades} textClass="text-purple-700" />
                                          </div>
                                       </div>

                                       {/* Direct Reports */}
                                       {pageDrs.length > 0 && (
                                           <div className="flex flex-col gap-2.5 mt-2">
                                              <div className="text-[11px] font-bold text-slate-500 uppercase tracking-wider border-b border-slate-200 pb-1 px-1">Direct Reports ({pageDrs.length})</div>
                                              {pageDrs.map(dr => <CompareReporteeTile key={dr._id} employee={dr} isMatrix={false} />)}
                                           </div>
                                       )}

                                       {/* Matrix Reports */}
                                       {pageMatrix.length > 0 && (
                                           <div className="flex flex-col gap-2.5 mt-2">
                                              <div className="text-[11px] font-bold text-purple-500 uppercase tracking-wider border-b border-purple-200 pb-1 px-1">Matrix Reports ({pageMatrix.length})</div>
                                              {pageMatrix.map(mr => <CompareReporteeTile key={mr._id} employee={mr} isMatrix={true} />)}
                                           </div>
                                       )}
                                       
                                       {pageDrs.length === 0 && pageMatrix.length === 0 && (
                                           <div className="text-center text-sm text-slate-400 italic mt-4 bg-white p-4 rounded-xl border border-slate-200 shadow-sm">
                                                Individual Contributor
                                           </div>
                                       )}
                                    </div>
                                )
                            })}
                        </div>
                     )}
                </div>
            </div>
        </div>
    )
};

function EmployeeCard({ employee, ceoId, globalMetrics, isActive, isMatrixNode, viewMode, onClick, onSelectDirect, onSelectMatrix, onContextMenu }) {
  const [showTooltip, setShowTooltip] = useState(false);
  const [showTimeline, setShowTimeline] = useState(false);
  
  const [gradeTooltip, setGradeTooltip] = useState(null); 
  const [tooltipPos, setTooltipPos] = useState({ top: 0, left: 0 });
  
  const hideTimeout = useRef(null);
  const hideGradeTimeout = useRef(null);

  const insights = employee._insights || { genderCount: { male: 0, female: 0, other: 0 }};
  const isIndividualContributor = insights.directCount === 0 && (insights.eaCount || 0) === 0 && insights.matrixCount === 0;

  const handleMouseEnterInfo = (e) => {
      clearTimeout(hideTimeout.current);
      const cardRect = e.currentTarget.closest('.group').getBoundingClientRect();
      const tooltipWidth = 360; 
      const estimatedHeight = 450;
      
      let style = { 
          position: 'fixed', 
          zIndex: 99999,
          maxHeight: 'calc(100vh - 100px)',
          overflowY: 'auto'
      };

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

      setTooltipPos(style);
      setShowTooltip(true);
  };

  const handleMouseLeaveInfo = () => hideTimeout.current = setTimeout(() => setShowTooltip(false), 200);

  const handleMouseEnterGrade = (e, type) => {
    clearTimeout(hideGradeTimeout.current);
    const pillRect = e.currentTarget.getBoundingClientRect();
    const cardRect = e.currentTarget.closest('.group').getBoundingClientRect();
    const tooltipWidth = 192;
    
    let style = { position: 'fixed', zIndex: 99999 };
    
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
    
    setTooltipPos(style);
    setGradeTooltip(type);
  };

  const handleMouseLeaveGrade = () => hideGradeTimeout.current = setTimeout(() => setGradeTooltip(null), 200);

  const totalGender = insights.genderCount.male + insights.genderCount.female + insights.genderCount.other;
  const malePct = totalGender > 0 ? Math.round((insights.genderCount.male / totalGender) * 100) : 0;
  const femalePct = totalGender > 0 ? Math.round((insights.genderCount.female / totalGender) * 100) : 0;
  
  const isTopNode = employee._id === ceoId;

  let cardClasses = "relative w-64 min-w-[16rem] mx-auto bg-white rounded-xl shadow-md border p-4 transition-all duration-200 flex flex-col group ";
  if (isActive) cardClasses += "border-blue-500 ring-4 ring-blue-100 shadow-xl scale-105 cursor-default z-10";
  else if (isMatrixNode) cardClasses += "border-purple-300 border-dashed hover:border-purple-500 hover:shadow-lg cursor-pointer";
  else cardClasses += "border-slate-200 hover:border-blue-400 hover:shadow-lg cursor-pointer";

  let popupHeaderClass = "px-3 py-2 border-b text-xs font-bold uppercase tracking-wider flex justify-between ";
  if (gradeTooltip === 'direct') popupHeaderClass += "bg-blue-100 text-blue-800 border-blue-200";
  else if (gradeTooltip === 'matrix') popupHeaderClass += "bg-purple-100 text-purple-800 border-purple-200";
  else popupHeaderClass += "bg-slate-100 text-slate-700 border-slate-200";

  return (
    <div id={isActive ? "active-employee-card" : undefined} className={`relative flex justify-center w-full ${isActive ? 'z-10' : 'z-0'}`}>
      <div className={cardClasses} onClick={!isActive ? onClick : undefined} onContextMenu={(e) => onContextMenu && onContextMenu(e, employee)}>
        
        {employee._isExCom && (
            <div className="absolute top-0 right-10 z-20 drop-shadow-sm" title="Executive Committee">
                <div className="bg-amber-100 text-amber-700 text-xs font-bold px-2 pt-1 pb-2.5 border-x border-b border-amber-200" style={{ clipPath: 'polygon(0 0, 100% 0, 100% 100%, 50% 80%, 0 100%)' }}>
                    Ex
                </div>
            </div>
        )}
        {employee._isOpCom && (
            <div className="absolute top-0 right-10 z-20 drop-shadow-sm" title="Operations Committee">
                <div className="bg-indigo-100 text-indigo-700 text-xs font-bold px-2 pt-1 pb-2.5 border-x border-b border-indigo-200" style={{ clipPath: 'polygon(0 0, 100% 0, 100% 100%, 50% 80%, 0 100%)' }}>
                    Op
                </div>
            </div>
        )}

        <div className="absolute top-3 right-3 text-slate-400 hover:text-blue-600 z-20 cursor-help" onMouseEnter={handleMouseEnterInfo} onMouseLeave={handleMouseLeaveInfo} onClick={(e) => e.stopPropagation()}>
          <Info size={18} />
        </div>

        <div className="flex items-center space-x-3 mb-3 pr-6">
          <div className={`w-12 h-12 rounded-full flex-shrink-0 flex items-center justify-center text-white font-bold shadow-sm ${isActive ? 'bg-blue-600' : isMatrixNode ? 'bg-purple-500' : 'bg-slate-700'}`}>
            {employee._initials}
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="font-bold text-slate-800 truncate text-sm" title={employee['Display Name']}>{employee._formattedName}</h3>
            <p className="text-xs text-slate-500 truncate mt-0.5" title={employee['Job Title']}>{employee['Job Title'] || 'No Title'}</p>
          </div>
        </div>
        
        <div className="text-xs text-slate-600 bg-slate-50 p-2 rounded-md flex flex-col gap-1.5">
          <div className="flex items-center justify-between">
              <div className="flex items-center space-x-1 truncate pr-2"><Building2 size={12} className="flex-shrink-0"/> <span className="truncate">{employee['Department (Label)'] || 'N/A'}</span></div>
              {employee._cardGrade && employee._cardGrade !== 'Unspecified' && (
                  <div className="flex items-center space-x-1 text-slate-500 font-bold whitespace-nowrap" title={employee._cardGrade}>
                      <Award size={12} className="flex-shrink-0"/> <span>{employee._cardGrade}</span>
                  </div>
              )}
          </div>
          <div className="flex items-center justify-between">
              <div className="flex items-center space-x-1 truncate pr-2"><MapPin size={12} className="flex-shrink-0"/> <span className="truncate">{employee['Location Name'] || 'N/A'}</span></div>
              <div className="flex items-center space-x-1 text-slate-500 flex-shrink-0 font-medium"><Clock size={12} /> <TenureDisplay employee={employee} /></div>
          </div>
        </div>

        {isIndividualContributor ? (
            <div className="mt-3 flex justify-center items-center text-[10px] font-semibold pt-2 border-t text-slate-400 italic">
                Individual Contributor
            </div>
        ) : (insights.directCount === 0 && (insights.eaCount || 0) === 0) && insights.matrixCount > 0 ? (
            <div className="mt-3 flex justify-end items-center text-[10px] font-semibold pt-2 border-t text-slate-600">
                <div onMouseEnter={(e) => handleMouseEnterGrade(e, 'matrix')} onMouseLeave={handleMouseLeaveGrade} onClick={(e) => { e.stopPropagation(); if(onSelectMatrix) onSelectMatrix(); }} className={`flex items-center px-1 py-0.5 rounded transition-colors cursor-pointer ${isActive && viewMode === 'matrix' ? 'bg-purple-100 text-purple-700 ring-1 ring-purple-300' : 'hover:bg-purple-50 text-purple-600'}`}>
                    <span>{formatNum(insights.matrixCount)} Matrix</span>
                </div>
            </div>
        ) : (
            <div className="mt-3 flex justify-between items-center text-[10px] font-semibold pt-2 border-t text-slate-600">
                <div onMouseEnter={(e) => handleMouseEnterGrade(e, 'direct')} onMouseLeave={handleMouseLeaveGrade} onClick={(e) => { e.stopPropagation(); if(onSelectDirect) onSelectDirect(); }} className={`flex items-center px-1 py-0.5 rounded transition-colors cursor-pointer ${isActive && viewMode === 'direct' ? 'bg-blue-100 text-blue-800 ring-1 ring-blue-300' : 'hover:bg-blue-50 text-slate-600'}`}>
                    <User size={12} className={`mr-1 ${isActive && viewMode === 'direct' ? 'text-blue-600' : 'text-blue-500'}`}/> 
                    {insights.eaCount > 0 ? `${insights.directCount} + EA` : `${formatNum(insights.directCount)} Direct`}
                </div>
                
                {insights.matrixCount > 0 && (
                    <div onMouseEnter={(e) => handleMouseEnterGrade(e, 'matrix')} onMouseLeave={handleMouseLeaveGrade} onClick={(e) => { e.stopPropagation(); if(onSelectMatrix) onSelectMatrix(); }} className={`flex items-center px-1 py-0.5 rounded transition-colors cursor-pointer ${isActive && viewMode === 'matrix' ? 'bg-purple-100 text-purple-700 ring-1 ring-purple-300' : 'hover:bg-purple-50 text-purple-600'}`}>
                        <span>{formatNum(insights.matrixCount)} Matrix</span>
                    </div>
                )}

                <div onMouseEnter={(e) => handleMouseEnterGrade(e, 'team')} onMouseLeave={handleMouseLeaveGrade} className="flex items-center cursor-help px-1 py-0.5 hover:bg-slate-100 rounded">
                    <Users size={12} className="mr-1 text-slate-500"/> {formatNum(insights.totalTeam)} Team
                </div>
            </div>
        )}
      </div>

      {gradeTooltip && createPortal(
          <div style={tooltipPos} className="fixed w-48 bg-white rounded-lg shadow-[0_0_20px_rgba(0,0,0,0.15)] border border-slate-200 text-sm overflow-hidden animate-scale-in z-[99999]" onMouseEnter={() => clearTimeout(hideGradeTimeout.current)} onMouseLeave={handleMouseLeaveGrade}>
              <div className={popupHeaderClass}><span>{gradeTooltip === 'direct' ? 'DR Summary' : gradeTooltip === 'matrix' ? 'Matrix Summary' : 'Team Summary'}</span></div>
              <div className="p-2 max-h-64 overflow-y-auto" style={{ scrollbarWidth: 'thin' }}>
                  {gradeTooltip === 'direct' && renderGradesList(insights.directGrades)}
                  {gradeTooltip === 'matrix' && renderGradesList(insights.matrixGrades)}
                  {gradeTooltip === 'team' && renderGradesList(insights.teamGrades)}
              </div>
          </div>,
          document.body
      )}

      {showTooltip && createPortal(
        <div style={tooltipPos} className="fixed w-[360px] bg-white rounded-xl shadow-[0_0_40px_rgba(0,0,0,0.2)] border border-slate-200 p-0 text-sm overflow-hidden flex flex-col animate-scale-in z-[99999]" onMouseEnter={() => clearTimeout(hideTimeout.current)} onMouseLeave={handleMouseLeaveInfo}>
          <div className="bg-slate-800 text-white px-5 py-4 border-b flex items-center flex-shrink-0">
            <Info size={18} className="mr-2" />
            <span className="font-bold text-base">Spotlight</span>
          </div>
          
          <div className="p-5 space-y-6 overflow-y-auto flex-1" style={{ scrollbarWidth: 'thin' }}>
            <div>
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-3 pb-1 border-b border-slate-100">Individual Context</h4>
                <div className="grid grid-cols-2 gap-3 text-sm mb-4">
                    <div className="bg-slate-50 p-2.5 rounded-lg border border-slate-100 flex flex-col items-center text-center"><span className="text-slate-400 font-medium mb-1 text-xs">Total Tenure</span><span className="font-bold text-slate-700 flex items-center"><CalendarDays size={14} className="mr-1.5"/> <TenureDisplay employee={employee} /></span></div>
                    <div className="bg-slate-50 p-2.5 rounded-lg border border-slate-100 flex flex-col items-center text-center"><span className="text-slate-400 font-medium mb-1 text-xs">Time in Role</span><span className="font-bold text-slate-700 flex items-center"><Clock size={14} className="mr-1.5"/> {employee._timeInRoleFormatted}</span></div>
                    <div className="bg-slate-50 p-2.5 rounded-lg border border-slate-100 flex flex-col items-center text-center"><span className="text-slate-400 font-medium mb-1 text-xs">Since Promoted</span><span className={`font-bold flex items-center ${employee._lastPromotionFormatted !== '-' ? 'text-green-700' : 'text-slate-400'}`}><Clock size={14} className="mr-1.5"/> {employee._lastPromotionFormatted}</span></div>
                    <div className="bg-slate-50 p-2.5 rounded-lg border border-slate-100 flex flex-col items-center text-center"><span className="text-slate-400 font-medium mb-1 text-xs">With Manager</span><span className="font-bold text-indigo-700 flex items-center"><Users size={14} className="mr-1.5"/> {employee._timeWithManagerFormatted}</span></div>
                </div>

                {employee._timeline && employee._timeline.length > 0 && (
                    <div className="mb-2 mt-4 pt-4 border-t border-slate-100">
                        <button className="text-sm text-blue-600 hover:text-blue-700 flex items-center font-bold w-full focus:outline-none transition-colors" onClick={(e) => { e.stopPropagation(); setShowTimeline(!showTimeline); }}>
                            {showTimeline ? 'Hide Role History' : 'View Role History'}
                        </button>
                        {showTimeline && (
                            <div className="mt-4 relative ml-3 space-y-0 pb-1">
                                {employee._timeline.map((item, i) => {
                                    if (item.isGap) return (<div key={i} className="relative pl-7 py-3"><div className="absolute left-0 top-0 bottom-0 w-px border-l-2 border-dashed border-slate-300"></div><div className="text-xs italic text-slate-50 flex items-center bg-white relative z-10 py-1 -ml-3 pl-3">Data Unavailable <span className="mx-2">•</span> {item.duration} gap</div></div>);
                                    let dotColor = 'bg-slate-400 ring-white';
                                    if (item.type === 'promotion') dotColor = 'bg-green-500 ring-white';
                                    else if (item.type === 'transfer') dotColor = 'bg-blue-500 ring-white';
                                    else if (item.type === 'hire') dotColor = 'bg-purple-500 ring-white';
                                    else if (item.type === 'company_transfer') dotColor = 'bg-orange-500 ring-white';
                                    const showLocation = employee._hasLocationTransfer && item.location && item.location !== employee['Location Name'];
                                    return (
                                        <div key={i} className="relative pl-7 pb-5">
                                            {i !== employee._timeline.length - 1 && <div className="absolute left-0 top-2 bottom-0 w-px bg-slate-200 -ml-[1px]"></div>}
                                            <div className={`absolute -left-[6px] top-1.5 h-3 w-3 rounded-full ring-[3px] ${dotColor} z-10`}></div>
                                            <div className="text-sm font-bold text-slate-800 leading-tight flex flex-wrap gap-2 items-center">
                                                <span>{item.title || 'Unknown Title'}</span>
                                                {item.grade && item.grade !== 'Unspecified' && <span className="text-slate-500 font-normal border border-slate-200 px-1.5 rounded text-[10px] bg-slate-50">{item.grade}</span>}
                                            </div>
                                            <div className="text-[11px] text-slate-500 mt-1.5 flex items-center gap-3 flex-wrap">
                                                <span className="flex items-center"><Clock size={10} className="mr-1.5"/> {item.durationFormatted}</span>
                                                {showLocation && <span className="flex items-center text-blue-600 font-medium"><MapPin size={10} className="mr-1.5"/> {item.location}</span>}
                                            </div>
                                        </div>
                                    );
                                })}
                            </div>
                        )}
                    </div>
                )}
            </div>

            {/* --- ORGANIZATIONAL CONTEXT --- */}
            {isIndividualContributor ? (
               <div className="pt-4 text-center border-t border-slate-100 mt-4">
                   <p className="text-base font-semibold text-slate-600">Individual Contributor</p>
                   <p className="text-sm text-slate-400 mt-1">No reports.</p>
               </div>
            ) : (
              <div className="mt-4">
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-3 pb-1 border-b border-slate-100">Organizational Context</h4>
                
                {!isTopNode && employee._isExCom && globalMetrics.exCom && (
                    <BenchmarkBox title="ExCom Benchmark" borderColor="border-amber-200" titleColor="text-amber-600" bgClass="bg-amber-50/20">
                        <MetricScale label="Direct Reports" min={globalMetrics.exCom.drMin} max={globalMetrics.exCom.drMax} median={globalMetrics.exCom.drMedian} value={insights.directCount} />
                        <MetricScale label="Total Team Size" min={globalMetrics.exCom.teamMin} max={globalMetrics.exCom.teamMax} median={globalMetrics.exCom.teamMedian} value={insights.totalTeam} />
                        {insights.pctOfManagerTeam !== undefined && insights.managerValidDrCount > 0 && (() => {
                            const expected = 100 / insights.managerValidDrCount;
                            let shareColor = "text-slate-700";
                            if (insights.pctOfManagerTeam <= expected * 0.92) shareColor = "text-red-600";
                            else if (insights.pctOfManagerTeam >= expected * 1.18) shareColor = "text-green-600";
                            else shareColor = "text-blue-600";
                            return (
                                <div className="mt-4 p-3 rounded-lg border border-slate-200 bg-slate-50 flex justify-between items-center">
                                    <span className="text-sm font-bold text-slate-600">Share of Manager's Team</span>
                                    <span className={`font-bold text-xl leading-tight ${shareColor}`}>{insights.pctOfManagerTeam}%</span>
                                </div>
                            );
                        })()}
                    </BenchmarkBox>
                )}
                
                {!isTopNode && !employee._isExCom && (
                    <>
                        {insights.peerMedianDirects !== undefined && (
                            <BenchmarkBox title="Peer Benchmark">
                                <MetricScale label="Direct Reports" min={insights.peerMinDirects} max={insights.peerMaxDirects} median={insights.peerMedianDirects} value={insights.directCount} />
                                {insights.pctOfManagerTeam !== undefined && insights.managerValidDrCount > 0 && (() => {
                                    const expected = 100 / insights.managerValidDrCount;
                                    let shareColor = "text-slate-700";
                                    if (insights.pctOfManagerTeam <= expected * 0.92) shareColor = "text-red-600";
                                    else if (insights.pctOfManagerTeam >= expected * 1.18) shareColor = "text-green-600";
                                    else shareColor = "text-blue-600";
                                    return (
                                        <div className="mt-4 p-3 rounded-lg border border-slate-200 bg-slate-50 flex justify-between items-center">
                                            <span className="text-sm font-bold text-slate-600">Share of Manager's Team</span>
                                            <span className={`font-bold text-xl leading-tight ${shareColor}`}>{insights.pctOfManagerTeam}%</span>
                                        </div>
                                    );
                                })()}
                            </BenchmarkBox>
                        )}
                        {employee._isOpCom && globalMetrics.opCom && (
                            <BenchmarkBox title="OpCom Benchmark" borderColor="border-indigo-200" titleColor="text-indigo-600" bgClass="bg-indigo-50/20">
                                <MetricScale label="Direct Reports" min={globalMetrics.opCom.drMin} max={globalMetrics.opCom.drMax} median={globalMetrics.opCom.drMedian} value={insights.directCount} />
                                <MetricScale label="Total Team Size" min={globalMetrics.opCom.teamMin} max={globalMetrics.opCom.teamMax} median={globalMetrics.opCom.teamMedian} value={insights.totalTeam} />
                            </BenchmarkBox>
                        )}
                        {!isTopNode && !employee._isExCom && !employee._isOpCom && globalMetrics.grade && globalMetrics.grade[employee._summaryGrade] && (
                            <BenchmarkBox title="Grade Benchmark" rightElement={<div className="flex items-center space-x-1.5 text-slate-500 font-bold bg-slate-100 px-2 py-0.5 rounded text-xs"><Award size={12} className="flex-shrink-0"/> <span>{employee._summaryGrade}</span></div>}>
                                <MetricScale label="Direct Reports" min={globalMetrics.grade[employee._summaryGrade].drMin} max={globalMetrics.grade[employee._summaryGrade].drMax} median={globalMetrics.grade[employee._summaryGrade].drMedian} value={insights.directCount} />
                                <MetricScale label="Total Team Size" min={globalMetrics.grade[employee._summaryGrade].teamMin} max={globalMetrics.grade[employee._summaryGrade].teamMax} median={globalMetrics.grade[employee._summaryGrade].teamMedian} value={insights.totalTeam} />
                            </BenchmarkBox>
                        )}
                    </>
                )}

                {/* 5. TEAM DIVERSITY */}
                {insights.directCount > 0 && (
                    <div className="mt-5 px-1">
                        <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-3">Team Diversity (DR)</h4>
                        <div className="w-full bg-slate-200 h-2.5 rounded-full overflow-hidden flex mt-2 shadow-inner">
                            {malePct > 0 && <div style={{ width: `${malePct}%` }} className="bg-blue-500 h-full"></div>}
                            {femalePct > 0 && <div style={{ width: `${femalePct}%` }} className="bg-pink-500 h-full"></div>}
                        </div>
                        <div className="flex justify-between text-sm mt-2 text-slate-600 font-medium">
                            <span>Male: <span className="font-bold text-slate-800">{malePct}%</span></span>
                            <span>Female: <span className="font-bold text-slate-800">{femalePct}%</span></span>
                        </div>
                    </div>
                )}
              </div>
            )}
          </div>
        </div>,
        document.body
      )}
    </div>
  );
}

// --- App Entry Point ---
const App = () => {
  const [appTab, setAppTab] = useState('org'); // 'org', 'table', 'compare'
  const [data, setData] = useState([]);
  const [employeeMap, setEmployeeMap] = useState({});
  const [activeEmployeeId, setActiveEmployeeId] = useState(null);
  const [ceoId, setCeoId] = useState(null);
  
  // Filtering & Search
  const [searchQuery, setSearchQuery] = useState('');
  const [isSearchOpen, setIsSearchOpen] = useState(false);
  
  const [showFilterPanel, setShowFilterPanel] = useState(false);
  const [filterMatchMode, setFilterMatchMode] = useState('and'); 
  const [filterConditions, setFilterConditions] = useState([]);
  const [openDropdown, setOpenDropdown] = useState(null);
  
  const [sortConfigs, setSortConfigs] = useState([{ field: 'Grade', dir: 'asc' }]); 
  const [activeCohortScale, setActiveCohortScale] = useState(null);

  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [xlsxLoaded, setXlsxLoaded] = useState(false);
  const [viewMode, setViewMode] = useState('direct'); 
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);

  const [compareList, setCompareList] = useState({ blue: [], green: [], purple: [], orange: [], red: [] });
  const [contextMenu, setContextMenu] = useState(null);
  const [printNodeId, setPrintNodeId] = useState(null);

  const searchRef = useRef(null);
  const tableContainerRef = useRef(null);

  // Global Click Handlers for Dropdowns
  useEffect(() => {
      const handleClickOutside = (e) => {
          if (searchRef.current && !searchRef.current.contains(e.target)) {
              setIsSearchOpen(false);
          }
          if (!e.target.closest('.filter-dropdown-wrapper')) {
              setOpenDropdown(null);
          }
          if (!e.target.closest('.context-menu')) {
              setContextMenu(null);
          }
      };
      document.addEventListener('mousedown', handleClickOutside);
      return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleDragOver = (e) => { e.preventDefault(); setIsDragging(true); };
  const handleDragLeave = () => setIsDragging(false);
  const handleDrop = (e) => {
      e.preventDefault();
      setIsDragging(false);
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
          handleFileUpload(e.dataTransfer.files[0]);
      }
  };

  useEffect(() => {
    if (window.XLSX) { setXlsxLoaded(true); return; }
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.onload = () => setXlsxLoaded(true);
    script.onerror = () => setError("Failed to load Excel parser.");
    document.head.appendChild(script);
  }, []);

  // Print Effect Lifecycle
  useEffect(() => {
    if (printNodeId) {
        const timer = setTimeout(() => {
            window.print();
        }, 500); 
        return () => clearTimeout(timer);
    }
  }, [printNodeId]);

  useEffect(() => {
    const handleAfterPrint = () => {
        if (printNodeId) setPrintNodeId(null);
    };
    window.addEventListener('afterprint', handleAfterPrint);
    return () => window.removeEventListener('afterprint', handleAfterPrint);
  }, [printNodeId]);

  // Handle Tab-specific auto-scrolling
  useEffect(() => { 
    if (appTab === 'org' && activeEmployeeId) {
        setViewMode('direct'); 
        const timer = setTimeout(() => {
            const activeEl = document.getElementById('active-employee-card');
            if (activeEl) {
                activeEl.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
            }
        }, 100);
        return () => clearTimeout(timer);
    }
  }, [activeEmployeeId, appTab]);

  useEffect(() => { 
    if (appTab === 'table' && activeEmployeeId) {
        const timer = setTimeout(() => {
            const activeEl = document.getElementById(`table-row-${activeEmployeeId}`);
            if (activeEl && tableContainerRef.current) {
                tableContainerRef.current.scrollTo({
                    top: Math.max(0, activeEl.offsetTop - 45), // 45px buffer for sticky header
                    behavior: 'smooth'
                });
            }
        }, 100);
        return () => clearTimeout(timer);
    }
  }, [appTab]); // Fire only when tab changes

  const handleEmployeeSelect = (id) => {
      setActiveEmployeeId(id);
      setAppTab('org');
      setViewMode('direct');
  };

  const handleFileUpload = async (file) => {
    if (!window.XLSX) return setError("Excel parsing library is still loading.");
    setLoading(true); setError(null);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(buffer, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const rawData = window.XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { defval: "" });
      let historyData = workbook.SheetNames.length > 1 ? window.XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[1]], { defval: "" }) : [];
      if (rawData.length === 0) throw new Error("Uploaded Excel file is empty.");
      processEmployeeData(rawData, historyData);
    } catch (err) {
      setError(err.message || "Failed to process file.");
    } finally {
      setLoading(false);
    }
  };

  const processEmployeeData = (rawData, historyData) => {
    const empMap = {};
    const directReportsMap = {};
    const matrixReportsMap = {};
    const historyMap = {};

    historyData.forEach(row => {
        const id = row['Users Sys Id']?.toString().trim();
        if (!id) return;
        if (!historyMap[id]) historyMap[id] = [];
        historyMap[id].push(row);
    });

    const isCompanyTransfer = (reasonLow) => ['transfer|company change', 'vol separation|to rpg group company', 'rehire|from rpg group company', 'promotion & company transfer', 'hire|from rpg group company'].some(r => reasonLow.includes(r));
    const corporateValues = ["Corporate", "Manufacturing", "Sales & Marketing", "Sales and Marketing"];

    rawData.forEach(emp => {
      const id = emp['Users Sys Id']?.toString().trim();
      if (!id) return;
      emp._id = id;
      
      // PRE-COMPUTE DATA ONCE FOR FAST UI
      emp._formattedName = formatDisplayFirstLast(emp['Display Name']);
      emp._formattedManagerName = formatDisplayFirstLast(emp['Line Manager Name'] || '');
      emp._initials = String(emp['Display Name'] || '?').split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
      emp._searchString = (emp._formattedName + ' ' + (emp['Job Title'] || '')).toLowerCase();

      emp._groupTenureDate = parseExcelDate(emp['Group Date of Joining']);
      emp._orgTenureDate = parseExcelDate(emp['Hire Date']); 
      emp._tenureDate = emp._groupTenureDate || emp._orgTenureDate;
      
      emp._isDiffTenure = false;
      if (emp._orgTenureDate && emp._groupTenureDate) {
          const diff = Math.abs(emp._orgTenureDate - emp._groupTenureDate);
          if (diff > 86400000) { 
              const d1 = emp._orgTenureDate;
              const d2 = emp._groupTenureDate;
              const isSwap = (d1.getFullYear() === d2.getFullYear()) && 
                             (d1.getDate() === d2.getMonth() + 1) && 
                             (d1.getMonth() + 1 === d2.getDate());
              if (!isSwap) {
                  emp._isDiffTenure = true;
              }
          }
      }

      emp._summaryGrade = processGrade(emp['Pay Grade (Name)']); 
      emp._cardGrade = formatCardGrade(emp['Pay Grade (Name)']); 
      emp['Job Title'] = formatJobTitle(emp['Job Title']);
      
      const business = String(emp['Business (Label)'] || '').trim();
      const companyStr = String(emp['Company (Label)'] || '').trim();
      emp._derivedCompany = corporateValues.includes(business) ? companyStr : business;
      if (!emp._derivedCompany) emp._derivedCompany = 'Unspecified';

      const events = historyMap[id] || [];
      const eventsAsc = [...events].sort((a,b) => (parseExcelDate(a['Effective Start Date']) || new Date(0)) - (parseExcelDate(b['Effective Start Date']) || new Date(0)));

      let timelineRaw = [];
      let currentBlock = null;

      eventsAsc.forEach(ev => {
          const rawTitle = ev['Job Title']?.trim() || emp['Job Title']?.trim(); 
          if (!rawTitle) return;
          const title = formatJobTitle(rawTitle);
          const grade = formatCardGrade(ev['Pay Grade (Name)']);
          const location = (ev['Location (Location Name)'] || ev['Location Name'] || ev['Location'])?.trim();
          const reasonLow = String(ev['Event Reason (Label)'] || '').trim().toLowerCase();
          const startDate = parseExcelDate(ev['Effective Start Date']);

          if (reasonLow.includes('separation')) {
              if (currentBlock) { currentBlock.endDate = startDate; timelineRaw.push(currentBlock); currentBlock = null; }
              return; 
          }

          let isNewBlock = false; let blockType = 'other';
          if (!currentBlock) {
              isNewBlock = true;
              if (isCompanyTransfer(reasonLow)) blockType = 'company_transfer';
              else if (reasonLow.includes('promotion')) blockType = 'promotion';
              else if (reasonLow.includes('transfer')) blockType = 'transfer';
              else blockType = 'hire'; 
          } else {
              const gradeChanged = currentBlock.grade !== 'Unspecified' && grade !== 'Unspecified' && currentBlock.grade !== grade;
              const titleChanged = currentBlock.title !== title;
              const isDataChangeOnly = reasonLow === 'data change' || reasonLow.includes('data change|') || reasonLow.includes('job change');
              if (gradeChanged) { isNewBlock = true; blockType = 'promotion'; }
              else if (isCompanyTransfer(reasonLow)) { if (titleChanged) { isNewBlock = true; blockType = 'company_transfer'; } }
              else if (reasonLow.includes('promotion')) { isNewBlock = true; blockType = 'promotion'; }
              else if (reasonLow.includes('transfer') || reasonLow.includes('rotation')) { if (titleChanged) { isNewBlock = true; blockType = 'transfer'; } }
              else if (!isDataChangeOnly && titleChanged) { isNewBlock = true; blockType = 'transfer'; }
          }

          if (isNewBlock) {
              if (currentBlock) { currentBlock.endDate = startDate; timelineRaw.push(currentBlock); }
              currentBlock = { title, grade, location, startDate, endDate: null, type: blockType };
          } else if (currentBlock) {
              if (location) currentBlock.location = location;
              currentBlock.title = title; 
          }
      });

      if (currentBlock) { currentBlock.endDate = new Date(); timelineRaw.push(currentBlock); }

      let finalTimeline = [];
      for (let i = 0; i < timelineRaw.length; i++) {
          let block = timelineRaw[i];
          if (i < timelineRaw.length - 1 && block.startDate && block.endDate && (block.endDate - block.startDate < 30 * 24 * 60 * 60 * 1000)) continue;
          if (finalTimeline.length > 0) {
              const prevBlock = finalTimeline[finalTimeline.length - 1];
              if (prevBlock.endDate && block.startDate && (block.startDate - prevBlock.endDate > 30 * 24 * 60 * 60 * 1000)) {
                  finalTimeline.push({ isGap: true, duration: formatDuration(prevBlock.endDate, block.startDate) });
              }
          }
          finalTimeline.push(block);
      }

      emp._timeline = finalTimeline.reverse().map(role => role.isGap ? role : { ...role, durationFormatted: formatDuration(role.startDate, role.endDate) });
      emp._timeInRoleDate = emp._timeline.find(r => !r.isGap)?.startDate || null;
      emp._lastPromotionDate = emp._timeline.find(r => r.type === 'promotion')?.startDate || null;
      emp._timeWithManagerDate = emp._tenureDate; 
      empMap[id] = { ...emp };
    });

    let ceos = Object.values(empMap).filter(emp => 
        (!emp['Line Manager UserID'] || String(emp['Line Manager UserID']).trim() === '') &&
        (!emp['Line Manager Name'] || String(emp['Line Manager Name']).trim() === '') &&
        (!emp['Matrix Manager ID'] || String(emp['Matrix Manager ID']).trim() === '') &&
        (!emp['Matrix Manager Name'] || String(emp['Matrix Manager Name']).trim() === '')
    );
    
    let actualCEO = null;
    let computedCeoId = null;
    if (ceos.length > 1) {
      actualCEO = ceos.find(c => {
        const title = String(c['Job Title'] || c['Designation'] || '').toLowerCase();
        return title.includes('ceo') || title.includes('chief') || title.includes('managing director');
      });
      if (!actualCEO) actualCEO = ceos[0];
    } else if (ceos.length === 1) {
      actualCEO = ceos[0];
    }

    if (actualCEO) {
      computedCeoId = actualCEO._id;
      ceos.forEach(c => {
        if (c._id !== actualCEO._id) {
          empMap[c._id]['Matrix Manager Name'] = actualCEO['Display Name'];
          empMap[c._id]['Matrix Manager ID'] = actualCEO['Username'];
        }
      });
    }

    const sysIdMap = {}; const usernameMap = {};
    Object.values(empMap).forEach(emp => {
        const sysId = emp['Users Sys Id']?.toString().trim().toLowerCase();
        const username = emp['Username']?.toString().trim().toLowerCase();
        if (sysId) sysIdMap[sysId] = emp._id;
        if (username) usernameMap[username] = emp._id;
    });

    Object.values(empMap).forEach(emp => {
      const lmIdStr = emp['Line Manager UserID']?.toString().trim().toLowerCase();
      const mmIdStr = emp['Matrix Manager ID']?.toString().trim().toLowerCase();
      const managerId = lmIdStr ? sysIdMap[lmIdStr] : null;
      const matrixId = mmIdStr ? usernameMap[mmIdStr] : null;

      if (managerId && managerId !== emp._id && empMap[managerId]) {
        emp._managerId = managerId;
        if (!directReportsMap[managerId]) directReportsMap[managerId] = [];
        directReportsMap[managerId].push(emp._id);
      }
      if (matrixId && matrixId !== emp._id && empMap[matrixId]) {
        emp._matrixId = matrixId;
        if (!matrixReportsMap[matrixId]) matrixReportsMap[matrixId] = [];
        matrixReportsMap[matrixId].push(emp._id);
      }
    });

    const calculateInsights = (empId, visited = new Set()) => {
      if (visited.has(empId)) return empMap[empId]._insights;
      visited.add(empId);
      const directs = directReportsMap[empId] || [];
      const matrix = matrixReportsMap[empId] || [];
      
      let totalTeam = 0, directCount = 0, eaCount = 0, genderCount = { male: 0, female: 0, other: 0 }, probationCount = 0, noticeCount = 0;
      const directGrades = {}, matrixGrades = {}, teamGrades = {};

      directs.forEach(childId => {
        const child = empMap[childId];
        if (!child) return;
        const childInsights = calculateInsights(childId, visited);
        
        // Include EA in grade summaries and overall team sizes
        const grade = child._summaryGrade;
        directGrades[grade] = (directGrades[grade] || 0) + 1;
        teamGrades[grade] = (teamGrades[grade] || 0) + 1;
        
        const empStatus = String(child['Employee Status (Picklist Label)'] || '').toLowerCase();
        if (empStatus.includes('probation')) probationCount++;
        if (empStatus.includes('notice') || String(child['Resignation Date'] || '').trim().length > 0) noticeCount++;
        
        totalTeam += 1 + (childInsights ? childInsights.totalTeam : 0);
        if (childInsights) Object.entries(childInsights.teamGrades).forEach(([g, c]) => teamGrades[g] = (teamGrades[g] || 0) + c);

        if (!isEA(child)) {
            // Exclude EA from diversity gender count and the core 'directCount' (used for min/max/median SoC)
            const gender = String(child['Gender'] || '').toLowerCase();
            if (gender.startsWith('m')) genderCount.male++; else if (gender.startsWith('f')) genderCount.female++; else genderCount.other++;
            directCount++;
        } else {
            eaCount++;
        }
      });
      matrix.forEach(childId => {
          const child = empMap[childId];
          if(child) matrixGrades[child._summaryGrade] = (matrixGrades[child._summaryGrade] || 0) + 1;
      });

      const insights = { directCount, eaCount, matrixCount: matrix.length, totalTeam, directGrades, matrixGrades, teamGrades, genderCount, probationCount, noticeCount };
      empMap[empId]._insights = insights;
      empMap[empId]._directs = directs;
      empMap[empId]._matrix = matrix;
      return insights;
    };
    Object.keys(empMap).forEach(id => calculateInsights(id));

    // Cap timeWithManagerDate to Manager's own tenure start
    Object.values(empMap).forEach(emp => {
      const managerId = emp._managerId;
      const mgr = empMap[managerId] || empMap[emp._matrixId];
      if (mgr) {
          const mgrJoinDate = mgr._orgTenureDate || mgr._groupTenureDate || mgr._tenureDate;
          if (mgrJoinDate && emp._timeWithManagerDate && emp._timeWithManagerDate < mgrJoinDate) {
              emp._timeWithManagerDate = mgrJoinDate;
          }
      }
    });

    Object.values(empMap).forEach(emp => {
        if (isEA(emp)) return;
        const rank = getRank(emp._summaryGrade);
        const hasTeam = (emp._insights?.totalTeam || 0) > 0;
        if (rank <= 2 && hasTeam) emp._isExCom = true;
    });

    Object.values(empMap).forEach(emp => {
        if (isEA(emp)) return;
        const rank = getRank(emp._summaryGrade);
        const manager = empMap[emp._managerId];
        const matrixManager = empMap[emp._matrixId];
        
        if (rank === 3 || rank === 4) {
            if (manager && manager._isExCom) {
                emp._isOpCom = true;
            } else if (!manager && matrixManager && matrixManager._isExCom) {
                emp._isOpCom = true;
            }
        }
    });

    Object.values(empMap).forEach(emp => {
      const managerId = emp._managerId;
      if (managerId && empMap[managerId]) {
        const manager = empMap[managerId];
        const peers = (manager._directs || []).filter(id => id !== emp._id && !isEA(empMap[id]));
        const managerTeamSize = Math.max(1, manager._insights?.totalTeam || 1);
        const myBranchSize = 1 + (emp._insights?.totalTeam || 0);
        if (managerTeamSize > 0 && !isEA(emp)) {
           emp._insights.pctOfManagerTeam = Math.round((myBranchSize / managerTeamSize) * 100);
           emp._insights.managerValidDrCount = peers.length + 1; 
        }
        if (peers.length > 0) {
          const peerDrs = peers.map(pId => empMap[pId]?._insights?.directCount || 0);
          emp._insights.peerMedianDirects = getMedian(peerDrs);
          emp._insights.peerMinDirects = Math.min(...peerDrs);
          emp._insights.peerMaxDirects = Math.max(...peerDrs);
        }
      }
      
      // Compute formatted duration strings once logic finishes
      emp._tenureFormatted = formatDuration(emp._groupTenureDate || emp._tenureDate);
      emp._orgTenureFormatted = formatDuration(emp._orgTenureDate);
      emp._timeInRoleFormatted = formatDuration(emp._timeInRoleDate || emp._tenureDate);
      emp._lastPromotionFormatted = formatDuration(emp._lastPromotionDate);
      emp._timeWithManagerFormatted = formatDuration(emp._timeWithManagerDate || emp._tenureDate);
    });

    // Globally sort base data by Grade, ExCom, OpCom
    const baseDataArr = Object.values(empMap).sort((a, b) => sortEmployees(a, b, computedCeoId));
    setData(baseDataArr);
    setEmployeeMap(empMap);
    
    if (actualCEO) {
        setActiveEmployeeId(actualCEO._id);
        setCeoId(actualCEO._id);
    }
  };

  const allUniqueGrades = useMemo(() => [...new Set(data.map(emp => emp._summaryGrade).filter(Boolean))].sort((a, b) => getRank(a) - getRank(b)), [data]);
  const allUniqueCompanies = useMemo(() => [...new Set(data.map(emp => emp._derivedCompany).filter(Boolean))].sort((a, b) => a.localeCompare(b)), [data]);

  const filteredSearch = useMemo(() => {
    if (!searchQuery) return [];
    const query = searchQuery.toLowerCase();
    // Using precomputed formatted names handles queries faster
    return data.filter(emp => emp._searchString.includes(query)).slice(0, 5);
  }, [searchQuery, data]);

  // Decoupled filtering logic
  const baseFilteredData = useMemo(() => {
      if (filterConditions.length === 0) return data; 
      return data.filter(emp => {
          const results = filterConditions.map(cond => {
              if (!cond.value && cond.value !== 0 && !['Grade','Company','Committee'].includes(cond.field)) return false;
              if (['Grade','Company','Committee'].includes(cond.field) && cond.value.length === 0) return false;

              let empVal;
              if (cond.field === 'Team Size') empVal = emp._insights?.totalTeam || 0;
              else if (cond.field === 'DR Size') empVal = emp._insights?.directCount || 0;
              else if (cond.field === 'Total Reportees') empVal = (emp._insights?.directCount || 0) + (emp._insights?.matrixCount || 0) + (emp._insights?.eaCount || 0);
              else if (cond.field === 'Grade') empVal = emp._summaryGrade || '';
              else if (cond.field === 'Company') empVal = emp._derivedCompany || '';
              else if (cond.field === 'Committee') empVal = emp._isExCom ? 'ExCom' : (emp._isOpCom ? 'OpCom' : 'None');

              if (['Team Size', 'DR Size', 'Total Reportees'].includes(cond.field)) {
                  const numVal = Number(cond.value);
                  if (isNaN(numVal) || cond.value === '') return false;
                  if (cond.operator === '=') return empVal === numVal;
                  if (cond.operator === '>') return empVal > numVal;
                  if (cond.operator === '<') return empVal < numVal;
              } else if (['Grade', 'Company', 'Committee'].includes(cond.field)) {
                  return cond.value.includes(empVal);
              }
              return false;
          });
          if (results.length === 0) return false;
          return filterMatchMode === 'and' ? results.every(r => r) : results.some(r => r);
      });
  }, [data, filterConditions, filterMatchMode]);

  // Tabular specific sorted view
  const tabularSortedData = useMemo(() => {
      let filtered = [...baseFilteredData];
      if (sortConfigs.length > 0) {
          filtered.sort((a, b) => {
              for (let config of sortConfigs) {
                  let valA, valB;
                  switch (config.field) {
                      case 'Employee': valA = a._formattedName; valB = b._formattedName; break;
                      case 'Grade': valA = getEmpSortRank(a, ceoId); valB = getEmpSortRank(b, ceoId); break;
                      case 'Department': valA = a['Department (Label)'] || ''; valB = b['Department (Label)'] || ''; break;
                      case 'DRSize': valA = a._insights?.directCount || 0; valB = b._insights?.directCount || 0; break;
                      case 'MatrixSize': valA = a._insights?.matrixCount || 0; valB = b._insights?.matrixCount || 0; break;
                      case 'TeamSize': valA = a._insights?.totalTeam || 0; valB = b._insights?.totalTeam || 0; break;
                      case 'Manager': valA = a._formattedManagerName; valB = b._formattedManagerName; break;
                      default: valA = ''; valB = '';
                  }
                  if (valA === valB) continue;
                  let cmp = (typeof valA === 'string' && typeof valB === 'string') ? valA.localeCompare(valB) : (valA > valB ? 1 : -1);
                  return config.dir === 'asc' ? cmp : -cmp;
              }
              return 0;
          });
      }
      return filtered;
  }, [baseFilteredData, sortConfigs, ceoId]);

  const cohortMetrics = useMemo(() => {
      const cohorts = { ExCom: [], OpCom: [], VP: [], GM: [], M3: [] };
      baseFilteredData.forEach(emp => {
          if (emp._isExCom && emp._id !== ceoId) cohorts.ExCom.push(emp);
          if (emp._isOpCom && emp._id !== ceoId) cohorts.OpCom.push(emp);
          if (emp._summaryGrade === 'VP') cohorts.VP.push(emp);
          if (emp._summaryGrade === 'GM') cohorts.GM.push(emp);
          if (emp._summaryGrade === 'M3') cohorts.M3.push(emp);
      });

      const getStats = (arr) => {
          if (arr.length === 0) return { count: 0 };
          const drs = arr.map(a => a._insights?.directCount || 0);
          const matrix = arr.map(a => a._insights?.matrixCount || 0);
          const teams = arr.map(a => a._insights?.totalTeam || 0);
          const totalReps = arr.map(a => (a._insights?.directCount || 0) + (a._insights?.matrixCount || 0));
          
          const nzMatrix = matrix.filter(m => m > 0);

          return {
              count: arr.length,
              drMin: drs.length ? Math.min(...drs) : 0,
              drMax: drs.length ? Math.max(...drs) : 0,
              drMedian: getMedian(drs),

              teamMin: teams.length ? Math.min(...teams) : 0,
              teamMax: teams.length ? Math.max(...teams) : 0,
              teamMedian: getMedian(teams),

              matrixMin: nzMatrix.length ? Math.min(...nzMatrix) : 0,
              matrixMax: nzMatrix.length ? Math.max(...nzMatrix) : 0,
              matrixMedian: getMedian(nzMatrix),
              matrixHasZeros: nzMatrix.length !== matrix.length,
              matrixNzCount: nzMatrix.length,

              totalRepMin: totalReps.length ? Math.min(...totalReps) : 0,
              totalRepMax: totalReps.length ? Math.max(...totalReps) : 0,
              totalRepMedian: getMedian(totalReps)
          };
      };

      return {
          ExCom: getStats(cohorts.ExCom),
          OpCom: getStats(cohorts.OpCom),
          VP: getStats(cohorts.VP),
          GM: getStats(cohorts.GM),
          M3: getStats(cohorts.M3)
      };
  }, [baseFilteredData, ceoId]);

  const dynamicGlobalMetrics = useMemo(() => {
    const grades = {};
    baseFilteredData.forEach(emp => {
        if (!isEA(emp)) {
            const g = emp._summaryGrade;
            if (!grades[g]) grades[g] = { drs: [], teams: [] };
            grades[g].drs.push(emp._insights?.directCount || 0);
            grades[g].teams.push(emp._insights?.totalTeam || 0);
        }
    });

    const gradeMetricsFinal = {};
    Object.keys(grades).forEach(g => {
        gradeMetricsFinal[g] = {
            drMin: Math.min(...grades[g].drs), drMax: Math.max(...grades[g].drs), drMedian: getMedian(grades[g].drs),
            teamMin: Math.min(...grades[g].teams), teamMax: Math.max(...grades[g].teams), teamMedian: getMedian(grades[g].teams)
        };
    });

    return {
        grade: gradeMetricsFinal,
        exCom: cohortMetrics.ExCom,
        opCom: cohortMetrics.OpCom
    };
  }, [baseFilteredData, cohortMetrics]);

  const heatmapStats = useMemo(() => {
      const depts = {};
      baseFilteredData.forEach(emp => {
          if (emp._insights?.directCount > 0) {
              const d = emp['Department (Label)'] || 'Unknown';
              if (!depts[d]) depts[d] = [];
              depts[d].push(emp._insights.directCount);
          }
      });
      return Object.entries(depts).map(([d, drs]) => ({ dept: d, medianDr: getMedian(drs), count: drs.length }))
            .sort((a,b) => b.medianDr - a.medianDr).slice(0, 10);
  }, [baseFilteredData]);

  const filterFieldsOrder = ['Company', 'Committee', 'Grade', 'DR Size', 'Total Reportees', 'Team Size'];

  const addFilterCondition = () => {
      const nextField = filterFieldsOrder[filterConditions.length % filterFieldsOrder.length];
      
      let defaultOp = '=', defaultVal = '';
      if (['Team Size', 'DR Size', 'Total Reportees'].includes(nextField)) { defaultOp = '='; defaultVal = ''; }
      if (['Grade', 'Company', 'Committee'].includes(nextField)) { defaultOp = 'in'; defaultVal = []; }

      setFilterConditions([...filterConditions, { id: Date.now(), field: nextField, operator: defaultOp, value: defaultVal }]);
      setAppTab('table');
      if (!isSidebarOpen) setIsSidebarOpen(true);
  };
  
  const updateFilterCondition = (id, key, val) => {
      setFilterConditions(filterConditions.map(c => c.id === id ? { ...c, [key]: val } : c));
      if (key === 'field') {
          setFilterConditions(prev => prev.map(c => {
              if (c.id === id) {
                  let defaultOp = '=', defaultVal = '';
                  if (['Team Size', 'DR Size', 'Total Reportees'].includes(val)) { defaultOp = '='; defaultVal = ''; }
                  if (['Grade', 'Company', 'Committee'].includes(val)) { defaultOp = 'in'; defaultVal = []; }
                  return { ...c, operator: defaultOp, value: defaultVal };
              }
              return c;
          }));
      }
      setAppTab('table');
  };

  const handleSummaryTileClick = (type) => {
      let newFilters = [...filterConditions];
      if (['VP', 'GM', 'M3'].includes(type)) {
          newFilters = newFilters.filter(f => f.field !== 'Grade');
          newFilters.push({ id: Date.now(), field: 'Grade', operator: 'in', value: [type] });
      } else if (['ExCom', 'OpCom'].includes(type)) {
          newFilters = newFilters.filter(f => f.field !== 'Committee');
          newFilters.push({ id: Date.now(), field: 'Committee', operator: 'in', value: [type] });
      }
      setFilterConditions(newFilters);
      setActiveCohortScale(type);
      setAppTab('table');
  };

  const removeFilterCondition = (id) => {
      setFilterConditions(filterConditions.filter(c => c.id !== id));
  };

  const handleSort = (field) => {
      setSortConfigs(prev => {
          const existingIdx = prev.findIndex(c => c.field === field);
          if (existingIdx === -1) return [...prev, { field, dir: 'asc' }];
          const newConfigs = [...prev];
          if (newConfigs[existingIdx].dir === 'asc') newConfigs[existingIdx] = { field, dir: 'desc' };
          else newConfigs.splice(existingIdx, 1);
          return newConfigs;
      });
  };

  const handleAddToCompare = (empId, color) => {
      setCompareList(prev => {
          const group = prev[color] || [];
          if (group.length >= 4 && !group.includes(empId)) { alert('Maximum 4 employees per group.'); return prev; }
          if (!group.includes(empId)) return { ...prev, [color]: [...group, empId] };
          return prev;
      });
      setContextMenu(null);
  };

  const activeEmployee = employeeMap[activeEmployeeId];
  const manager = activeEmployee?._managerId ? employeeMap[activeEmployee._managerId] : null;
  const directReports = (activeEmployee?._directs || []).map(id => employeeMap[id]).filter(Boolean).filter(emp => filterConditions.length === 0 || baseFilteredData.find(f => f._id === emp._id)).sort((a, b) => sortEmployees(a, b, ceoId));
  const matrixReports = (activeEmployee?._matrix || []).map(id => employeeMap[id]).filter(Boolean).filter(emp => filterConditions.length === 0 || baseFilteredData.find(f => f._id === emp._id)).sort((a, b) => sortEmployees(a, b, ceoId));
  const isMatrixView = viewMode === 'matrix';
  const displayedReports = isMatrixView ? matrixReports : directReports;

  // Reset table scroll if filtering changes
  useEffect(() => {
      if (tableContainerRef.current) {
          tableContainerRef.current.scrollTop = 0;
      }
  }, [tabularSortedData]);

  // --- RENDER ---
  if (data.length === 0) {
    return (
      <div className="h-screen w-full bg-slate-100 flex items-center justify-center p-6">
        <div className={`max-w-xl w-full bg-white p-10 rounded-2xl shadow-xl border-2 border-dashed transition-colors ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-slate-300'}`} onDragOver={handleDragOver} onDragLeave={handleDragLeave} onDrop={handleDrop}>
          <div className="flex flex-col items-center text-center space-y-4">
            <div className="w-20 h-20 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center"><Upload size={40} /></div>
            <h2 className="text-2xl font-bold text-slate-800">Upload Employee Data</h2>
            <p className="text-slate-500 text-sm">Drag and drop your Excel (.xlsx) file here.<br/>Ensure it contains standard employee details.</p>
            <input type="file" accept=".xlsx, .xls" className="hidden" id="file-upload" disabled={!xlsxLoaded} onChange={(e) => e.target.files[0] && handleFileUpload(e.target.files[0])} />
            <label htmlFor="file-upload" className={`px-6 py-3 text-white font-medium rounded-lg transition-colors shadow-md ${!xlsxLoaded || loading ? 'bg-slate-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700 cursor-pointer'}`}>
              {!xlsxLoaded ? 'Loading Library...' : loading ? 'Processing...' : 'Select Excel File'}
            </label>
            {error && <p className="text-red-500 text-sm mt-4 font-medium">{error}</p>}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen w-full flex flex-col font-sans text-slate-800 bg-slate-100 overflow-hidden">
      
      {/* Dynamic Print CSS Injection */}
      <style dangerouslySetInnerHTML={{__html: `
        @media print {
            @page { size: landscape; margin: 8mm; }
            html, body { background-color: white !important; margin: 0; padding: 0; }
            body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        }
      `}} />

      {contextMenu && createPortal(
          <div className="fixed bg-white border border-slate-200 shadow-xl rounded-xl p-3 z-[999999] animate-scale-in context-menu" style={{ top: contextMenu.y, left: contextMenu.x }}>
              <div className="text-xs font-bold mb-3 text-slate-500 uppercase tracking-wider">Add to Compare</div>
              <div className="flex gap-2.5 mb-3">
                  {[ {id:'blue', bg:'bg-blue-500'}, {id:'green', bg:'bg-green-500'}, {id:'red', bg:'bg-red-500'}, {id:'orange', bg:'bg-orange-500'}, {id:'purple', bg:'bg-purple-500'} ].map(c => (
                      <button key={c.id} onClick={() => handleAddToCompare(contextMenu.empId, c.id)} className={`w-6 h-6 rounded-md ${c.bg} shadow-sm hover:scale-110 transition-transform hover:ring-2 hover:ring-offset-1 hover:ring-${c.id}-400`}></button>
                  ))}
              </div>
              <div className="w-full h-px bg-slate-100 my-2"></div>
              <button 
                  onClick={() => { setPrintNodeId(contextMenu.empId); setContextMenu(null); }}
                  className="w-full flex items-center justify-center gap-2 text-[11px] font-bold text-slate-600 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 py-1.5 rounded transition-colors"
              >
                  <Printer size={12}/> Print Structure
              </button>
          </div>,
          document.body
      )}

      {/* MAIN APPLICATION (Hidden during print) */}
      <div className={`flex-col h-screen w-full overflow-hidden ${printNodeId ? 'hidden' : 'flex'} print:hidden`}>
          <header className="bg-white border-b px-6 py-4 flex items-center justify-between shadow-sm z-30 flex-shrink-0">
            <div className="flex items-center space-x-3 w-1/3">
              <div className="bg-blue-600 p-2 rounded-lg"><Users className="text-white" size={24} /></div>
              <h1 className="text-2xl font-serif italic text-blue-800 tracking-tight hidden sm:block font-bold">Org Sense</h1>
            </div>
            <div className="flex bg-slate-50 p-1 rounded-lg border border-slate-200 w-fit mx-auto justify-center">
                <button onClick={() => setAppTab('org')} className={`px-4 py-1.5 rounded-md text-sm font-semibold transition-all flex items-center gap-1.5 ${appTab === 'org' ? 'bg-white text-blue-600 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}>Structure</button>
                <button onClick={() => setAppTab('table')} className={`px-4 py-1.5 rounded-md text-sm font-semibold transition-all flex items-center gap-1.5 ${appTab === 'table' ? 'bg-white text-blue-600 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}>Table</button>
                <button onClick={() => setAppTab('compare')} className={`px-4 py-1.5 rounded-md text-sm font-semibold transition-all flex items-center gap-1.5 ${appTab === 'compare' ? 'bg-white text-blue-600 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}><BarChart2 size={14} /> Compare</button>
            </div>
            <div className="flex items-center justify-end space-x-4 w-1/3">
              {(appTab === 'org' || appTab === 'table') && (
                  <>
                      <div className="relative w-64 hidden md:block" ref={searchRef}>
                        <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-slate-400" size={18} />
                        <input 
                            type="text" 
                            placeholder="Search employee..." 
                            className="w-full pl-10 pr-4 py-1.5 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 bg-slate-50 text-sm" 
                            value={searchQuery} 
                            onChange={(e) => { setSearchQuery(e.target.value); setIsSearchOpen(true); }}
                            onFocus={() => setIsSearchOpen(true)} 
                        />
                        {isSearchOpen && searchQuery && (
                          <div className="absolute top-full right-0 mt-2 w-80 bg-white rounded-lg shadow-xl border overflow-hidden z-50">
                            {filteredSearch.length > 0 ? filteredSearch.map(emp => (
                                <button key={emp._id} className="w-full text-left px-4 py-3 hover:bg-slate-50 border-b last:border-0 flex flex-col" onClick={() => { handleEmployeeSelect(emp._id); setSearchQuery(''); setIsSearchOpen(false); }}>
                                  <span className="font-semibold text-slate-800">{emp._formattedName}</span>
                                  <span className="text-xs text-slate-500">{emp['Job Title']} • {emp['Department (Label)']}</span>
                                </button>
                              )) : <div className="px-4 py-3 text-slate-500 text-sm">No employees found.</div>}
                          </div>
                        )}
                      </div>
                      {ceoId && filterConditions.length === 0 && (
                        <button onClick={() => handleEmployeeSelect(ceoId)} className="px-4 py-1.5 bg-slate-50 hover:bg-slate-200 text-slate-700 rounded-lg font-medium transition-colors text-sm border border-slate-200 whitespace-nowrap">Go to Top</button>
                      )}
                  </>
              )}
            </div>
          </header>

          <main className="flex-1 overflow-hidden flex flex-row w-full relative">
            
            {/* LEFT SIDEBAR (DASHBOARD) */}
            {(appTab === 'org' || appTab === 'table') && (
                <aside className={`${isSidebarOpen ? 'w-72 md:w-80' : 'w-12'} bg-white border-r border-slate-200 flex-shrink-0 flex flex-col relative transition-all duration-300 z-10 shadow-[2px_0_10px_rgba(0,0,0,0.02)] hidden sm:flex`} >
                    
                    <button 
                        onClick={() => setIsSidebarOpen(!isSidebarOpen)}
                        className="absolute -right-3.5 top-6 bg-slate-50 border border-slate-300 shadow-md rounded-full p-1.5 z-[60] text-slate-500 hover:text-blue-600 focus:outline-none"
                    >
                        {isSidebarOpen ? <ChevronLeft size={16}/> : <ChevronRight size={16}/>}
                    </button>

                    <div className="flex-1 overflow-hidden relative">
                      {isSidebarOpen ? (
                        <div className="h-full overflow-y-auto flex flex-col" style={{ scrollbarWidth: 'thin' }}>
                            
                            {/* Filters Section (Collapsible) */}
                            <div className="p-5 border-b border-slate-100 bg-white filter-dropdown-wrapper">
                                <button 
                                    onClick={() => setShowFilterPanel(!showFilterPanel)} 
                                    className="flex justify-between items-center w-full text-left focus:outline-none"
                                >
                                    <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-2">
                                        <Filter size={14} /> Filters {filterConditions.length > 0 && `(${filterConditions.length})`}
                                    </h3>
                                    {showFilterPanel ? <ChevronDown size={14} className="text-slate-400"/> : <ChevronRight size={14} className="text-slate-400"/>}
                                </button>

                                {showFilterPanel && (
                                    <div className="mt-4 space-y-3 animate-fade-in-down">
                                        <div className="flex bg-slate-100 rounded p-0.5 border border-slate-200">
                                            <button onClick={() => setFilterMatchMode('and')} className={`flex-1 px-2 py-1 text-[10px] uppercase font-bold rounded transition-colors ${filterMatchMode === 'and' ? 'bg-white shadow-sm text-blue-600' : 'text-slate-500 hover:text-slate-700'}`}>Match All</button>
                                            <button onClick={() => setFilterMatchMode('or')} className={`flex-1 px-2 py-1 text-[10px] uppercase font-bold rounded transition-colors ${filterMatchMode === 'or' ? 'bg-white shadow-sm text-purple-600' : 'text-slate-500 hover:text-slate-700'}`}>Match Any</button>
                                        </div>

                                        <div className="space-y-1.5">
                                            {filterConditions.map((cond, index) => (
                                                <div key={cond.id} className="flex items-center gap-1.5 p-1 rounded hover:bg-slate-50 border border-transparent hover:border-slate-200 group transition-colors">
                                                    
                                                    <select className="bg-transparent text-xs font-bold text-slate-700 focus:outline-none cursor-pointer p-0 border-none w-[95px] flex-shrink-0 appearance-none" value={cond.field} onChange={(e) => updateFilterCondition(cond.id, 'field', e.target.value)}>
                                                        <option value="Company">Company</option>
                                                        <option value="Committee">Leadership</option>
                                                        <option value="Grade">Grade</option>
                                                        <option value="DR Size">Direct Reports Count</option>
                                                        <option value="Total Reportees">Total Reportees</option>
                                                        <option value="Team Size">Team Size</option>
                                                    </select>

                                                    {['Team Size', 'DR Size', 'Total Reportees'].includes(cond.field) && (
                                                        <select className="bg-white border border-slate-200 rounded px-1 text-xs font-medium text-slate-700 focus:outline-none h-6" value={cond.operator} onChange={(e) => updateFilterCondition(cond.id, 'operator', e.target.value)}>
                                                            <option value="=">=</option>
                                                            <option value=">">&gt;</option>
                                                            <option value="<">&lt;</option>
                                                        </select>
                                                    )}

                                                    <div className="flex-1 min-w-0 relative">
                                                        {['Team Size', 'DR Size', 'Total Reportees'].includes(cond.field) ? (
                                                            <input type="number" className="w-full border border-slate-200 rounded px-2 text-xs font-medium text-slate-700 focus:outline-none focus:border-blue-400 h-6" placeholder="0" value={cond.value} onChange={(e) => updateFilterCondition(cond.id, 'value', e.target.value)} />
                                                        ) : (
                                                            <div className="relative">
                                                                <button onClick={() => setOpenDropdown(openDropdown === cond.id ? null : cond.id)} className="w-full border border-slate-200 rounded px-2 text-xs font-medium bg-white text-left flex justify-between items-center focus:border-blue-400 h-6 truncate">
                                                                    <span className="truncate text-slate-700">{cond.value.length > 0 ? `${cond.value.length} Selected` : `Select...`}</span><ChevronDown size={12} className="text-slate-400 flex-shrink-0 ml-1" />
                                                                </button>
                                                                {openDropdown === cond.id && (
                                                                    <div className="absolute top-full left-0 mt-1 w-full max-h-48 overflow-y-auto bg-white border border-slate-200 shadow-xl rounded-md z-50 p-1 flex flex-col" style={{ scrollbarWidth: 'thin' }}>
                                                                        {(cond.field === 'Grade' ? allUniqueGrades : (cond.field === 'Company' ? allUniqueCompanies : ['ExCom', 'OpCom'])).map(item => (
                                                                            <label key={item} className="flex items-center gap-2 text-xs p-1.5 hover:bg-slate-50 rounded cursor-pointer border border-transparent transition-colors">
                                                                                <input type="checkbox" className="rounded text-blue-600 focus:ring-blue-500 w-3 h-3 m-0" checked={cond.value.includes(item)} onChange={(e) => {
                                                                                        const newVals = e.target.checked ? [...cond.value, item] : cond.value.filter(v => v !== item);
                                                                                        updateFilterCondition(cond.id, 'value', newVals);
                                                                                    }} 
                                                                                />
                                                                                <span className="truncate text-slate-700" title={item}>{item}</span>
                                                                            </label>
                                                                        ))}
                                                                    </div>
                                                                )}
                                                            </div>
                                                        )}
                                                    </div>
                                                    <button onClick={() => removeFilterCondition(cond.id)} className="text-slate-300 hover:text-red-500 p-1 rounded transition-colors opacity-0 group-hover:opacity-100 flex-shrink-0"><X size={14} /></button>
                                                </div>
                                            ))}
                                        </div>

                                        <button onClick={addFilterCondition} className="w-full flex justify-center items-center text-xs font-semibold text-blue-600 hover:text-blue-700 bg-blue-50 hover:bg-blue-100 py-1.5 rounded transition-colors border border-blue-100">
                                            <Plus size={14} className="mr-1" /> Add Rule
                                        </button>
                                    </div>
                                )}
                            </div>

                            {/* Cohort Summaries */}
                            <div className="p-5">
                                <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-4">Cohort Summaries</h3>

                                <div className="flex flex-col gap-3">
                                    {['ExCom', 'OpCom', 'VP', 'GM', 'M3'].map(k => {
                                        const s = cohortMetrics[k];
                                        if (!s || s.count === 0) return null;
                                        
                                        let bgClass = 'bg-white border-slate-200 hover:border-slate-400';
                                        let titleClass = 'text-slate-700';
                                        let Icon = null;
                                        
                                        if (k === 'ExCom') { bgClass = 'bg-amber-50 border-amber-200 hover:border-amber-400'; titleClass = 'text-amber-800'; }
                                        else if (k === 'OpCom') { bgClass = 'bg-indigo-50 border-indigo-200 hover:border-indigo-400'; titleClass = 'text-indigo-800'; }
                                        else { Icon = Award; }

                                        return (
                                            <button key={k} onClick={() => handleSummaryTileClick(k)} className={`rounded-xl p-3 shadow-sm transition-all text-left border group ${bgClass} ${activeCohortScale === k ? 'ring-2 ring-blue-500 ring-offset-1' : ''}`}>
                                                <div className="flex justify-between items-end mb-2.5 border-b border-slate-200/50 pb-2">
                                                    <span className={`font-bold text-sm flex items-center gap-1.5 ${titleClass}`}>{Icon && <Icon size={14}/>}{k}</span>
                                                    <span className="text-xs text-slate-500 bg-white/60 px-1.5 py-0.5 rounded font-bold">{s.count}</span>
                                                </div>
                                                <div className="grid grid-cols-3 gap-1 text-center divide-x divide-slate-200/50">
                                                    <div title="Median"><div className="text-[9px] text-slate-500 mb-0.5 font-bold uppercase">Direct</div><div className="font-bold text-blue-600">{formatNum(s.drMedian)}</div></div>
                                                    <div title={s.matrixHasZeros ? `Median for ${s.matrixNzCount} employees` : "Median"}><div className="text-[9px] text-slate-500 mb-0.5 font-bold uppercase">Matrix</div><div className="font-bold text-purple-600">{formatNum(s.matrixMedian)}{s.matrixHasZeros && s.matrixNzCount > 0 ? '*' : ''}</div></div>
                                                    <div title="Median"><div className="text-[9px] text-slate-500 mb-0.5 font-bold uppercase">Total</div><div className="font-bold text-orange-600">{formatNum(s.totalRepMedian)}</div></div>
                                                </div>
                                            </button>
                                        );
                                    })}
                                </div>

                                {/* Active Cohort Scales & Heatmap Collapse */}
                                {activeCohortScale && cohortMetrics[activeCohortScale] && cohortMetrics[activeCohortScale].count > 0 && (
                                   <div className="mt-6 pt-5 border-t border-slate-200 animate-fade-in-down">
                                      <div className="flex justify-between items-center mb-4">
                                          <h4 className="text-sm font-bold text-slate-700 flex items-center gap-2">
                                              {activeCohortScale} Benchmark <span className="text-xs font-bold bg-slate-100 text-slate-600 px-2 py-0.5 rounded-full">{cohortMetrics[activeCohortScale].count}</span>
                                      </h4>
                                      <button onClick={() => setActiveCohortScale(null)} className="text-slate-400 hover:text-slate-600 bg-slate-50 border border-slate-100 shadow-sm p-1 rounded-md"><X size={16}/></button>
                                  </div>
                                  
                                  <div className="flex flex-col gap-2 mb-6">
                                      <MetricScale label="Direct Reports" min={cohortMetrics[activeCohortScale].drMin} max={cohortMetrics[activeCohortScale].drMax} median={cohortMetrics[activeCohortScale].drMedian} value={0} hideCurrent />
                                      <div className="relative">
                                          <MetricScale label="Matrix Reports" min={cohortMetrics[activeCohortScale].matrixMin} max={cohortMetrics[activeCohortScale].matrixMax} median={cohortMetrics[activeCohortScale].matrixMedian} value={0} hideCurrent />
                                          {cohortMetrics[activeCohortScale].matrixHasZeros && cohortMetrics[activeCohortScale].matrixNzCount > 0 && (
                                              <p className="text-[9px] text-slate-400 italic absolute -bottom-2">* {cohortMetrics[activeCohortScale].matrixNzCount} employees in this cohort have matrix reports</p>
                                          )}
                                      </div>
                                      <MetricScale label="Total Reportees" min={cohortMetrics[activeCohortScale].totalRepMin} max={cohortMetrics[activeCohortScale].totalRepMax} median={cohortMetrics[activeCohortScale].totalRepMedian} value={0} hideCurrent />
                                      <MetricScale label="Team Size" min={cohortMetrics[activeCohortScale].teamMin} max={cohortMetrics[activeCohortScale].teamMax} median={cohortMetrics[activeCohortScale].teamMedian} value={0} hideCurrent />
                                  </div>

                                  {/* Heatmap */}
                                  {heatmapStats.length > 0 && (
                                      <div className="mt-6 pt-5 border-t border-slate-100">
                                          <h3 className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-4 flex items-center"><BarChart2 size={14} className="mr-1.5"/> Median Span of Control</h3>
                                          <div className="flex flex-col gap-2">
                                              {heatmapStats.map(hs => {
                                                  const maxVal = Math.max(...heatmapStats.map(d=>d.medianDr));
                                                  const intensity = maxVal > 0 ? (hs.medianDr / maxVal) : 0;
                                                  let colorClass = 'bg-slate-50 border-slate-200 text-slate-700';
                                                  if (intensity > 0.7) colorClass = 'bg-blue-500 border-blue-600 text-white';
                                                  else if (intensity > 0.3) colorClass = 'bg-blue-100 border-blue-200 text-blue-900';

                                                  return (
                                                      <div key={hs.dept} className={`border rounded-lg px-3 py-2 text-xs font-semibold flex items-center justify-between shadow-sm ${colorClass}`}>
                                                          <span className="truncate pr-2">{hs.dept}</span>
                                                          <span className="bg-white/40 px-2 py-0.5 rounded shadow-sm text-sm">{formatNum(hs.medianDr)}</span>
                                                      </div>
                                                  );
                                              })}
                                          </div>
                                      </div>
                                  )}
                               </div>
                            )}
                        </div>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center h-full pt-10 text-slate-400 font-bold uppercase tracking-widest text-xs">
                        <span className="rotate-90 whitespace-nowrap mt-16">Dashboard</span>
                    </div>
                  )}
                </div>
            </aside>
        )}

        {/* RIGHT CONTENT AREA */}
        <div className="flex-1 flex flex-col relative bg-slate-50 min-h-0 overflow-hidden" id="chart-container">
            
            {/* Active Filters Header (Sticky) */}
            {((appTab === 'org' || appTab === 'table') && filterConditions.length > 0) && (
                <div className="w-full bg-slate-50/90 backdrop-blur-md border-b border-slate-200 px-4 py-2.5 sm:px-8 flex flex-wrap items-center gap-2 shadow-sm min-h-[44px]">
                    <span className="text-[11px] font-bold text-slate-500 uppercase tracking-wider mr-1">
                        {filterMatchMode === 'or' ? 'Matches any:' : 'Matches all:'}
                    </span>
                    {filterConditions.flatMap(cond => {
                        const pills = [];
                        let displayField = cond.field;
                        if (cond.field === 'Committee') displayField = 'Leadership';
                        if (cond.field === 'DR Size') displayField = 'Directs';

                        if (['Grade', 'Company', 'Committee'].includes(cond.field)) {
                            if (Array.isArray(cond.value)) {
                                cond.value.forEach(val => {
                                    pills.push({ condId: cond.id, type: 'array', val: val, display: val });
                                });
                            }
                        } else if (['Team Size', 'DR Size', 'Total Reportees'].includes(cond.field)) {
                            pills.push({ condId: cond.id, type: 'single', display: `${displayField} ${cond.operator} ${cond.value}` });
                        }
                        return pills;
                    }).map((pill, i) => (
                        <div key={`${pill.condId}-${i}`} className={`flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-bold border shadow-sm transition-colors ${filterMatchMode === 'or' ? 'bg-purple-100 text-purple-800 border-purple-200 hover:bg-purple-200' : 'bg-blue-100 text-blue-800 border-blue-200 hover:bg-blue-200'}`}>
                            <span>{pill.display}</span>
                            <button onClick={() => {
                                if (pill.type === 'single') {
                                    removeFilterCondition(pill.condId);
                                } else {
                                    const cond = filterConditions.find(c => c.id === pill.condId);
                                    if (cond) {
                                        const newVals = cond.value.filter(v => v !== pill.val);
                                        if (newVals.length === 0) removeFilterCondition(cond.id);
                                        else updateFilterCondition(cond.id, 'value', newVals);
                                    }
                                }
                            }} className="opacity-50 hover:opacity-100 bg-white/50 rounded-full p-0.5"><X size={10}/></button>
                        </div>
                    ))}
                    <button onClick={() => { setFilterConditions([]); setActiveCohortScale(null); }} className="text-[10px] font-bold text-slate-400 hover:text-red-600 uppercase tracking-wider ml-auto flex items-center gap-1 transition-colors"><Trash2 size={12}/> Clear All</button>
                </div>
            )}

            {/* TABULAR VIEW CONTAINER */}
            <div className={`bg-white m-4 md:m-8 rounded-xl shadow-sm border border-slate-200 flex-1 flex-col overflow-hidden min-h-0 animate-fade-in-up ${appTab === 'table' ? 'flex' : 'hidden'}`}>
                <div className="p-4 border-b border-slate-200 flex items-center justify-between bg-slate-50">
                    <h2 className="text-lg font-bold text-slate-800 flex items-center gap-2">Filtered Results <span className="text-xs font-medium text-slate-500 bg-white border border-slate-200 px-2 py-0.5 rounded-full">{tabularSortedData.length} records</span></h2>
                </div>
                
                {tabularSortedData.length === 0 ? (
                    <div className="p-10 text-center text-slate-500">No employees match your current filter conditions.</div>
                ) : (
                    <div className="flex-1 overflow-auto" style={{ scrollbarWidth: 'thin' }} ref={tableContainerRef}>
                        <table className="w-full text-left text-sm min-w-[900px]">
                            <thead className="text-slate-600 border-b border-slate-200 sticky top-0 z-10 bg-slate-50 shadow-sm">
                                <tr>
                                    <SortableHeader label="Employee" field="Employee" sortConfigs={sortConfigs} handleSort={handleSort} />
                                    <SortableHeader label="Grade" field="Grade" sortConfigs={sortConfigs} handleSort={handleSort} />
                                    <SortableHeader label="Job Title" field="JobTitle" sortConfigs={sortConfigs} handleSort={handleSort} />
                                    <SortableHeader label="Department" field="Department" sortConfigs={sortConfigs} handleSort={handleSort} />
                                    <SortableHeader label="DR" field="DRSize" align="center" sortConfigs={sortConfigs} handleSort={handleSort} />
                                    <SortableHeader label="Mat" field="MatrixSize" align="center" sortConfigs={sortConfigs} handleSort={handleSort} />
                                    <SortableHeader label="Team" field="TeamSize" align="center" sortConfigs={sortConfigs} handleSort={handleSort} />
                                    <SortableHeader label="Line Manager" field="Manager" sortConfigs={sortConfigs} handleSort={handleSort} />
                                </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                                {tabularSortedData.map((emp) => (
                                    <tr key={emp._id} id={`table-row-${emp._id}`} className="hover:bg-blue-50/50 bg-white cursor-pointer transition-colors" onClick={() => handleEmployeeSelect(emp._id)}>
                                        <td className="px-4 py-3"><div className="font-bold text-slate-800 flex items-center gap-1.5"><span className="truncate max-w-[200px]">{emp._formattedName}</span>{emp._isExCom && <span className="text-[8px] bg-amber-100 text-amber-700 px-1 rounded uppercase font-bold flex-shrink-0">Ex</span>}{emp._isOpCom && <span className="text-[8px] bg-indigo-100 text-indigo-700 px-1 rounded uppercase font-bold flex-shrink-0">Op</span>}</div></td>
                                        <td className="px-4 py-3">{emp._cardGrade && emp._cardGrade !== 'Unspecified' ? <span className="bg-slate-100 text-slate-600 px-1.5 py-0.5 rounded text-[10px] font-bold border border-slate-200">{emp._cardGrade}</span> : <span className="text-slate-400">-</span>}</td>
                                        <td className="px-4 py-3 text-slate-700"><div className="truncate max-w-[200px]" title={emp['Job Title']}>{emp['Job Title']}</div></td>
                                        <td className="px-4 py-3 text-slate-600"><div className="truncate max-w-[150px]" title={emp['Department (Label)']}>{emp['Department (Label)'] || ''}</div></td>
                                        <td className="px-4 py-3 text-center font-medium text-blue-700">{formatNum(emp._insights?.directCount)}</td>
                                        <td className="px-4 py-3 text-center font-medium text-purple-600">{formatNum(emp._insights?.matrixCount)}</td>
                                        <td className="px-4 py-3 text-center font-medium text-orange-600">{formatNum(emp._insights?.totalTeam)}</td>
                                        <td className="px-4 py-3 text-slate-600"><div className="truncate max-w-[150px]" title={emp._formattedManagerName}>{emp._formattedManagerName || '-'}</div></td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}
            </div>

            {/* ORG CHART VIEW */}
            <div className={`w-full mx-auto flex-col items-center pb-32 p-4 sm:p-8 overflow-y-auto ${appTab === 'org' ? 'flex' : 'hidden'}`}>
                {manager && (
                <div className="flex flex-col items-center animate-fade-in-down w-full">
                    <EmployeeCard employee={manager} ceoId={ceoId} globalMetrics={dynamicGlobalMetrics} onClick={() => handleEmployeeSelect(manager._id)} onSelectDirect={() => { setActiveEmployeeId(manager._id); setViewMode('direct'); }} onSelectMatrix={() => { setActiveEmployeeId(manager._id); setViewMode('matrix'); }} onContextMenu={(e, emp) => { e.preventDefault(); setContextMenu({x: e.clientX, y: e.clientY, empId: emp._id}); }} />
                    <div className="h-10 w-px bg-slate-300 my-2"></div>
                </div>
                )}

                {activeEmployee && (
                <div className="relative flex justify-center items-center my-4 animate-scale-in z-10 w-full max-w-sm">
                    <EmployeeCard employee={activeEmployee} ceoId={ceoId} globalMetrics={dynamicGlobalMetrics} isActive viewMode={viewMode} onSelectDirect={() => setViewMode('direct')} onSelectMatrix={() => setViewMode('matrix')} onContextMenu={(e, emp) => { e.preventDefault(); setContextMenu({x: e.clientX, y: e.clientY, empId: emp._id}); }} />
                </div>
                )}

                {/* TEAMS STYLE MULTI-LINE LAYOUT */}
                {(() => {
                    const totalUnfilteredReports = isMatrixView ? (activeEmployee?._matrix || []).length : (activeEmployee?._directs || []).length;
                    const hasFilteredReports = filterConditions.length > 0 && totalUnfilteredReports > displayedReports.length;
                    const isCompletelyFiltered = totalUnfilteredReports > 0 && displayedReports.length === 0;

                    if (totalUnfilteredReports === 0) return null;

                    let pillClasses = `text-[10px] font-bold uppercase tracking-wider px-4 py-1.5 rounded-full shadow-sm border flex items-center gap-2 `;
                    
                    if (hasFilteredReports || isCompletelyFiltered) {
                        pillClasses += `bg-slate-100 text-slate-500 border-slate-200`;
                    } else {
                        pillClasses += isMatrixView ? 'bg-purple-50 text-purple-700 border-purple-200' : 'bg-white text-slate-600 border-slate-200';
                    }

                    return (
                        <div className="flex flex-col items-center animate-fade-in-up w-full mt-2">
                            <div className={`h-6 w-px ${isMatrixView ? 'bg-purple-400' : 'bg-slate-300'}`}></div>
                            
                            <div className="flex flex-col items-center gap-1.5 mb-6">
                                <div className={pillClasses}>
                                   <span>{isMatrixView ? 'Matrix Reports' : 'Direct Reports'} ({displayedReports.length}{(hasFilteredReports || isCompletelyFiltered) ? ` / ${totalUnfilteredReports}` : ''})</span>
                                   
                                   {(hasFilteredReports || isCompletelyFiltered) && (
                                       <>
                                           <div className="w-px h-3 bg-slate-300"></div>
                                           <span className="text-slate-400 flex items-center gap-1"><Filter size={10}/> Filters Applied</span>
                                       </>
                                   )}
                                </div>
                            </div>
                            
                            {displayedReports.length > 0 ? (
                                <div className="flex justify-center flex-wrap gap-6 w-full px-4">
                                {displayedReports.map(emp => (
                                    <div key={emp._id} className="flex flex-col items-center relative w-full sm:w-auto">
                                    <EmployeeCard employee={emp} ceoId={ceoId} globalMetrics={dynamicGlobalMetrics} isMatrixNode={isMatrixView} onClick={() => handleEmployeeSelect(emp._id)} onSelectDirect={() => { setActiveEmployeeId(emp._id); setViewMode('direct'); }} onSelectMatrix={() => { setActiveEmployeeId(emp._id); setViewMode('matrix'); }} onContextMenu={(e, emp) => { e.preventDefault(); setContextMenu({x: e.clientX, y: e.clientY, empId: emp._id}); }} />
                                    </div>
                                ))}
                                </div>
                            ) : (
                                <div className="text-sm text-slate-400 italic bg-white px-6 py-4 rounded-xl border border-slate-200 shadow-sm mt-2">
                                    All {isMatrixView ? 'matrix' : 'direct'} reports for this employee have been hidden by the current filters.
                                </div>
                            )}
                        </div>
                    );
                })()}
            </div>

            {/* COMPARE VIEW */}
            <div className={`w-full h-full flex-col overflow-hidden p-0 bg-slate-50 min-h-0 ${appTab === 'compare' ? 'flex' : 'hidden'}`}>
                <CompareView compareList={compareList} employeeMap={employeeMap} ceoId={ceoId} />
            </div>
        </div>
      </main>
      </div>

      {/* PRINT LAYOUT (Only visible during print) */}
      {printNodeId && (
        <div className="hidden print:block w-full">
            <PrintLayout rootId={printNodeId} employeeMap={employeeMap} ceoId={ceoId} />
        </div>
      )}
    </div>
  );
}

export default App;