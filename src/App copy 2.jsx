import React, { useState, useMemo, useEffect, useRef } from 'react';
import { Upload, Search, Info, Users, MapPin, Building2, UserCircle2, Clock, CalendarDays, Award, ChevronDown, ChevronRight, UserMinus, TrendingDown, X, Filter, Plus, Trash2, ArrowUp, ArrowDown } from 'lucide-react';

// --- Helper for formatting Job Titles ---
const formatJobTitle = (title) => {
    if (!title) return '';
    return String(title)
        .replace(/\bSenior Vice President\b/ig, 'SVP')
        .replace(/\bVice President\b/ig, 'VP')
        .replace(/\bGeneral Manager\b/ig, 'GM')
        .replace(/\bSenior Manager\b/ig, 'Sr Manager')
        .replace(/Sr\.\s*Manager/ig, 'Sr Manager')
        .replace(/\bAssistant Manager\b/ig, 'Asst Manager');
};

// --- Helper for processing Pay Grade mapping rules (Summaries) ---
const processGrade = (gradeStr) => {
    let grade = String(gradeStr || '').split('-')[0].trim() || 'Unspecified';
    const gradeUp = grade.toUpperCase();
    
    if (['CMT', 'MT', 'GET', 'OT'].includes(gradeUp) || gradeUp.includes('TRAINEE')) {
        return 'Trainee (CMT / MT / GET / OT)';
    }
    if (gradeUp === 'DEPUTY GENERAL MANAGER') return 'DGM';
    if (gradeUp === 'DEPUTY MANAGER') return 'DM';
    
    return grade;
};

// --- Helper for processing Card Grades (Uncombined, abbreviated) ---
const formatCardGrade = (gradeStr) => {
    let grade = String(gradeStr || '').split('-')[0].trim() || 'Unspecified';
    const gradeUp = grade.toUpperCase();
    if (gradeUp === 'DEPUTY GENERAL MANAGER') return 'DGM';
    if (gradeUp === 'DEPUTY MANAGER') return 'DM';
    return grade;
};

// --- Custom sort order for grades ---
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
    if (emp._id === ceoId) return -1; // Top of the hierarchy
    const rank = getRank(emp._summaryGrade);
    
    // Custom Logic: Parent MD is handled by ceoId (-1). Other MDs rank after ED.
    // ED is at index 1. So Other MDs get rank 1.5 (between ED and SVP).
    if (rank === 0) return 1.5; 
    return rank;
};

const getMedian = (arr) => {
    if (!arr || arr.length === 0) return 0;
    const s = [...arr].sort((a,b) => a - b);
    const mid = Math.floor(s.length / 2);
    return s.length % 2 !== 0 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
};

// --- Helper for formatting grades (using object from counting Pay Grades) ---
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

// --- Helper for name normalization & Proper Casing ---
const toProperCase = (str) => {
    return str.replace(/\b\w+/g, function(txt) {
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
};

const formatDisplayFirstLast = (name) => {
    if (!name) return '';
    const parts = String(name).trim().split(/\s+/);
    let display = String(name);
    if (parts.length >= 3) {
        display = `${parts[0]} ${parts[parts.length - 1]}`;
    }
    return toProperCase(display);
};

// --- Helper for date parsing ---
const parseExcelDate = (excelDate) => {
    if (excelDate === undefined || excelDate === null || excelDate === '') return null;
    if (typeof excelDate === 'number') {
        return new Date(Math.round((excelDate - 25569) * 86400 * 1000));
    }
    const d = new Date(excelDate);
    return isNaN(d.getTime()) ? null : d;
};

const formatDateUI = (dateObj) => {
    if (!dateObj) return '-';
    return dateObj.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
};

// --- Helper for granular tenure calculations ---
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

// --- Visual Scale Component ---
const MetricScale = ({ label, min, max, median, value }) => {
    const safeMax = Math.max(max, value, 1);
    const safeMin = Math.min(min, value); 
    const range = safeMax - safeMin;
    const getPos = (v) => range === 0 ? 50 : ((v - safeMin) / range) * 100;

    const isValMin = value === min;
    const isValMed = value === median;
    const isValMax = value === max;

    const baseCircle = "absolute top-1/2 rounded-full shadow-sm transform -translate-x-1/2 -translate-y-1/2";
    const blueHollow = `${baseCircle} h-3.5 w-3.5 border-[2px] border-blue-500 bg-white z-10 cursor-help`;
    const orangeHollow = `${baseCircle} h-4 w-4 border-[2px] border-orange-500 bg-white z-20 cursor-help`;
    const blueWithOrangeFill = `${baseCircle} h-4 w-4 border-[2px] border-blue-500 bg-orange-500 z-20 cursor-help`;

    return (
        <div className="mb-5 mt-2">
            <div className="mb-1.5 text-sm text-slate-700 font-semibold leading-none">
                {label}
            </div>
            {/* Scale Container */}
            <div className="relative w-full h-4 mt-2 mb-1.5">
                {/* Track Bounds wrapper */}
                <div className="absolute left-2 right-2 top-1/2 transform -translate-y-1/2 h-1.5 bg-blue-100 rounded-full overflow-hidden">
                    {/* Orange fill up to the current value */}
                    <div className="absolute top-0 bottom-0 left-0 bg-orange-400" style={{ width: `${getPos(value)}%` }}></div>
                </div>
                
                {/* Pin Container */}
                <div className="absolute left-2 right-2 top-0 bottom-0">
                    {!isValMin && <div className={blueHollow} style={{ left: `${getPos(min)}%` }} title={`Min: ${min}`}></div>}
                    {!isValMed && median !== min && median !== max && <div className={blueHollow} style={{ left: `${getPos(median)}%` }} title={`Median: ${median}`}></div>}
                    {!isValMax && max !== min && <div className={blueHollow} style={{ left: `${getPos(max)}%` }} title={`Max: ${max}`}></div>}
                    
                    { (isValMin || isValMed || isValMax) ? (
                        <div className={blueWithOrangeFill} style={{ left: `${getPos(value)}%` }} title={`Current: ${value} (Overlaps with Benchmark)`}></div>
                    ) : (
                        <div className={orangeHollow} style={{ left: `${getPos(value)}%` }} title={`Current: ${value}`}></div>
                    )}
                </div>
            </div>
            {/* Labels Container */}
            <div className="relative w-full h-4 mt-1 text-xs font-bold px-2">
                <span title={`Min: ${min}`} className={`absolute transform -translate-x-1/2 cursor-help ${isValMin ? 'text-orange-600 z-10' : 'text-blue-500'}`} style={{ left: `${getPos(min)}%` }}>{min}</span>
                {median !== min && median !== max && (
                    <span title={`Median: ${median}`} className={`absolute transform -translate-x-1/2 cursor-help ${isValMed ? 'text-orange-600 z-10' : 'text-blue-500'}`} style={{ left: `${getPos(median)}%` }}>{median}</span>
                )}
                {max !== min && (
                    <span title={`Max: ${max}`} className={`absolute transform -translate-x-1/2 cursor-help ${isValMax ? 'text-orange-600 z-10' : 'text-blue-500'}`} style={{ left: `${getPos(max)}%` }}>{max}</span>
                )}
                {!(isValMin || isValMed || isValMax) && (
                    <span title={`Current: ${value}`} className="absolute text-orange-600 z-10 bg-white/90 px-0.5 rounded transform -translate-x-1/2 cursor-help" style={{ left: `${getPos(value)}%` }}>{value}</span>
                )}
            </div>
        </div>
    );
};

// --- Custom Benchmark Layout Box ---
const BenchmarkBox = ({ title, rightElement, borderColor = 'border-slate-200', titleColor = 'text-slate-500', children, bgClass = '' }) => (
    <div className={`relative border ${borderColor} rounded-xl p-4 pt-5 mb-6 mt-4 ${bgClass}`}>
        <div className={`absolute -top-2.5 left-3 bg-white px-2 text-xs font-bold ${titleColor} uppercase tracking-wider`}>
            {title}
        </div>
        {rightElement && (
            <div className="absolute -top-3 right-3 bg-white px-1">
                {rightElement}
            </div>
        )}
        {children}
    </div>
);

export default function App() {
  const [appTab, setAppTab] = useState('org'); 
  const [data, setData] = useState([]);
  const [employeeMap, setEmployeeMap] = useState({});
  const [globalMetrics, setGlobalMetrics] = useState({ grade: {}, exCom: null, opCom: null });
  const [attritionSummary, setAttritionSummary] = useState(null);
  const [attritionModal, setAttritionModal] = useState(null); 
  const [activeEmployeeId, setActiveEmployeeId] = useState(null);
  const [ceoId, setCeoId] = useState(null);
  
  // Search & Filters
  const [searchQuery, setSearchQuery] = useState('');
  const [showFilterPanel, setShowFilterPanel] = useState(false);
  const [showTabularResults, setShowTabularResults] = useState(false);
  const [filterMatchMode, setFilterMatchMode] = useState('and'); 
  const [filterConditions, setFilterConditions] = useState([]);
  const [openDropdown, setOpenDropdown] = useState(null);
  const [sortConfigs, setSortConfigs] = useState([]); 

  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [xlsxLoaded, setXlsxLoaded] = useState(false);
  const [viewMode, setViewMode] = useState('direct'); 
  
  const [expandedManagers, setExpandedManagers] = useState([]);

  useEffect(() => {
    if (window.XLSX) {
      setXlsxLoaded(true);
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.onload = () => setXlsxLoaded(true);
    script.onerror = () => setError("Failed to load Excel parser.");
    document.head.appendChild(script);
  }, []);

  useEffect(() => {
    setViewMode('direct');
  }, [activeEmployeeId]);

  const handleFileUpload = async (file) => {
    if (!window.XLSX) {
      setError("Excel parsing library is still loading. Please try again in a moment.");
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(buffer, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const rawData = window.XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { defval: "" });
      
      let historyData = [];
      if (workbook.SheetNames.length > 1) {
          const secondSheetName = workbook.SheetNames[1];
          historyData = window.XLSX.utils.sheet_to_json(workbook.Sheets[secondSheetName], { defval: "" });
      }

      let attritionData = [];
      if (workbook.SheetNames.length > 2) {
          const thirdSheetName = workbook.SheetNames[2];
          attritionData = window.XLSX.utils.sheet_to_json(workbook.Sheets[thirdSheetName], { defval: "" });
      }

      if (rawData.length === 0) {
        throw new Error("The uploaded Excel file is empty.");
      }

      processEmployeeData(rawData, historyData, attritionData);
    } catch (err) {
      setError(err.message || "Failed to process the Excel file. Please check the format.");
    } finally {
      setLoading(false);
    }
  };

  const processEmployeeData = (rawData, historyData, attritionData) => {
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

    const normalizeTitleForCompare = (t) => {
        return String(t).toLowerCase()
            .replace(/\bsenior\b/g, 'sr')
            .replace(/\bassistant\b/g, 'asst')
            .replace(/\bgeneral manager\b/g, 'gm')
            .replace(/\bvice president\b/g, 'vp')
            .replace(/\badministration\b/g, 'admin')
            .replace(/\btechnology\b/g, 'tech')
            .replace(/\band\b/g, '&')
            .replace(/\bhuman resources\b/g, 'hr')
            .replace(/[^a-z0-9]/g, '') 
            .trim();
    };

    const isCompanyTransfer = (reasonLow) => {
        return [
            'transfer|company change',
            'vol separation|to rpg group company',
            'rehire|from rpg group company',
            'promotion & company transfer',
            'hire|from rpg group company'
        ].some(r => reasonLow.includes(r));
    };

    const isEA = (e) => {
        if (!e) return false;
        const title = String(e['Job Title'] || '').toLowerCase();
        return title.includes('executive assistant') || title.includes('executive secretary');
    };
    
    // Corporate/Business extraction mapping
    const corporateValues = ["Corporate", "Manufacturing", "Sales & Marketing", "Sales and Marketing"];

    rawData.forEach(emp => {
      const id = emp['Users Sys Id']?.toString().trim();
      if (!id) return;
      
      emp._id = id;
      emp._tenureDate = parseExcelDate(emp['Group Date of Joining']);
      emp._summaryGrade = processGrade(emp['Pay Grade (Name)']); 
      emp._cardGrade = formatCardGrade(emp['Pay Grade (Name)']); 
      emp['Job Title'] = formatJobTitle(emp['Job Title']);
      
      // Determine Company filter value
      const business = String(emp['Business (Label)'] || '').trim();
      const companyStr = String(emp['Company (Label)'] || '').trim();
      emp._derivedCompany = corporateValues.includes(business) ? companyStr : business;
      if (!emp._derivedCompany) emp._derivedCompany = 'Unspecified';

      const events = historyMap[id] || [];
      const eventsAsc = [...events].sort((a,b) => {
          const dateA = parseExcelDate(a['Effective Start Date']) || new Date(0);
          const dateB = parseExcelDate(b['Effective Start Date']) || new Date(0);
          return dateA - dateB;
      });

      let timelineRaw = [];
      let currentBlock = null;
      let pendingGroupMovement = false;
      let hasLocationTransfer = false;

      eventsAsc.forEach(ev => {
          const rawTitle = ev['Job Title']?.trim() || emp['Job Title']?.trim(); 
          if (!rawTitle) return;
          
          const title = formatJobTitle(rawTitle);
          const normTitle = normalizeTitleForCompare(title);
          const grade = formatCardGrade(ev['Pay Grade (Name)']);
          const location = (ev['Location (Location Name)'] || ev['Location Name'] || ev['Location'])?.trim();
          const reason = String(ev['Event Reason (Label)'] || '').trim();
          const reasonLow = reason.toLowerCase();
          const startDate = parseExcelDate(ev['Effective Start Date']);
          const jobCode = String(ev['Job Code (Job Code)'] || '').trim();
          const position = String(ev['Position'] || '').trim();

          if (reasonLow.includes('location transfer') || reasonLow.includes('location change')) {
              hasLocationTransfer = true;
          }

          if (reasonLow.includes('separation')) {
              if (currentBlock) {
                  currentBlock.endDate = startDate;
                  timelineRaw.push(currentBlock);
                  currentBlock = null;
              }
              if (reasonLow.includes('to rpg group company')) {
                  pendingGroupMovement = true;
              }
              return; 
          }

          let isNewBlock = false;
          let blockType = 'other';

          if (!currentBlock) {
              isNewBlock = true;
              if (pendingGroupMovement || isCompanyTransfer(reasonLow)) blockType = 'company_transfer';
              else if (reasonLow.includes('promotion')) blockType = 'promotion';
              else if (reasonLow.includes('transfer')) blockType = 'transfer';
              else if (reasonLow.includes('hire')) blockType = 'hire';
              else blockType = 'hire'; 
              pendingGroupMovement = false;
          } else {
              const gradeChanged = currentBlock.grade !== 'Unspecified' && grade !== 'Unspecified' && currentBlock.grade !== grade;
              const titleChanged = currentBlock.normTitle !== normTitle;
              const jcChanged = currentBlock.jobCode && jobCode && currentBlock.jobCode !== jobCode;
              const posChanged = currentBlock.position && position && currentBlock.position !== position;
              const isDataChangeOnly = reasonLow === 'data change' || reasonLow.includes('data change|');

              if (gradeChanged) {
                  isNewBlock = true;
                  blockType = 'promotion';
              } else if (isCompanyTransfer(reasonLow)) {
                  if (titleChanged) {
                      isNewBlock = true;
                      blockType = 'company_transfer';
                  }
              } else if (reasonLow.includes('promotion')) {
                  isNewBlock = true;
                  blockType = 'promotion';
              } else if (reasonLow.includes('transfer') || reasonLow.includes('rotation')) {
                  if (jcChanged || posChanged || titleChanged) {
                      isNewBlock = true;
                      blockType = 'transfer';
                  }
              } else if (isDataChangeOnly) {
                  isNewBlock = false;
              } else if (reasonLow.includes('job change')) {
                   isNewBlock = false; 
              } else {
                  if (titleChanged || jcChanged) {
                      isNewBlock = true;
                      blockType = 'transfer';
                  }
              }
          }

          if (isNewBlock) {
              if (currentBlock) {
                  currentBlock.endDate = startDate;
                  timelineRaw.push(currentBlock);
              }
              currentBlock = {
                  title, normTitle, grade, location, startDate, endDate: null, type: blockType, jobCode, position
              };
          } else if (currentBlock) {
              if (location) currentBlock.location = location;
              currentBlock.title = title; 
              currentBlock.normTitle = normTitle;
              currentBlock.jobCode = jobCode;
              currentBlock.position = position;

              if (startDate && currentBlock.startDate && Math.abs(startDate - currentBlock.startDate) < 172800000) {
                  if (reasonLow.includes('promotion')) currentBlock.type = 'promotion';
                  else if (isCompanyTransfer(reasonLow) && currentBlock.type !== 'promotion') currentBlock.type = 'company_transfer';
              }
          }
      });

      if (currentBlock) {
          currentBlock.endDate = new Date();
          timelineRaw.push(currentBlock);
      }

      let finalTimeline = [];
      for (let i = 0; i < timelineRaw.length; i++) {
          let block = timelineRaw[i];
          if (i < timelineRaw.length - 1 && block.startDate && block.endDate) {
              const durationMs = block.endDate - block.startDate;
              if (durationMs < 30 * 24 * 60 * 60 * 1000) continue;
          }
          if (finalTimeline.length > 0) {
              const prevBlock = finalTimeline[finalTimeline.length - 1];
              if (prevBlock.endDate && block.startDate) {
                  const gapMs = block.startDate - prevBlock.endDate;
                  if (gapMs > 30 * 24 * 60 * 60 * 1000) {
                      finalTimeline.push({
                          isGap: true,
                          duration: formatDuration(prevBlock.endDate, block.startDate)
                      });
                  }
              }
          }
          finalTimeline.push(block);
      }

      finalTimeline = finalTimeline.reverse();
      let timeInRoleDate = finalTimeline.find(r => !r.isGap)?.startDate || null;
      let lastPromotionDate = finalTimeline.find(r => r.type === 'promotion')?.startDate || null;

      const eventsDesc = [...events].sort((a,b) => {
          const dateA = parseExcelDate(a['Effective Start Date']) || new Date(0);
          const dateB = parseExcelDate(b['Effective Start Date']) || new Date(0);
          return dateB - dateA; 
      });

      let timeWithManagerDate = null;
      const currentManagerId = emp['Line Manager UserID'] ? String(emp['Line Manager UserID']).trim() : undefined;
      const currentManagerName = emp['Line Manager Name'] ? String(emp['Line Manager Name']).trim() : undefined;
      
      if (eventsDesc.length > 0 && (currentManagerId || currentManagerName)) {
          let oldestValidDate = parseExcelDate(eventsDesc[0]['Effective Start Date']);
          for (let i = 0; i < eventsDesc.length; i++) {
              const ev = eventsDesc[i];
              const evManagerStr = String(ev['Line Manager'] || '').trim();
              const reason = String(ev['Event Reason (Label)'] || '').toLowerCase();
              const isMatch = (evManagerStr === currentManagerId) || (evManagerStr === currentManagerName);
              if (!isMatch && evManagerStr !== '') break; 
              
              const evDate = parseExcelDate(ev['Effective Start Date']);
              if (evDate) oldestValidDate = evDate;
              if (reason.includes('manager change')) break;
          }
          timeWithManagerDate = oldestValidDate;
      }

      emp._timeline = finalTimeline.map(role => {
          if (role.isGap) return role;
          return { ...role, durationFormatted: formatDuration(role.startDate, role.endDate) };
      });

      emp._hasLocationTransfer = hasLocationTransfer;
      emp._timeInRoleDate = timeInRoleDate;
      emp._lastPromotionDate = lastPromotionDate;
      emp._timeWithManagerDate = timeWithManagerDate;

      empMap[id] = { ...emp };
    });

    let ceos = Object.values(empMap).filter(emp => 
        (!emp['Line Manager Name'] || String(emp['Line Manager Name']).trim() === '') &&
        (!emp['Line Manager UserID'] || String(emp['Line Manager UserID']).trim() === '') &&
        (!emp['Matrix Manager Name'] || String(emp['Matrix Manager Name']).trim() === '') &&
        (!emp['Matrix Manager ID'] || String(emp['Matrix Manager ID']).trim() === '') &&
        (!emp['Matrix Manager UserID'] || String(emp['Matrix Manager UserID']).trim() === '')
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

    // Exact ID Mapping for hierarchical links
    const sysIdMap = {};
    const usernameMap = {};
    
    Object.values(empMap).forEach(emp => {
        const sysId = emp['Users Sys Id']?.toString().trim().toLowerCase();
        const username = emp['Username']?.toString().trim().toLowerCase();

        if (sysId) sysIdMap[sysId] = emp._id;
        if (username) usernameMap[username] = emp._id;
    });

    Object.values(empMap).forEach(emp => {
      const lineManagerUserId = emp['Line Manager UserID']?.toString().trim().toLowerCase();
      const matrixManagerId = emp['Matrix Manager ID']?.toString().trim().toLowerCase();

      const managerId = lineManagerUserId ? sysIdMap[lineManagerUserId] : null;
      const matrixId = matrixManagerId ? usernameMap[matrixManagerId] : null;

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

    const addGrades = (target, source) => {
        for (const [grade, count] of Object.entries(source)) {
            target[grade] = (target[grade] || 0) + count;
        }
    };

    const calculateInsights = (empId, visited = new Set()) => {
      if (visited.has(empId)) return empMap[empId]._insights;
      visited.add(empId);

      const directs = directReportsMap[empId] || [];
      const matrix = matrixReportsMap[empId] || [];
      
      let totalTeam = 0;
      let directCount = 0;
      let genderCount = { male: 0, female: 0, other: 0 };
      let probationCount = 0;
      let noticeCount = 0;

      const directGrades = {};
      const matrixGrades = {};
      const teamGrades = {};

      directs.forEach(childId => {
        const child = empMap[childId];
        if (!child) return;

        const childInsights = calculateInsights(childId, visited);

        // DO NOT aggregate insights if the DR is an EA
        if (!isEA(child)) {
            const grade = child._summaryGrade;
            directGrades[grade] = (directGrades[grade] || 0) + 1;
            teamGrades[grade] = (teamGrades[grade] || 0) + 1;

            const gender = String(child['Gender'] || '').toLowerCase();
            if (gender.startsWith('m')) genderCount.male++;
            else if (gender.startsWith('f')) genderCount.female++;
            else genderCount.other++;

            const empStatus = String(child['Employee Status (Picklist Label)'] || '').toLowerCase();
            const confStatus = String(child['Confirmation Status (Picklist Label)'] || '').toLowerCase();
            
            if (empStatus.includes('probation') || confStatus.includes('probation')) {
                probationCount++;
            }

            const lwdDate = String(child['Last Working Date'] || '').trim();
            const hasResigned = lwdDate.length > 0 && lwdDate.toLowerCase() !== 'na' && lwdDate.toLowerCase() !== 'null' && lwdDate !== '-';
            
            if (empStatus.includes('notice') || hasResigned) {
                noticeCount++;
            }

            directCount++;
            totalTeam += 1 + (childInsights ? childInsights.totalTeam : 0);
            if (childInsights) {
                addGrades(teamGrades, childInsights.teamGrades);
            }
        }
      });

      matrix.forEach(childId => {
          const child = empMap[childId];
          if(!child) return;
          matrixGrades[child._summaryGrade] = (matrixGrades[child._summaryGrade] || 0) + 1;
      });

      const insights = {
        directCount,
        matrixCount: matrix.length,
        totalTeam,
        directGrades,
        matrixGrades,
        teamGrades,
        genderCount,
        probationCount,
        noticeCount
      };

      empMap[empId]._insights = insights;
      empMap[empId]._directs = directs;
      empMap[empId]._matrix = matrix;

      return insights;
    };

    Object.keys(empMap).forEach(id => calculateInsights(id));

    // ExCom & OpCom Identification 
    const exComIds = [];
    const opComIds = [];

    Object.values(empMap).forEach(emp => {
        if (isEA(emp)) return;
        const rank = getRank(emp._summaryGrade);
        const hasTeam = (emp._insights?.totalTeam || 0) > 0;
        
        // ExCom: SVP & above (rank <= 2) with a team
        if (rank <= 2 && hasTeam) {
            emp._isExCom = true;
            if (emp._id !== computedCeoId) exComIds.push(emp._id);
        }
    });

    Object.values(empMap).forEach(emp => {
        if (isEA(emp)) return;
        const rank = getRank(emp._summaryGrade);
        const manager = empMap[emp._managerId];
        
        // OpCom: VP & GM (rank 3 or 4) reporting directly to an ExCom
        if ((rank === 3 || rank === 4) && manager && manager._isExCom) {
            emp._isOpCom = true;
            opComIds.push(emp._id);
        }
    });

    // Compile Grade Global Metrics for Benchmarking (excluding EAs)
    const gradeStatsBuilder = {};
    Object.values(empMap).forEach(emp => {
        if (isEA(emp)) return;
        const grade = emp._summaryGrade;
        if (!gradeStatsBuilder[grade]) gradeStatsBuilder[grade] = { drs: [], teams: [] };
        gradeStatsBuilder[grade].drs.push(emp._insights?.directCount || 0);
        gradeStatsBuilder[grade].teams.push(emp._insights?.totalTeam || 0);
    });

    const gradeMetricsFinal = {};
    Object.keys(gradeStatsBuilder).forEach(g => {
        const drs = gradeStatsBuilder[g].drs;
        const teams = gradeStatsBuilder[g].teams;
        gradeMetricsFinal[g] = {
            drMin: Math.min(...drs),
            drMax: Math.max(...drs),
            drMedian: getMedian(drs),
            teamMin: Math.min(...teams),
            teamMax: Math.max(...teams),
            teamMedian: getMedian(teams)
        };
    });

    // ExCom Metrics (Excluding Top Node)
    const exComDrs = exComIds.map(id => empMap[id]._insights.directCount);
    const exComTeams = exComIds.map(id => empMap[id]._insights.totalTeam);
    const exComMetrics = {
        drMin: exComDrs.length ? Math.min(...exComDrs) : 0,
        drMax: exComDrs.length ? Math.max(...exComDrs) : 0,
        drMedian: getMedian(exComDrs),
        teamMin: exComTeams.length ? Math.min(...exComTeams) : 0,
        teamMax: exComTeams.length ? Math.max(...exComTeams) : 0,
        teamMedian: getMedian(exComTeams)
    };

    // OpCom Metrics
    const opComDrs = opComIds.map(id => empMap[id]._insights.directCount);
    const opComTeams = opComIds.map(id => empMap[id]._insights.totalTeam);
    const opComMetrics = {
        drMin: opComDrs.length ? Math.min(...opComDrs) : 0,
        drMax: opComDrs.length ? Math.max(...opComDrs) : 0,
        drMedian: getMedian(opComDrs),
        teamMin: opComTeams.length ? Math.min(...opComTeams) : 0,
        teamMax: opComTeams.length ? Math.max(...opComTeams) : 0,
        teamMedian: getMedian(opComTeams)
    };

    // Calculate Cross-Peer / Manager-Relative Insights
    Object.values(empMap).forEach(emp => {
      const managerId = emp._managerId;
      if (managerId && empMap[managerId]) {
        const manager = empMap[managerId];
        
        const peers = (manager._directs || []).filter(id => {
            if (id === emp._id) return false;
            return !isEA(empMap[id]);
        });
        
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
    });

    setData(Object.values(empMap));
    setEmployeeMap(empMap);
    setGlobalMetrics({ grade: gradeMetricsFinal, exCom: exComMetrics, opCom: opComMetrics });
    
    if (actualCEO) {
        setCeoId(actualCEO._id);
        setActiveEmployeeId(actualCEO._id);
    }

    if (attritionData && attritionData.length > 0) {
        const attStats = {
            total: 0,
            vol: 0,
            invol: 0,
            ret: 0,
            tree: []
        };

        const directAttritionsMap = {};
        const attritionBySysId = {};

        attritionData.forEach(row => {
            const managerId = String(row['Line Manager Userid'] || row['Line Manager User ID'] || '').trim();
            const managerName = String(row['Line Manager'] || '').trim();
            const reason = String(row['Voluntary/In-voluntary/Retirement'] || '').toLowerCase();
            const sysId = String(row['Users Sys Id']).trim();
            
            if (sysId) attritionBySysId[sysId] = row;
            
            attStats.total++;
            if (reason.includes('voluntary') && !reason.includes('in-voluntary') && !reason.includes('involuntary')) attStats.vol++;
            else if (reason.includes('in-voluntary') || reason.includes('involuntary')) attStats.invol++;
            else if (reason.includes('retire')) attStats.ret++;

            const record = {
                name: formatDisplayFirstLast(row['Display Name'] || row['Username']),
                title: formatJobTitle(row['Job Title']),
                reason: toProperCase(row['Voluntary/In-voluntary/Retirement'] || 'Unknown'),
                lwd: parseExcelDate(row['Last Working Date']),
                positionCode: String(row['Position code'] || row['Position'] || '').trim(),
                managerName
            };

            if (managerId) {
                if (!directAttritionsMap[managerId]) directAttritionsMap[managerId] = [];
                directAttritionsMap[managerId].push(record);
            }
        });

        const inactiveManagerNodes = [];
        Object.keys(directAttritionsMap).forEach(mgrId => {
            if (!empMap[mgrId]) {
                const records = directAttritionsMap[mgrId];
                const mgrDetails = attritionBySysId[mgrId];
                
                let targetManagerId = null;
                const firstRecWithPos = records.find(r => r.positionCode);
                if (firstRecWithPos && firstRecWithPos.positionCode) {
                    const incumbent = Object.values(empMap).find(e => String(e['Position'] || '').trim() === firstRecWithPos.positionCode);
                    if (incumbent && incumbent._managerId) {
                        targetManagerId = incumbent._managerId;
                    }
                }

                inactiveManagerNodes.push({
                    managerId: mgrId,
                    name: mgrDetails ? formatDisplayFirstLast(mgrDetails['Display Name'] || mgrDetails['Username']) : (records[0].managerName || mgrId),
                    title: mgrDetails ? formatJobTitle(mgrDetails['Job Title']) : '-',
                    department: mgrDetails ? (mgrDetails['Department (Label)'] || 'Unknown') : '-',
                    currentTeamSize: 0,
                    directAttritions: records.sort((a,b) => (b.lwd || 0) - (a.lwd || 0)),
                    teamCount: records.length,
                    children: [],
                    isInactive: true,
                    targetManagerId
                });
            }
        });

        const inactiveNodesByTarget = {};
        const orphanedInactiveNodes = [];
        inactiveManagerNodes.forEach(node => {
            if (node.targetManagerId) {
                if (!inactiveNodesByTarget[node.targetManagerId]) inactiveNodesByTarget[node.targetManagerId] = [];
                inactiveNodesByTarget[node.targetManagerId].push(node);
            } else {
                orphanedInactiveNodes.push(node);
            }
        });

        const buildNode = (empId) => {
            const emp = empMap[empId];
            const directAttritions = directAttritionsMap[empId] || [];
            const childrenIds = emp._directs || [];
            
            const children = [];
            let teamCount = directAttritions.length;
            
            childrenIds.forEach(childId => {
                const childNode = buildNode(childId);
                if (childNode.teamCount > 0) {
                    children.push(childNode);
                    teamCount += childNode.teamCount;
                }
            });

            const attachedInactive = inactiveNodesByTarget[empId] || [];
            attachedInactive.forEach(inactNode => {
                children.push(inactNode);
                teamCount += inactNode.teamCount;
            });
            
            children.sort((a,b) => {
                const rankA = getEmpSortRank(empMap[a.managerId], computedCeoId);
                const rankB = getEmpSortRank(empMap[b.managerId], computedCeoId);
                if (rankA !== rankB) return rankA - rankB;
                return b.teamCount - a.teamCount;
            });

            return {
                managerId: empId,
                name: formatDisplayFirstLast(emp['Display Name']),
                title: formatJobTitle(emp['Job Title']),
                department: emp['Department (Label)'] || '-',
                currentTeamSize: emp._insights?.totalTeam || 0,
                directAttritions: directAttritions.sort((a,b) => (b.lwd || 0) - (a.lwd || 0)),
                teamCount,
                children,
                isInactive: false
            };
        };

        const rootNodes = [];
        if (actualCEO) {
            const ceoNode = buildNode(actualCEO._id);
            if (ceoNode.teamCount > 0) rootNodes.push(ceoNode);
            setExpandedManagers([actualCEO._id]);
        } else if (ceos.length > 0) {
            ceos.forEach(c => {
                const n = buildNode(c._id);
                if (n.teamCount > 0) rootNodes.push(n);
            });
        }

        orphanedInactiveNodes.forEach(node => {
            if (node.teamCount > 0) rootNodes.push(node);
        });

        rootNodes.sort((a,b) => {
            const rankA = getEmpSortRank(empMap[a.managerId], computedCeoId);
            const rankB = getEmpSortRank(empMap[b.managerId], computedCeoId);
            if (rankA !== rankB) return rankA - rankB;
            return b.teamCount - a.teamCount;
        });
        attStats.tree = rootNodes;
        setAttritionSummary(attStats);
    }
  };

  const handleDragOver = (e) => { e.preventDefault(); setIsDragging(true); };
  const handleDragLeave = () => setIsDragging(false);
  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFileUpload(e.dataTransfer.files[0]);
    }
  };

  // --- Filtering & Sorting Logic ---
  const allUniqueGrades = useMemo(() => {
      const grades = data.map(emp => emp._summaryGrade).filter(Boolean);
      return [...new Set(grades)].sort((a, b) => getRank(a) - getRank(b));
  }, [data]);

  const allUniqueCompanies = useMemo(() => {
      const comps = data.map(emp => emp._derivedCompany).filter(Boolean);
      return [...new Set(comps)].sort((a, b) => a.localeCompare(b));
  }, [data]);

  const filteredSearch = useMemo(() => {
    if (!searchQuery) return [];
    const query = searchQuery.toLowerCase();
    return data
      .filter(emp => String(emp['Display Name'] || '').toLowerCase().includes(query) || formatJobTitle(emp['Job Title']).toLowerCase().includes(query))
      .slice(0, 5);
  }, [searchQuery, data]);

  const tabularFilteredData = useMemo(() => {
      if (filterConditions.length === 0) return [];
      let filtered = data.filter(emp => {
          const results = filterConditions.map(cond => {
              if (!cond.value && cond.value !== 0 && cond.field !== 'Grade' && cond.field !== 'Company') return false;
              if ((cond.field === 'Grade' || cond.field === 'Company') && cond.value.length === 0) return false;

              let empVal;
              if (cond.field === 'Designation') empVal = emp['Job Title'] || '';
              else if (cond.field === 'Team Size') empVal = emp._insights?.totalTeam || 0;
              else if (cond.field === 'DR Size') empVal = emp._insights?.directCount || 0;
              else if (cond.field === 'Grade') empVal = emp._summaryGrade || '';
              else if (cond.field === 'Company') empVal = emp._derivedCompany || '';

              if (cond.field === 'Designation') {
                  const valLower = cond.value.toLowerCase();
                  const empValLower = empVal.toLowerCase();
                  if (cond.operator === 'contains') return empValLower.includes(valLower);
                  if (cond.operator === 'equals') return empValLower === valLower;
              } else if (cond.field === 'Team Size' || cond.field === 'DR Size') {
                  const numVal = Number(cond.value);
                  if (isNaN(numVal)) return false;
                  if (cond.operator === '=') return empVal === numVal;
                  if (cond.operator === '>') return empVal > numVal;
                  if (cond.operator === '<') return empVal < numVal;
                  if (cond.operator === '>=') return empVal >= numVal;
                  if (cond.operator === '<=') return empVal <= numVal;
              } else if (cond.field === 'Grade' || cond.field === 'Company') {
                  return cond.value.includes(empVal);
              }
              return false;
          });

          if (results.length === 0) return false;
          if (filterMatchMode === 'and') return results.every(r => r);
          if (filterMatchMode === 'or') return results.some(r => r);
          return false;
      });

      // Apply Reverse Cascading Sort
      if (sortConfigs.length > 0) {
          filtered.sort((a, b) => {
              for (let config of sortConfigs) {
                  let valA, valB;
                  switch (config.field) {
                      case 'Employee': 
                          valA = formatDisplayFirstLast(a['Display Name']); 
                          valB = formatDisplayFirstLast(b['Display Name']); 
                          break;
                      case 'JobTitle': 
                          valA = a['Job Title'] || ''; 
                          valB = b['Job Title'] || ''; 
                          break;
                      case 'Grade': 
                          valA = getEmpSortRank(a, ceoId); 
                          valB = getEmpSortRank(b, ceoId); 
                          break;
                      case 'Department': 
                          valA = a['Department (Label)'] || ''; 
                          valB = b['Department (Label)'] || ''; 
                          break;
                      case 'DRSize': 
                          valA = a._insights?.directCount || 0; 
                          valB = b._insights?.directCount || 0; 
                          break;
                      case 'MatrixSize': 
                          valA = a._insights?.matrixCount || 0; 
                          valB = b._insights?.matrixCount || 0; 
                          break;
                      case 'TeamSize': 
                          valA = a._insights?.totalTeam || 0; 
                          valB = b._insights?.totalTeam || 0; 
                          break;
                      case 'Manager': 
                          valA = formatDisplayFirstLast(a['Line Manager Name'] || ''); 
                          valB = formatDisplayFirstLast(b['Line Manager Name'] || ''); 
                          break;
                      default: valA = ''; valB = '';
                  }

                  if (valA === valB) continue;

                  let cmp = 0;
                  if (typeof valA === 'string' && typeof valB === 'string') {
                      cmp = valA.localeCompare(valB);
                  } else {
                      cmp = valA > valB ? 1 : -1;
                  }
                  return config.dir === 'asc' ? cmp : -cmp;
              }
              return 0;
          });
      }
      return filtered;
  }, [data, filterConditions, filterMatchMode, sortConfigs, ceoId]);

  const addFilterCondition = () => {
      setFilterConditions([...filterConditions, { id: Date.now(), field: 'Designation', operator: 'contains', value: '' }]);
      setShowTabularResults(true); // Auto-show table on new filter
  };

  const updateFilterCondition = (id, key, val) => {
      setFilterConditions(filterConditions.map(c => c.id === id ? { ...c, [key]: val } : c));
      if (key === 'field') {
          setFilterConditions(prev => prev.map(c => {
              if (c.id === id) {
                  let defaultOp = 'contains';
                  let defaultVal = '';
                  if (val === 'Team Size' || val === 'DR Size') defaultOp = '>=';
                  if (val === 'Grade' || val === 'Company') { defaultOp = 'in'; defaultVal = []; }
                  return { ...c, operator: defaultOp, value: defaultVal };
              }
              return c;
          }));
      }
      setShowTabularResults(true); // Auto-show table on filter update
  };

  const removeFilterCondition = (id) => {
      setFilterConditions(filterConditions.filter(c => c.id !== id));
  };

  const handleSort = (field) => {
      setSortConfigs(prev => {
          const existingIdx = prev.findIndex(c => c.field === field);
          if (existingIdx === -1) {
              return [...prev, { field, dir: 'asc' }];
          } else {
              const existing = prev[existingIdx];
              const newConfigs = [...prev];
              if (existing.dir === 'asc') {
                  newConfigs[existingIdx] = { field, dir: 'desc' };
                  return newConfigs;
              } else {
                  newConfigs.splice(existingIdx, 1);
                  return newConfigs;
              }
          }
      });
  };

  const renderSortIcon = (field) => {
      const config = sortConfigs.find(c => c.field === field);
      if (!config) return <div className="w-4 inline-block"></div>;
      return config.dir === 'asc' ? <ArrowUp size={14} className="inline ml-1 text-blue-600"/> : <ArrowDown size={14} className="inline ml-1 text-blue-600"/>;
  };

  const SortableHeader = ({ label, field, align = 'left' }) => (
      <th 
          className={`px-6 py-3 font-semibold cursor-pointer bg-slate-50 hover:bg-slate-200 select-none transition-colors ${align === 'center' ? 'text-center' : 'text-left'}`}
          onClick={() => handleSort(field)}
      >
          <div className={`flex items-center ${align === 'center' ? 'justify-center' : 'justify-start'}`}>
              {label} {renderSortIcon(field)}
          </div>
      </th>
  );

  const renderManagerRows = (nodes, level = 0) => {
      return nodes.map((node) => {
          const isExpanded = expandedManagers.includes(node.managerId);
          const hasChildren = node.children.length > 0;
          
          return (
              <React.Fragment key={`${node.managerId}-${level}`}>
                  <tr 
                      className={`border-b border-slate-100 hover:bg-slate-50 cursor-pointer transition-colors ${isExpanded ? 'bg-blue-50/50' : ''}`}
                      onClick={() => {
                          if (!hasChildren) return;
                          setExpandedManagers(prev => 
                              prev.includes(node.managerId) 
                                  ? prev.filter(id => id !== node.managerId) 
                                  : [...prev, node.managerId]
                          )
                      }}
                  >
                      <td className="py-3 text-slate-400 relative" style={{ paddingLeft: `${1.5 + level * 2}rem`, width: '3rem' }}>
                          {level > 0 && (
                              <div className="absolute top-0 bottom-0 w-1 rounded-full bg-blue-200" style={{ left: `${1.5 + (level - 1) * 2 + 0.3}rem` }}></div>
                          )}
                          {hasChildren ? (isExpanded ? <ChevronDown size={18}/> : <ChevronRight size={18}/>) : <div className="w-[18px]"></div>}
                      </td>
                      <td className="px-6 py-3">
                          <div className="flex items-center gap-2">
                              <span className={`font-bold ${node.isInactive ? 'text-slate-600' : 'text-slate-800'}`}>{node.name}</span>
                              {node.isInactive && (
                                  <span className="flex items-center text-slate-400 italic text-xs font-normal">
                                      <div className="h-1.5 w-1.5 rounded-full bg-orange-500 mr-1.5"></div>
                                      Separated
                                  </span>
                              )}
                          </div>
                          <div className="text-xs text-slate-500 mt-0.5">{node.title}</div>
                      </td>
                      <td className="px-6 py-3 text-slate-600">{node.department}</td>
                      <td className="px-6 py-3 text-center">
                          <span className="bg-slate-100 text-slate-600 px-2.5 py-1 rounded-full font-medium">{node.isInactive ? '-' : (node.currentTeamSize || '-')}</span>
                      </td>
                      <td className="px-6 py-3 text-center">
                          {node.directAttritions.length > 0 ? (
                              <button 
                                  onClick={(e) => {
                                      e.stopPropagation();
                                      setAttritionModal(node);
                                  }}
                                  className="px-3 py-1 bg-blue-100 text-blue-700 rounded-full font-bold shadow-sm border border-blue-200 hover:bg-blue-200 transition-colors"
                              >
                                  {node.directAttritions.length}
                              </button>
                          ) : (
                              <span className="text-slate-400">-</span>
                          )}
                      </td>
                      <td className="px-6 py-3 text-center font-bold text-slate-700">
                          {node.teamCount > 0 ? node.teamCount : '-'}
                      </td>
                  </tr>
                  
                  {isExpanded && node.children.length > 0 && renderManagerRows(node.children, level + 1)}
              </React.Fragment>
          );
      });
  };

  if (data.length === 0) {
    return (
      <div className="h-screen w-full bg-slate-50 flex items-center justify-center p-6">
        <div 
          className={`max-w-xl w-full bg-white p-10 rounded-2xl shadow-xl border-2 border-dashed transition-colors ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-slate-300'}`}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
        >
          <div className="flex flex-col items-center text-center space-y-4">
            <div className="w-20 h-20 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center">
              <Upload size={40} />
            </div>
            <h2 className="text-2xl font-bold text-slate-800">Upload Employee Data</h2>
            <p className="text-slate-500 text-sm">
              Drag and drop your Excel (.xlsx) file here.<br/>
              Ensure it contains standard employee details, an optional 2nd sheet for Job History, and an optional 3rd sheet for Attrition.
            </p>
            <input 
              type="file" 
              accept=".xlsx, .xls" 
              className="hidden" 
              id="file-upload"
              disabled={!xlsxLoaded}
              onChange={(e) => e.target.files[0] && handleFileUpload(e.target.files[0])}
            />
            <label 
              htmlFor="file-upload" 
              className={`px-6 py-3 text-white font-medium rounded-lg transition-colors shadow-md ${!xlsxLoaded || loading ? 'bg-slate-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700 cursor-pointer'}`}
            >
              {!xlsxLoaded ? 'Loading Library...' : loading ? 'Processing...' : 'Select Excel File'}
            </label>
            {error && <p className="text-red-500 text-sm mt-4 font-medium">{error}</p>}
          </div>
        </div>
      </div>
    );
  }

  const activeEmployee = employeeMap[activeEmployeeId];
  const manager = activeEmployee?._managerId ? employeeMap[activeEmployee._managerId] : null;
  const directReports = (activeEmployee?._directs || [])
    .map(id => employeeMap[id])
    .filter(Boolean)
    .sort((a, b) => {
        const rankA = getEmpSortRank(a, ceoId);
        const rankB = getEmpSortRank(b, ceoId);
        if (rankA !== rankB) return rankA - rankB;
        return (b._insights?.totalTeam || 0) - (a._insights?.totalTeam || 0);
    });
  const matrixReports = (activeEmployee?._matrix || [])
    .map(id => employeeMap[id])
    .filter(Boolean)
    .sort((a, b) => {
        const rankA = getEmpSortRank(a, ceoId);
        const rankB = getEmpSortRank(b, ceoId);
        if (rankA !== rankB) return rankA - rankB;
        return (b._insights?.totalTeam || 0) - (a._insights?.totalTeam || 0);
    });

  const isMatrixView = viewMode === 'matrix';
  const displayedReports = isMatrixView ? matrixReports : directReports;

  return (
    <div className="h-screen w-full flex flex-col font-sans text-slate-800 bg-slate-50 overflow-hidden">
      {/* Header */}
      <header className="bg-white border-b px-6 py-4 flex items-center justify-between shadow-sm sticky top-0 z-20 flex-shrink-0">
        <div className="flex items-center space-x-3 w-1/3">
          <div className="bg-blue-600 p-2 rounded-lg"><Users className="text-white" size={24} /></div>
          <h1 className="text-xl font-bold text-slate-800 hidden sm:block">Org Insights</h1>
        </div>

        {/* View Toggle */}
        <div className="flex bg-slate-100 p-1 rounded-lg border border-slate-200 w-fit mx-auto justify-center">
            <button 
                onClick={() => { setAppTab('org'); setFilterConditions([]); setShowFilterPanel(false); setShowTabularResults(false); }}
                className={`px-4 py-1.5 rounded-md text-sm font-semibold transition-all ${appTab === 'org' ? 'bg-white text-blue-600 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
            >
                Org Chart
            </button>
            {attritionSummary && (
                <button 
                    onClick={() => { setAppTab('attrition'); setFilterConditions([]); setShowFilterPanel(false); setShowTabularResults(false); }}
                    className={`px-4 py-1.5 rounded-md text-sm font-semibold transition-all flex items-center gap-2 ${appTab === 'attrition' ? 'bg-white text-blue-600 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
                >
                    <UserMinus size={14} /> Attrition
                </button>
            )}
        </div>

        {/* Search Bar & Actions */}
        <div className="flex items-center justify-end space-x-4 w-1/3">
          {appTab === 'org' && (
              <>
                  <button 
                    onClick={() => setShowFilterPanel(!showFilterPanel)}
                    className={`p-2 rounded-lg transition-colors border ${showFilterPanel || filterConditions.length > 0 ? 'bg-blue-50 border-blue-200 text-blue-600' : 'bg-slate-50 border-slate-200 text-slate-500 hover:bg-slate-100'}`}
                    title="Advanced Filters"
                  >
                    <Filter size={18} />
                  </button>

                  <div className="relative w-64 hidden md:block">
                    <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-slate-400" size={18} />
                    <input 
                      type="text" 
                      placeholder="Search employee..." 
                      className="w-full pl-10 pr-4 py-1.5 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 bg-slate-50 text-sm"
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                    />
                    {searchQuery && (
                      <div className="absolute top-full right-0 mt-2 w-80 bg-white rounded-lg shadow-xl border overflow-hidden z-50">
                        {filteredSearch.length > 0 ? (
                          filteredSearch.map(emp => (
                            <button
                              key={emp._id}
                              className="w-full text-left px-4 py-3 hover:bg-slate-50 border-b last:border-0 flex flex-col"
                              onClick={() => {
                                setActiveEmployeeId(emp._id);
                                setSearchQuery('');
                                setShowTabularResults(false);
                              }}
                            >
                              <span className="font-semibold text-slate-800">{formatDisplayFirstLast(emp['Display Name'])}</span>
                              <span className="text-xs text-slate-500">{emp['Job Title']} • {emp['Department (Label)']}</span>
                            </button>
                          ))
                        ) : (
                          <div className="px-4 py-3 text-slate-500 text-sm">No employees found.</div>
                        )}
                      </div>
                    )}
                  </div>
                  
                  {appTab === 'org' && filterConditions.length > 0 && !showTabularResults && (
                      <button
                          onClick={() => setShowTabularResults(true)}
                          className="px-4 py-1.5 bg-blue-50 text-blue-600 hover:bg-blue-100 rounded-lg font-medium transition-colors text-sm whitespace-nowrap"
                      >
                          Return to Tabular View
                      </button>
                  )}

                  {ceoId && filterConditions.length === 0 && (
                    <button
                      onClick={() => { 
                          setActiveEmployeeId(ceoId); 
                          setViewMode('direct');
                          const container = document.getElementById('chart-container');
                          if (container) container.scrollTo({ top: 0, behavior: 'smooth' });
                      }}
                      className="px-4 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg font-medium transition-colors text-sm whitespace-nowrap"
                    >
                      Go to Top
                    </button>
                  )}
              </>
          )}
        </div>
      </header>

      {/* Advanced Filter Panel */}
      {appTab === 'org' && showFilterPanel && (
          <div className="bg-white border-b border-slate-200 p-6 shadow-sm z-10 flex-shrink-0 animate-fade-in-down relative">
              <div className="max-w-5xl mx-auto">
                  <div className="flex justify-between items-center mb-4">
                      <h3 className="font-bold text-slate-700 flex items-center gap-2"><Filter size={16} /> Advanced Filters</h3>
                      <div className="flex items-center gap-3">
                          <div className="flex bg-slate-100 rounded-md p-0.5 border border-slate-200">
                              <button onClick={() => setFilterMatchMode('and')} className={`px-4 py-1 text-xs rounded transition-colors ${filterMatchMode === 'and' ? 'bg-white shadow-sm font-bold text-blue-600' : 'text-slate-500 hover:text-slate-700'}`}>Match All</button>
                              <button onClick={() => setFilterMatchMode('or')} className={`px-4 py-1 text-xs rounded transition-colors ${filterMatchMode === 'or' ? 'bg-white shadow-sm font-bold text-blue-600' : 'text-slate-500 hover:text-slate-700'}`}>Match Any</button>
                          </div>
                      </div>
                  </div>

                  <div className="space-y-3">
                      {filterConditions.map((cond, index) => (
                          <div key={cond.id} className="flex flex-wrap items-center gap-3 bg-slate-50 p-2 rounded-lg border border-slate-100">
                              <div className="text-xs font-bold text-slate-400 w-6 text-center">{index + 1}.</div>
                              
                              <select 
                                  className="border border-slate-300 rounded-md p-1.5 text-sm bg-white focus:ring-2 focus:ring-blue-500 outline-none w-40"
                                  value={cond.field}
                                  onChange={(e) => updateFilterCondition(cond.id, 'field', e.target.value)}
                              >
                                  <option value="Designation">Designation</option>
                                  <option value="Company">Company</option>
                                  <option value="Team Size">Team Size</option>
                                  <option value="DR Size">Direct Reports Size</option>
                                  <option value="Grade">Grade</option>
                              </select>

                              <select 
                                  className="border border-slate-300 rounded-md p-1.5 text-sm bg-white focus:ring-2 focus:ring-blue-500 outline-none w-32"
                                  value={cond.operator}
                                  onChange={(e) => updateFilterCondition(cond.id, 'operator', e.target.value)}
                              >
                                  {cond.field === 'Designation' && (
                                      <>
                                          <option value="contains">Contains</option>
                                          <option value="equals">Equals</option>
                                      </>
                                  )}
                                  {(cond.field === 'Team Size' || cond.field === 'DR Size') && (
                                      <>
                                          <option value=">=">{'>='}</option>
                                          <option value="<=">{'<='}</option>
                                          <option value="=">Equals</option>
                                          <option value=">">Greater than</option>
                                          <option value="<">Less than</option>
                                      </>
                                  )}
                                  {(cond.field === 'Grade' || cond.field === 'Company') && <option value="in">Includes</option>}
                              </select>

                              <div className="flex-1 min-w-[200px] relative">
                                  {cond.field === 'Designation' && (
                                      <input 
                                          type="text" 
                                          className="w-full border border-slate-300 rounded-md p-1.5 text-sm outline-none focus:ring-2 focus:ring-blue-500"
                                          placeholder="e.g. Manager"
                                          value={cond.value}
                                          onChange={(e) => updateFilterCondition(cond.id, 'value', e.target.value)}
                                      />
                                  )}
                                  {(cond.field === 'Team Size' || cond.field === 'DR Size') && (
                                      <input 
                                          type="number" 
                                          className="w-full border border-slate-300 rounded-md p-1.5 text-sm outline-none focus:ring-2 focus:ring-blue-500"
                                          placeholder="0"
                                          value={cond.value}
                                          onChange={(e) => updateFilterCondition(cond.id, 'value', e.target.value)}
                                      />
                                  )}
                                  {(cond.field === 'Grade' || cond.field === 'Company') && (
                                      <div className="relative">
                                          <button 
                                              onClick={() => setOpenDropdown(openDropdown === cond.id ? null : cond.id)}
                                              className="w-full border border-slate-300 rounded-md p-1.5 text-sm bg-white text-left flex justify-between items-center focus:ring-2 focus:ring-blue-500"
                                          >
                                              <span className="truncate">{cond.value.length > 0 ? `${cond.value.length} Selected` : `Select ${cond.field}s...`}</span>
                                              <ChevronDown size={16} className="text-slate-400 flex-shrink-0" />
                                          </button>
                                          {openDropdown === cond.id && (
                                              <>
                                                  <div className="fixed inset-0 z-40" onClick={() => setOpenDropdown(null)}></div>
                                                  <div className="absolute top-full left-0 mt-1 w-full max-h-60 overflow-y-auto bg-white border border-slate-200 shadow-xl rounded-md z-50 p-2 grid grid-cols-2 gap-2" style={{ scrollbarWidth: 'thin' }}>
                                                      {(cond.field === 'Grade' ? allUniqueGrades : allUniqueCompanies).map(item => (
                                                          <label key={item} className="flex items-center gap-2 text-sm p-1.5 hover:bg-slate-50 rounded cursor-pointer border border-transparent hover:border-slate-200 transition-colors">
                                                              <input 
                                                                  type="checkbox" 
                                                                  className="rounded text-blue-600 focus:ring-blue-500"
                                                                  checked={cond.value.includes(item)} 
                                                                  onChange={(e) => {
                                                                      const newVals = e.target.checked 
                                                                          ? [...cond.value, item] 
                                                                          : cond.value.filter(v => v !== item);
                                                                      updateFilterCondition(cond.id, 'value', newVals);
                                                                  }} 
                                                              />
                                                              <span className="truncate" title={item}>{item}</span>
                                                          </label>
                                                      ))}
                                                  </div>
                                              </>
                                          )}
                                      </div>
                                  )}
                              </div>

                              <button 
                                  onClick={() => removeFilterCondition(cond.id)}
                                  className="p-1.5 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-md transition-colors"
                                  title="Remove Condition"
                              >
                                  <Trash2 size={16} />
                              </button>
                          </div>
                      ))}
                      
                      <div className="flex gap-3 mt-4">
                          <button 
                              onClick={addFilterCondition}
                              className="flex items-center text-sm font-semibold text-blue-600 hover:text-blue-700 bg-blue-50 hover:bg-blue-100 px-3 py-1.5 rounded-md transition-colors"
                          >
                              <Plus size={16} className="mr-1" /> Add Condition
                          </button>
                          {filterConditions.length > 0 && (
                              <button 
                                  onClick={() => { setFilterConditions([]); setShowTabularResults(false); }}
                                  className="text-sm font-semibold text-slate-500 hover:text-slate-700 px-3 py-1.5"
                              >
                                  Clear All
                              </button>
                          )}
                      </div>
                  </div>
              </div>
          </div>
      )}

      {/* Main Content Area */}
      <main className="flex-1 overflow-auto relative" id="chart-container">
        
        {/* TABULAR VIEW (Overrides Org Chart if filters applied & explicitly shown) */}
        {appTab === 'org' && filterConditions.length > 0 && showTabularResults && (
            <div className="w-full max-w-6xl mx-auto p-4 sm:p-8 animate-fade-in-up">
                <div className="flex items-center justify-between mb-4">
                    <h2 className="text-xl font-bold text-slate-800">Filtered Results <span className="text-slate-400 font-normal ml-2">({tabularFilteredData.length})</span></h2>
                    <button 
                        onClick={() => setShowTabularResults(false)}
                        className="text-sm font-semibold text-blue-600 bg-blue-50 px-3 py-1.5 rounded hover:bg-blue-100 transition-colors"
                    >
                        Return to Org Chart
                    </button>
                </div>
                
                {tabularFilteredData.length === 0 ? (
                    <div className="bg-white p-10 rounded-xl border border-slate-200 text-center text-slate-500 shadow-sm">
                        No employees match your current filter conditions.
                    </div>
                ) : (
                    <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                        <div className="overflow-x-auto max-h-[65vh]" style={{ scrollbarWidth: 'thin' }}>
                            <table className="w-full text-left text-sm">
                                <thead className="text-slate-600 border-b border-slate-200 sticky top-0 z-20 shadow-sm bg-slate-50">
                                    <tr>
                                        <SortableHeader label="Employee" field="Employee" />
                                        <SortableHeader label="Grade" field="Grade" />
                                        <SortableHeader label="Job Title" field="JobTitle" />
                                        <SortableHeader label="Department" field="Department" />
                                        <SortableHeader label="DR Size" field="DRSize" align="center" />
                                        <SortableHeader label="Matrix Size" field="MatrixSize" align="center" />
                                        <SortableHeader label="Team Size" field="TeamSize" align="center" />
                                        <SortableHeader label="Line Manager" field="Manager" />
                                    </tr>
                                </thead>
                                <tbody>
                                    {tabularFilteredData.map((emp) => (
                                        <tr 
                                            key={emp._id} 
                                            className="border-b border-slate-100 hover:bg-blue-50 bg-white cursor-pointer transition-colors"
                                            onClick={() => {
                                                setActiveEmployeeId(emp._id);
                                                setShowTabularResults(false);
                                            }}
                                        >
                                            <td className="px-6 py-3">
                                                <div className="font-bold text-slate-800">{formatDisplayFirstLast(emp['Display Name'])}</div>
                                            </td>
                                            <td className="px-6 py-3">
                                                {emp._cardGrade && emp._cardGrade !== 'Unspecified' ? (
                                                    <span className="bg-slate-100 text-slate-600 px-2 py-1 rounded text-[10px] font-bold uppercase tracking-wider border border-slate-200 whitespace-nowrap">
                                                        {emp._cardGrade}
                                                    </span>
                                                ) : (
                                                    <span className="text-slate-400">-</span>
                                                )}
                                            </td>
                                            <td className="px-6 py-3 text-slate-700">
                                                {emp['Job Title']}
                                            </td>
                                            <td className="px-6 py-3 text-slate-600">{emp['Department (Label)']}</td>
                                            <td className="px-6 py-3 text-center font-medium text-slate-700">{emp._insights?.directCount || '-'}</td>
                                            <td className="px-6 py-3 text-center font-medium text-slate-700">{emp._insights?.matrixCount || '-'}</td>
                                            <td className="px-6 py-3 text-center font-medium text-slate-700">{emp._insights?.totalTeam || '-'}</td>
                                            <td className="px-6 py-3 text-slate-600">{formatDisplayFirstLast(emp['Line Manager Name']) || '-'}</td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                )}
            </div>
        )}

        {/* ORG CHART VIEW */}
        {appTab === 'org' && (!showTabularResults || filterConditions.length === 0) && (
            <div className="w-full max-w-7xl mx-auto min-h-full flex flex-col items-center pb-32 p-4 sm:p-8">
                {manager && (
                <div className="flex flex-col items-center animate-fade-in-down">
                    <EmployeeCard 
                        employee={manager} 
                        ceoId={ceoId}
                        globalMetrics={globalMetrics}
                        onClick={() => setActiveEmployeeId(manager._id)} 
                        onSelectDirect={() => { setActiveEmployeeId(manager._id); setViewMode('direct'); }}
                        onSelectMatrix={() => { setActiveEmployeeId(manager._id); setViewMode('matrix'); }}
                    />
                    <div className="h-10 w-px bg-slate-300 my-2"></div>
                </div>
                )}

                {activeEmployee && (
                <div className="relative flex justify-center items-center my-4 animate-scale-in">
                    <EmployeeCard 
                        employee={activeEmployee} 
                        ceoId={ceoId}
                        globalMetrics={globalMetrics}
                        isActive 
                        viewMode={viewMode}
                        onSelectDirect={() => setViewMode('direct')}
                        onSelectMatrix={() => setViewMode('matrix')}
                    />
                </div>
                )}

                {/* TEAMS STYLE MULTI-LINE LAYOUT */}
                {displayedReports.length > 0 && (
                <div className="flex flex-col items-center animate-fade-in-up w-full mt-2">
                    {/* Single Vertical Drop Line */}
                    <div className={`h-6 w-px ${isMatrixView ? 'bg-purple-400' : 'bg-slate-300'}`}></div>
                    
                    {/* Visual Label for the Group */}
                    <div className={`mb-6 text-[10px] font-bold uppercase tracking-wider px-4 py-1.5 rounded-full shadow-sm border ${isMatrixView ? 'bg-purple-50 text-purple-700 border-purple-200' : 'bg-white text-slate-500 border-slate-200'}`}>
                       {isMatrixView ? 'Matrix Reports' : 'Direct Reports'} ({displayedReports.length})
                    </div>
                    
                    {/* Wrapping Grid Container */}
                    <div className="flex justify-center flex-wrap gap-6 w-full max-w-6xl px-4">
                    {displayedReports.map(emp => (
                        <div key={emp._id} className="flex flex-col items-center w-64">
                        <EmployeeCard 
                            employee={emp}
                            ceoId={ceoId}
                            globalMetrics={globalMetrics}
                            isMatrixNode={isMatrixView}
                            onClick={() => setActiveEmployeeId(emp._id)} 
                            onSelectDirect={() => {
                                setActiveEmployeeId(emp._id);
                                setViewMode('direct');
                            }}
                            onSelectMatrix={() => {
                                setActiveEmployeeId(emp._id);
                                setViewMode('matrix');
                            }}
                        />
                        </div>
                    ))}
                    </div>
                </div>
                )}
            </div>
        )}

        {/* ATTRITION VIEW */}
        {appTab === 'attrition' && attritionSummary && (
            <div className="w-full max-w-5xl mx-auto p-4 sm:p-8 animate-fade-in-up">
                <div className="flex items-center justify-between mb-6">
                    <div>
                        <h2 className="text-2xl font-bold text-slate-800">Attrition Dashboard</h2>
                        <p className="text-slate-500 text-sm mt-1">Cascading exit data across the organizational hierarchy</p>
                    </div>
                    <div className="bg-blue-50 text-blue-700 px-4 py-2 rounded-lg font-bold border border-blue-100 flex items-center">
                        <TrendingDown size={18} className="mr-2"/> Global Total: {attritionSummary.total}
                    </div>
                </div>

                {/* KPI Cards */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-8">
                    <div className="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col">
                        <span className="text-sm font-semibold text-slate-500 uppercase tracking-wider mb-2">Voluntary</span>
                        <div className="flex items-end gap-3">
                            <span className="text-3xl font-bold text-orange-600">{attritionSummary.vol}</span>
                            <span className="text-sm text-slate-400 mb-1 pb-0.5">employees</span>
                        </div>
                    </div>
                    <div className="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col">
                        <span className="text-sm font-semibold text-slate-500 uppercase tracking-wider mb-2">Involuntary</span>
                        <div className="flex items-end gap-3">
                            <span className="text-3xl font-bold text-slate-600">{attritionSummary.invol}</span>
                            <span className="text-sm text-slate-400 mb-1 pb-0.5">employees</span>
                        </div>
                    </div>
                    <div className="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col">
                        <span className="text-sm font-semibold text-slate-500 uppercase tracking-wider mb-2">Retirement</span>
                        <div className="flex items-end gap-3">
                            <span className="text-3xl font-bold text-blue-600">{attritionSummary.ret}</span>
                            <span className="text-sm text-slate-400 mb-1 pb-0.5">employees</span>
                        </div>
                    </div>
                </div>

                {/* Manager-wise Cascading Table */}
                <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                    <div className="px-6 py-4 border-b border-slate-200 bg-slate-50">
                        <h3 className="font-bold text-slate-700">Manager Team Attrition Hierarchy</h3>
                    </div>
                    <div className="overflow-x-auto pb-32">
                        <table className="w-full text-left text-sm">
                            <thead className="bg-white text-slate-500 border-b">
                                <tr>
                                    <th className="py-3 font-semibold w-12 pl-6"></th>
                                    <th className="px-6 py-3 font-semibold">Manager Name</th>
                                    <th className="px-6 py-3 font-semibold">Department</th>
                                    <th className="px-6 py-3 font-semibold text-center">Current Team Size</th>
                                    <th className="px-6 py-3 font-semibold text-center">Direct Attrition</th>
                                    <th className="px-6 py-3 font-semibold text-center">Total Attrition</th>
                                </tr>
                            </thead>
                            <tbody>
                                {renderManagerRows(attritionSummary.tree)}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        )}

      </main>

      {/* Direct Attritions Detail Modal Overlay */}
      {attritionModal && (
        <div className="fixed inset-0 bg-slate-900/60 z-[99999] flex items-center justify-center p-4 animate-fade-in-down" onClick={() => setAttritionModal(null)}>
            <div className="bg-white rounded-xl shadow-2xl w-full max-w-3xl max-h-[85vh] flex flex-col" onClick={e => e.stopPropagation()}>
                <div className="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-xl">
                    <div>
                        <h3 className="font-bold text-slate-800 text-lg">Direct Attritions</h3>
                        <p className="text-sm text-slate-500">Former direct reports of <span className="font-semibold text-slate-700">{attritionModal.name}</span></p>
                    </div>
                    <button onClick={() => setAttritionModal(null)} className="text-slate-400 hover:text-slate-600 bg-white hover:bg-slate-200 p-1.5 rounded-full transition-colors">
                        <X size={20} />
                    </button>
                </div>
                <div className="p-6 overflow-y-auto" style={{ scrollbarWidth: 'thin' }}>
                    <table className="w-full text-sm text-left">
                        <thead>
                            <tr className="text-slate-400 border-b border-slate-200">
                                <th className="pb-3 font-semibold">Employee Name</th>
                                <th className="pb-3 font-semibold">Job Title</th>
                                <th className="pb-3 font-semibold">Exit Type</th>
                                <th className="pb-3 font-semibold text-right">Last Working Date</th>
                            </tr>
                        </thead>
                        <tbody>
                            {attritionModal.directAttritions.map((rec, i) => {
                                const reasonLow = String(rec.reason).toLowerCase();
                                let badgeClass = 'bg-slate-100 text-slate-700 border-slate-200';
                                
                                if (reasonLow.includes('voluntary') && !reasonLow.includes('in-voluntary') && !reasonLow.includes('involuntary')) {
                                    badgeClass = 'bg-orange-50 text-orange-700 border-orange-200';
                                } else if (reasonLow.includes('involuntary') || reasonLow.includes('in-voluntary')) {
                                    badgeClass = 'bg-slate-100 text-slate-700 border-slate-300';
                                } else if (reasonLow.includes('retire')) {
                                    badgeClass = 'bg-blue-50 text-blue-700 border-blue-200';
                                }

                                return (
                                    <tr key={i} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 transition-colors">
                                        <td className="py-3 font-semibold text-slate-700">{rec.name}</td>
                                        <td className="py-3 text-slate-600">{rec.title}</td>
                                        <td className="py-3">
                                            <span className={`px-2 py-0.5 rounded text-[10px] font-bold uppercase tracking-wider border ${badgeClass}`}>
                                                {rec.reason}
                                            </span>
                                        </td>
                                        <td className="py-3 text-right text-slate-500 font-medium">{formatDateUI(rec.lwd)}</td>
                                    </tr>
                                );
                            })}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
      )}
    </div>
  );
}

function EmployeeCard({ employee, ceoId, globalMetrics, isActive, isMatrixNode, viewMode, onClick, onSelectDirect, onSelectMatrix }) {
  const [showTooltip, setShowTooltip] = useState(false);
  const [showTimeline, setShowTimeline] = useState(false);
  
  const [gradeTooltip, setGradeTooltip] = useState(null); // 'direct', 'matrix', or 'team'
  const [tooltipPos, setTooltipPos] = useState({ h: 'right', v: 'top' });
  
  const hideTimeout = useRef(null);
  const hideGradeTimeout = useRef(null);

  const handleMouseEnterInfo = (e) => {
      clearTimeout(hideTimeout.current);
      const iconRect = e.currentTarget.getBoundingClientRect();
      
      const tooltipWidth = 360; 
      const tooltipHeight = 450; 
      
      let h = (iconRect.right + tooltipWidth + 20 > window.innerWidth) ? 'left' : 'right';
      let v = (iconRect.top + tooltipHeight > window.innerHeight) ? 'bottom' : 'top';
      
      setTooltipPos({ h, v });
      setShowTooltip(true);
  };

  const handleMouseLeaveInfo = () => {
    hideTimeout.current = setTimeout(() => setShowTooltip(false), 200);
  };

  const handleMouseEnterGrade = (e, type) => {
    clearTimeout(hideGradeTimeout.current);
    const rect = e.currentTarget.getBoundingClientRect();
    
    let h = (type === 'team') ? 'right' : 'left';
    if (h === 'left' && rect.left - 200 < 0) h = 'right';
    if (h === 'right' && rect.right + 200 > window.innerWidth) h = 'left';

    const v = (rect.top + 250 > window.innerHeight) ? 'bottom' : 'top';
    
    setTooltipPos({ h, v });
    setGradeTooltip(type);
  };

  const handleMouseLeaveGrade = () => {
      hideGradeTimeout.current = setTimeout(() => setGradeTooltip(null), 200);
  };

  const initials = String(employee['Display Name'] || '?').split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
  const displayName = formatDisplayFirstLast(employee['Display Name']);
  const insights = employee._insights || { genderCount: { male: 0, female: 0, other: 0 }};
  
  const tenureFormatted = formatDuration(employee._tenureDate);
  const timeInRoleFormatted = formatDuration(employee._timeInRoleDate || employee._tenureDate); 
  const lastPromotionFormatted = formatDuration(employee._lastPromotionDate); 
  const timeWithManagerFormatted = formatDuration(employee._timeWithManagerDate || employee._tenureDate);

  const totalGender = insights.genderCount.male + insights.genderCount.female + insights.genderCount.other;
  const malePct = totalGender > 0 ? Math.round((insights.genderCount.male / totalGender) * 100) : 0;
  const femalePct = totalGender > 0 ? Math.round((insights.genderCount.female / totalGender) * 100) : 0;
  const isIndividualContributor = insights.directCount === 0 && insights.matrixCount === 0;

  const isFirstLevel = !employee._managerId || employee._managerId === ceoId;

  let cardClasses = "w-full bg-white rounded-xl shadow-md border p-4 transition-all duration-200 flex flex-col group ";
  
  if (isActive) {
    cardClasses += "border-blue-500 ring-4 ring-blue-100 shadow-xl scale-105 cursor-default";
  } else if (isMatrixNode) {
    cardClasses += "border-purple-300 border-dashed hover:border-purple-500 hover:shadow-lg cursor-pointer";
  } else {
    cardClasses += "border-slate-200 hover:border-blue-400 hover:shadow-lg cursor-pointer";
  }

  let popupHeaderClass = "px-3 py-2 border-b text-xs font-bold uppercase tracking-wider flex justify-between ";
  if (gradeTooltip === 'direct') popupHeaderClass += "bg-blue-100 text-blue-800 border-blue-200";
  else if (gradeTooltip === 'matrix') popupHeaderClass += "bg-purple-100 text-purple-800 border-purple-200";
  else popupHeaderClass += "bg-slate-100 text-slate-700 border-slate-200";

  return (
    <div className={`relative w-full ${(showTooltip || gradeTooltip) ? 'z-[99999]' : isActive ? 'z-10' : 'z-0'}`}>
      <div className={cardClasses} onClick={!isActive ? onClick : undefined}>
        
        {/* Bookmark Tags for ExCom / OpCom */}
        {employee._isExCom && (
            <div className="absolute top-0 right-10 z-20 drop-shadow-md" title="Executive Committee">
                <div className="bg-amber-100 text-amber-800 text-xs font-extrabold px-2 pt-1 pb-2.5" style={{ clipPath: 'polygon(0 0, 100% 0, 100% 100%, 50% 75%, 0 100%)' }}>
                    Ex
                </div>
            </div>
        )}
        {employee._isOpCom && (
            <div className="absolute top-0 right-10 z-20 drop-shadow-md" title="Operations Committee">
                <div className="bg-indigo-100 text-indigo-800 text-xs font-extrabold px-2 pt-1 pb-2.5" style={{ clipPath: 'polygon(0 0, 100% 0, 100% 100%, 50% 75%, 0 100%)' }}>
                    Op
                </div>
            </div>
        )}

        <div 
          className="absolute top-3 right-3 text-slate-400 hover:text-blue-600 z-20 cursor-help"
          onMouseEnter={handleMouseEnterInfo}
          onMouseLeave={handleMouseLeaveInfo}
          onClick={(e) => e.stopPropagation()}
        >
          <Info size={18} />
        </div>

        <div className="flex items-center space-x-3 mb-3 pr-6">
          <div className={`w-12 h-12 rounded-full flex-shrink-0 flex items-center justify-center text-white font-bold shadow-sm ${isActive ? 'bg-blue-600' : isMatrixNode ? 'bg-purple-500' : 'bg-slate-700'}`}>
            {initials}
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="font-bold text-slate-800 truncate text-sm" title={employee['Display Name']}>
              {displayName}
            </h3>
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
              <div className="flex items-center space-x-1 text-slate-500 flex-shrink-0 font-medium" title="Tenure"><Clock size={12} /> <span>{tenureFormatted}</span></div>
          </div>
        </div>

        {/* Bottom Row Counters: Direct, Matrix, Team */}
        <div className="mt-3 flex justify-between items-center text-[10px] font-semibold pt-2 border-t text-slate-600">
            <div 
                onMouseEnter={(e) => handleMouseEnterGrade(e, 'direct')}
                onMouseLeave={handleMouseLeaveGrade}
                onClick={(e) => { e.stopPropagation(); if(onSelectDirect) onSelectDirect(); }}
                className={`flex items-center px-1 py-0.5 rounded transition-colors cursor-pointer ${isActive && viewMode === 'direct' ? 'bg-blue-100 text-blue-800 ring-1 ring-blue-300' : 'hover:bg-blue-50 text-slate-600'}`}
            >
                <UserCircle2 size={12} className={`mr-1 ${isActive && viewMode === 'direct' ? 'text-blue-600' : 'text-blue-500'}`}/> {insights.directCount} Direct
            </div>
            
            {insights.matrixCount > 0 && (
                <div 
                    onMouseEnter={(e) => handleMouseEnterGrade(e, 'matrix')}
                    onMouseLeave={handleMouseLeaveGrade}
                    onClick={(e) => { e.stopPropagation(); if(onSelectMatrix) onSelectMatrix(); }} 
                    className={`flex items-center px-1 py-0.5 rounded transition-colors cursor-pointer ${isActive && viewMode === 'matrix' ? 'bg-purple-100 text-purple-700 ring-1 ring-purple-300' : 'hover:bg-purple-50 text-purple-600'}`}
                >
                    <span>{insights.matrixCount} Matrix</span>
                </div>
            )}

            <div 
                onMouseEnter={(e) => handleMouseEnterGrade(e, 'team')}
                onMouseLeave={handleMouseLeaveGrade}
                className="flex items-center cursor-help px-1 py-0.5 hover:bg-slate-100 rounded"
            >
                <Users size={12} className="mr-1 text-slate-500"/> {insights.totalTeam} Team
            </div>
        </div>
      </div>

      {/* Pay Grade Details Popup Component using Absolute Layout Relative to the Card */}
      {gradeTooltip && (
          <div 
              className={`absolute w-48 bg-white rounded-lg shadow-[0_0_20px_rgba(0,0,0,0.15)] border border-slate-200 text-sm overflow-hidden animate-scale-in z-[99999] ${tooltipPos.h === 'right' ? 'left-full ml-3' : 'right-full mr-3'} ${tooltipPos.v === 'top' ? 'top-10' : 'bottom-10'}`}
              onMouseEnter={() => clearTimeout(hideGradeTimeout.current)}
              onMouseLeave={handleMouseLeaveGrade}
          >
              <div className={popupHeaderClass}>
                  <span>
                    {gradeTooltip === 'direct' ? 'DR Summary' : gradeTooltip === 'matrix' ? 'Matrix Summary' : 'Team Summary'}
                  </span>
              </div>
              <div className="p-2 max-h-64 overflow-y-auto" style={{ scrollbarWidth: 'thin' }}>
                  {gradeTooltip === 'direct' && renderGradesList(insights.directGrades)}
                  {gradeTooltip === 'matrix' && renderGradesList(insights.matrixGrades)}
                  {gradeTooltip === 'team' && renderGradesList(insights.teamGrades)}
              </div>
          </div>
      )}

      {/* Advanced Insights Tooltip (Using dynamically assigned CSS absolute classes for stability) */}
      {showTooltip && (
        <div 
          className={`absolute w-[360px] bg-white rounded-xl shadow-[0_0_40px_rgba(0,0,0,0.2)] border border-slate-200 p-0 text-sm overflow-hidden animate-scale-in z-[99999] flex flex-col ${tooltipPos.h === 'right' ? 'left-full ml-3' : 'right-full mr-3'} ${tooltipPos.v === 'top' ? 'top-0' : 'bottom-0'}`}
          style={{ maxHeight: '90vh' }}
          onMouseEnter={() => clearTimeout(hideTimeout.current)}
          onMouseLeave={handleMouseLeaveInfo}
        >
          <div className="bg-slate-800 text-white px-5 py-4 border-b flex items-center flex-shrink-0">
            <Info size={18} className="mr-2" />
            <span className="font-bold text-base">Advanced Insights</span>
          </div>
          
          <div className="p-5 space-y-6 overflow-y-auto" style={{ scrollbarWidth: 'thin' }}>

            {/* --- INDIVIDUAL CONTEXT --- */}
            <div>
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-3 pb-1 border-b border-slate-100">Individual Context</h4>
                <div className="grid grid-cols-2 gap-3 text-sm mb-4">
                    <div className="bg-slate-50 p-2.5 rounded-lg border border-slate-100 flex flex-col items-center text-center">
                        <span className="text-slate-400 font-medium mb-1 text-xs">Total Tenure</span>
                        <span className="font-bold text-slate-700 flex items-center"><CalendarDays size={14} className="mr-1.5"/> {tenureFormatted}</span>
                    </div>
                    <div className="bg-slate-50 p-2.5 rounded-lg border border-slate-100 flex flex-col items-center text-center">
                        <span className="text-slate-400 font-medium mb-1 text-xs">Time in Role</span>
                        <span className="font-bold text-slate-700 flex items-center"><Clock size={14} className="mr-1.5"/> {timeInRoleFormatted}</span>
                    </div>
                    <div className="bg-slate-50 p-2.5 rounded-lg border border-slate-100 flex flex-col items-center text-center">
                        <span className="text-slate-400 font-medium mb-1 text-xs">Since Promoted</span>
                        <span className={`font-bold flex items-center ${lastPromotionFormatted !== '-' ? 'text-green-700' : 'text-slate-400'}`}><Clock size={14} className="mr-1.5"/> {lastPromotionFormatted}</span>
                    </div>
                    <div className="bg-slate-50 p-2.5 rounded-lg border border-slate-100 flex flex-col items-center text-center">
                        <span className="text-slate-400 font-medium mb-1 text-xs">With Manager</span>
                        <span className="font-bold text-indigo-700 flex items-center"><Users size={14} className="mr-1.5"/> {timeWithManagerFormatted}</span>
                    </div>
                </div>

                {/* Expandable Role History Timeline */}
                {employee._timeline && employee._timeline.length > 0 && (
                    <div className="mb-2">
                        <button
                            className="text-sm text-blue-600 hover:text-blue-700 flex items-center font-bold w-full focus:outline-none transition-colors"
                            onClick={(e) => { e.stopPropagation(); setShowTimeline(!showTimeline); }}
                        >
                            {showTimeline ? 'Hide Role History' : 'View Role History'}
                        </button>
                        
                        {showTimeline && (
                            <div className="mt-4 relative ml-3 space-y-0 pb-1">
                                {employee._timeline.map((item, i) => {
                                    if (item.isGap) {
                                        return (
                                            <div key={i} className="relative pl-7 py-3">
                                                <div className="absolute left-0 top-0 bottom-0 w-px border-l-2 border-dashed border-slate-300"></div>
                                                <div className="text-xs italic text-slate-500 flex items-center bg-white relative z-10 py-1 -ml-3 pl-3">
                                                    Data Unavailable <span className="mx-2">•</span> {item.duration} gap
                                                </div>
                                            </div>
                                        );
                                    }

                                    let dotColor = 'bg-slate-400 ring-white';
                                    if (item.type === 'promotion') dotColor = 'bg-green-500 ring-white';
                                    else if (item.type === 'transfer') dotColor = 'bg-blue-500 ring-white';
                                    else if (item.type === 'hire') dotColor = 'bg-purple-500 ring-white';
                                    else if (item.type === 'company_transfer') dotColor = 'bg-orange-500 ring-white';

                                    const showLocation = employee._hasLocationTransfer && item.location && item.location !== employee['Location Name'];

                                    return (
                                        <div key={i} className="relative pl-7 pb-5">
                                            {i !== employee._timeline.length - 1 && (
                                                <div className="absolute left-0 top-2 bottom-0 w-px bg-slate-200 -ml-[1px]"></div>
                                            )}
                                            <div className={`absolute -left-[6px] top-1.5 h-3 w-3 rounded-full ring-[3px] ${dotColor} z-10`}></div>
                                            
                                            <div className="text-sm font-bold text-slate-800 leading-tight flex flex-wrap gap-2 items-center">
                                                <span>{item.title || 'Unknown Title'}</span>
                                                {item.grade && item.grade !== 'Unspecified' && (
                                                    <span className="text-slate-500 font-normal border border-slate-200 px-1.5 rounded text-[10px] bg-slate-50">
                                                        {item.grade}
                                                    </span>
                                                )}
                                            </div>
                                            <div className="text-xs text-slate-500 mt-1.5 flex items-center gap-3 flex-wrap">
                                                <span className="flex items-center"><Clock size={12} className="mr-1.5"/> {item.durationFormatted}</span>
                                                {showLocation && (
                                                    <span className="flex items-center text-blue-600 font-medium"><MapPin size={12} className="mr-1.5"/> {item.location}</span>
                                                )}
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
               <div className="pt-4 text-center border-t border-slate-100">
                   <p className="text-base font-semibold text-slate-600">Individual Contributor</p>
                   <p className="text-sm text-slate-400 mt-1">No direct or matrix reports.</p>
               </div>
            ) : (
              <div>
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-4 pt-1 border-t border-slate-100">Organizational Context</h4>
                
                {/* 1. EXCOM BENCHMARK (Overrides standard peer benchmark for ExComs) */}
                {employee._isExCom ? (
                    globalMetrics.exCom && (
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
                    )
                ) : (
                    <>
                        {/* 2. STANDARD PEER BENCHMARK */}
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
                        
                        {/* 3. OPCOM BENCHMARK (For qualifying VPs/GMs) */}
                        {employee._isOpCom && globalMetrics.opCom && (
                            <BenchmarkBox title="OpCom Benchmark" borderColor="border-indigo-200" titleColor="text-indigo-600" bgClass="bg-indigo-50/20">
                                <MetricScale label="Direct Reports" min={globalMetrics.opCom.drMin} max={globalMetrics.opCom.drMax} median={globalMetrics.opCom.drMedian} value={insights.directCount} />
                                <MetricScale label="Total Team Size" min={globalMetrics.opCom.teamMin} max={globalMetrics.opCom.teamMax} median={globalMetrics.opCom.teamMedian} value={insights.totalTeam} />
                            </BenchmarkBox>
                        )}
                        
                        {/* 4. GRADE BENCHMARK */}
                        {!isFirstLevel && globalMetrics.grade && globalMetrics.grade[employee._summaryGrade] && (
                            <BenchmarkBox 
                                title="Grade Benchmark" 
                                rightElement={
                                    <div className="flex items-center space-x-1.5 text-slate-500 font-bold bg-slate-100 px-2 py-0.5 rounded text-xs">
                                        <Award size={12} className="flex-shrink-0"/> <span>{employee._summaryGrade}</span>
                                    </div>
                                }
                            >
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
        </div>
      )}
    </div>
  );
}