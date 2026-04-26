import React, { useState, useMemo, useEffect, useRef } from 'react';
import { Upload, Search, Info, Users, MapPin, Building2, UserCircle2 } from 'lucide-react';

// --- Helper for formatting grades ---
const formatGrades = (gradesObj) => {
    if (!gradesObj) return "No data";
    const entries = Object.entries(gradesObj);
    if(entries.length === 0) return "No data";
    // Sort by count descending
    return entries.sort((a,b) => b[1] - a[1]).map(([g, c]) => `${c} ${g}`).join(' | ');
};

// --- Helper for name normalization (fuzzy matching) ---
const normalizeName = (name) => name ? name.toLowerCase().replace(/\s+/g, ' ').trim() : '';
const getFirstLast = (name) => {
    if (!name) return '';
    const parts = normalizeName(name).split(' ');
    if (parts.length <= 1) return parts[0];
    return `${parts[0]} ${parts[parts.length - 1]}`;
};

export default function App() {
  const [data, setData] = useState([]);
  const [employeeMap, setEmployeeMap] = useState({});
  const [activeEmployeeId, setActiveEmployeeId] = useState(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [xlsxLoaded, setXlsxLoaded] = useState(false);
  const [viewMode, setViewMode] = useState('direct'); // 'direct' or 'matrix'

  // Load XLSX dynamically
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

  // Reset view mode when active employee changes naturally
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
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = window.XLSX.utils.sheet_to_json(worksheet, { defval: "" });

      if (jsonData.length === 0) {
        throw new Error("The uploaded Excel file is empty.");
      }

      processEmployeeData(jsonData);
    } catch (err) {
      setError(err.message || "Failed to process the Excel file. Please check the format.");
    } finally {
      setLoading(false);
    }
  };

  const processEmployeeData = (rawData) => {
    const empMap = {};
    const directReportsMap = {};
    const matrixReportsMap = {};
    const exactNameMap = {};
    const firstLastNameMap = {};

    // 1. Initial cleanup and strict ID mapping with fuzzy name fallback
    rawData.forEach(emp => {
      // Prioritize robust ID columns if they exist
      const empId = (emp['Employee ID'] || emp['Users Sys Id'])?.toString().trim();
      const displayName = emp['Display Name']?.trim();
      
      const id = empId || displayName;
      if (!id) return;
      
      emp._id = id;
      empMap[id] = { ...emp };

      if (empId) {
          exactNameMap[empId.toLowerCase()] = id;
      }
      if (displayName) {
          exactNameMap[normalizeName(displayName)] = id;
          firstLastNameMap[getFirstLast(displayName)] = id;
      }
    });

    // Advanced lookup function
    const findEmployeeId = (userId, nameQuery) => {
        // First attempt exact ID match if provided
        if (userId) {
            const idStr = userId.toString().trim().toLowerCase();
            if (exactNameMap[idStr]) return exactNameMap[idStr];
        }
        // Fallback to name match
        if (nameQuery) {
            const norm = normalizeName(nameQuery);
            if (exactNameMap[norm]) return exactNameMap[norm];
            const fl = getFirstLast(nameQuery);
            if (firstLastNameMap[fl]) return firstLastNameMap[fl];
        }
        return null;
    };

    // 2. Identify CEO and fix Line Managers
    let ceos = Object.values(empMap).filter(emp => 
        (!emp['Line Manager Name'] || emp['Line Manager Name'].trim() === '') &&
        (!emp['Line Manager UserID'] || emp['Line Manager UserID'].toString().trim() === '')
    );
    
    let actualCEO = null;
    if (ceos.length > 1) {
      actualCEO = ceos.find(c => {
        const title = (c['Job Title'] || c['Designation'] || '').toLowerCase();
        return title.includes('ceo') || title.includes('chief') || title.includes('managing director');
      });
      if (!actualCEO) actualCEO = ceos[0];
    } else if (ceos.length === 1) {
      actualCEO = ceos[0];
    }

    // Assign matrix managers to the 'other' empty line managers
    if (actualCEO) {
      ceos.forEach(c => {
        if (c._id !== actualCEO._id) {
          empMap[c._id]['Matrix Manager Name'] = actualCEO['Display Name'];
          empMap[c._id]['Matrix Manager ID'] = actualCEO['Employee ID'] || actualCEO['Users Sys Id'];
        }
      });
    }

    // 3. Build Adjacency Lists
    Object.values(empMap).forEach(emp => {
      const managerId = findEmployeeId(emp['Line Manager UserID'], emp['Line Manager Name']);
      const matrixId = findEmployeeId(emp['Matrix Manager ID'] || emp['Matrix Manager UserID'], emp['Matrix Manager Name']);

      if (managerId && managerId !== emp._id) {
        emp._managerId = managerId;
        if (!directReportsMap[managerId]) directReportsMap[managerId] = [];
        directReportsMap[managerId].push(emp._id);
      }

      if (matrixId && matrixId !== emp._id) {
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

    // 4. Calculate Insights (Bottom-Up using DFS)
    const calculateInsights = (empId, visited = new Set()) => {
      if (visited.has(empId)) return empMap[empId]._insights;
      visited.add(empId);

      const directs = directReportsMap[empId] || [];
      const matrix = matrixReportsMap[empId] || [];
      
      let totalTeam = 0;
      let genderCount = { male: 0, female: 0, other: 0 };
      let probationCount = 0;
      let noticeCount = 0;

      const directGrades = {};
      const matrixGrades = {};
      const teamGrades = {};

      directs.forEach(childId => {
        const child = empMap[childId];
        if (!child) return;

        const grade = child['Grade Level (Picklist Label)'] || 'Unspecified';
        directGrades[grade] = (directGrades[grade] || 0) + 1;
        teamGrades[grade] = (teamGrades[grade] || 0) + 1;

        const gender = (child['Gender'] || '').toLowerCase();
        if (gender.startsWith('m')) genderCount.male++;
        else if (gender.startsWith('f')) genderCount.female++;
        else genderCount.other++;

        const empStatus = String(child['Employee Status (Picklist Label)'] || '').toLowerCase();
        const confStatus = String(child['Confirmation Status (Picklist Label)'] || '').toLowerCase();
        
        if (empStatus.includes('probation') || confStatus.includes('probation')) {
            probationCount++;
        }

        const resDate = String(child['Resignation Date'] || '').trim();
        const hasResigned = resDate.length > 0 && resDate.toLowerCase() !== 'na' && resDate.toLowerCase() !== 'null' && resDate !== '-';
        
        if (empStatus.includes('notice') || hasResigned) {
            noticeCount++;
        }

        const childInsights = calculateInsights(childId, visited);
        if (childInsights) {
          totalTeam += 1 + childInsights.totalTeam;
          addGrades(teamGrades, childInsights.teamGrades);
        }
      });

      matrix.forEach(childId => {
          const child = empMap[childId];
          if(!child) return;
          const grade = child['Grade Level (Picklist Label)'] || 'Unspecified';
          matrixGrades[grade] = (matrixGrades[grade] || 0) + 1;
      });

      const insights = {
        directCount: directs.length,
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

    // 5. Calculate Cross-Peer / Manager-Relative Insights
    Object.values(empMap).forEach(emp => {
      const managerId = emp._managerId;
      if (managerId && empMap[managerId]) {
        const manager = empMap[managerId];
        const peers = (manager._directs || []).filter(id => id !== emp._id);
        
        const myBranchSize = 1 + (emp._insights?.totalTeam || 0);
        const managerTeamSize = manager._insights?.totalTeam || 1;
        
        if (managerTeamSize > 0) {
           emp._insights.pctOfManagerTeam = Math.round((myBranchSize / managerTeamSize) * 100);
        }
        
        if (peers.length > 0) {
          const totalPeerDirects = peers.reduce((sum, peerId) => sum + (empMap[peerId]?._insights?.directCount || 0), 0);
          emp._insights.peerAvgDirects = (totalPeerDirects / peers.length).toFixed(1);
        } else {
          emp._insights.peerAvgDirects = 0;
        }
      }
    });

    setData(Object.values(empMap));
    setEmployeeMap(empMap);
    
    if (actualCEO) setActiveEmployeeId(actualCEO._id);
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

  const filteredSearch = useMemo(() => {
    if (!searchQuery) return [];
    return data
      .filter(emp => emp['Display Name']?.toLowerCase().includes(searchQuery.toLowerCase()))
      .slice(0, 5);
  }, [searchQuery, data]);

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
            <p className="text-slate-500">
              Drag and drop your Excel (.xlsx) file here or click to browse.<br/>
              Ensure it contains standard employee details and manager columns.
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
    .sort((a, b) => (b._insights?.totalTeam || 0) - (a._insights?.totalTeam || 0));
  const matrixReports = (activeEmployee?._matrix || [])
    .map(id => employeeMap[id])
    .filter(Boolean)
    .sort((a, b) => (b._insights?.totalTeam || 0) - (a._insights?.totalTeam || 0));

  const isMatrixView = viewMode === 'matrix';
  const displayedReports = isMatrixView ? matrixReports : directReports;

  return (
    <div className="h-screen w-full flex flex-col font-sans text-slate-800 overflow-hidden bg-slate-50">
      {/* Header */}
      <header className="bg-white border-b px-6 py-4 flex items-center justify-between shadow-sm sticky top-0 z-20 flex-shrink-0">
        <div className="flex items-center space-x-3">
          <div className="bg-blue-600 p-2 rounded-lg"><Users className="text-white" size={24} /></div>
          <h1 className="text-xl font-bold text-slate-800">Org Chart & Insights</h1>
        </div>

        {/* Search Bar */}
        <div className="relative w-80">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-slate-400" size={18} />
            <input 
              type="text" 
              placeholder="Search employee by name..." 
              className="w-full pl-10 pr-4 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 bg-slate-50"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
            />
          </div>
          {searchQuery && (
            <div className="absolute top-full left-0 right-0 mt-2 bg-white rounded-lg shadow-xl border overflow-hidden z-50">
              {filteredSearch.length > 0 ? (
                filteredSearch.map(emp => (
                  <button
                    key={emp._id}
                    className="w-full text-left px-4 py-3 hover:bg-slate-50 border-b last:border-0 flex flex-col"
                    onClick={() => {
                      setActiveEmployeeId(emp._id);
                      setSearchQuery('');
                    }}
                  >
                    <span className="font-semibold text-slate-800">{emp['Display Name']}</span>
                    <span className="text-xs text-slate-500">{emp['Job Title']} • {emp['Department (Label)']}</span>
                  </button>
                ))
              ) : (
                <div className="px-4 py-3 text-slate-500 text-sm">No employees found.</div>
              )}
            </div>
          )}
        </div>
      </header>

      {/* Main Org Chart Area */}
      <main className="flex-1 overflow-y-auto overflow-x-hidden p-4 sm:p-8 relative" id="chart-container">
        <div className="w-full max-w-7xl mx-auto min-h-full flex flex-col items-center pb-32">
            
            {manager && (
            <div className="flex flex-col items-center animate-fade-in-down">
                <EmployeeCard 
                    employee={manager} 
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
      </main>
    </div>
  );
}

function EmployeeCard({ employee, isActive, isMatrixNode, viewMode, onClick, onSelectDirect, onSelectMatrix }) {
  const [showTooltip, setShowTooltip] = useState(false);
  const [tooltipPos, setTooltipPos] = useState({ h: 'right', v: 'top' });
  const hideTimeout = useRef(null);

  const handleMouseEnter = (e) => {
    clearTimeout(hideTimeout.current);
    const iconRect = e.currentTarget.getBoundingClientRect();
    
    // Evaluate horizontal bounds to decide which side of the card it opens on
    // Tooltip is roughly 300px wide
    const h = (iconRect.right + 300 > window.innerWidth) ? 'left' : 'right';
    
    // Evaluate vertical bounds to prevent it from going past the bottom edge
    // Tooltip is roughly 350px tall
    const v = (iconRect.top + 350 > window.innerHeight) ? 'bottom' : 'top';
    
    setTooltipPos({ h, v });
    setShowTooltip(true);
  };

  const handleMouseLeave = () => {
    hideTimeout.current = setTimeout(() => {
      setShowTooltip(false);
    }, 200);
  };

  const initials = (employee['Display Name'] || '?').split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
  const insights = employee._insights || { genderCount: { male: 0, female: 0, other: 0 }};
  
  const totalGender = insights.genderCount.male + insights.genderCount.female + insights.genderCount.other;
  const malePct = totalGender > 0 ? Math.round((insights.genderCount.male / totalGender) * 100) : 0;
  const femalePct = totalGender > 0 ? Math.round((insights.genderCount.female / totalGender) * 100) : 0;

  let cardClasses = "relative w-full bg-white rounded-xl shadow-md border p-4 transition-all duration-200 flex flex-col group ";
  
  if (isActive) {
    cardClasses += "border-blue-500 ring-4 ring-blue-100 shadow-xl scale-105 cursor-default";
  } else if (isMatrixNode) {
    cardClasses += "border-purple-300 border-dashed hover:border-purple-500 hover:shadow-lg cursor-pointer";
  } else {
    cardClasses += "border-slate-200 hover:border-blue-400 hover:shadow-lg cursor-pointer";
  }

  // Elevating z-index strictly on the hovered wrapper forces it to the absolute top across all elements.
  return (
    <div className={`relative w-full ${showTooltip ? 'z-[100]' : isActive ? 'z-10' : 'z-0'}`}>
      <div className={cardClasses} onClick={!isActive ? onClick : undefined}>
        <div 
          className="absolute top-3 right-3 text-slate-400 hover:text-blue-600 z-20 cursor-help"
          onMouseEnter={handleMouseEnter}
          onMouseLeave={handleMouseLeave}
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
              {employee['Display Name']}
            </h3>
            <p className="text-xs text-slate-500 truncate" title={employee['Job Title']}>{employee['Job Title'] || 'No Title'}</p>
          </div>
        </div>
        
        <div className="text-xs text-slate-600 space-y-1 bg-slate-50 p-2 rounded-md">
          <div className="flex items-center space-x-1 truncate"><Building2 size={12} className="flex-shrink-0"/> <span className="truncate">{employee['Department (Label)'] || 'N/A'}</span></div>
          <div className="flex items-center space-x-1 truncate"><MapPin size={12} className="flex-shrink-0"/> <span className="truncate">{employee['Location Name'] || 'N/A'}</span></div>
        </div>

        {/* Bottom Row Counters: Direct, Matrix, Team */}
        <div className="mt-3 flex justify-between items-center text-[10px] font-semibold pt-2 border-t text-slate-600">
            <div 
                title={`Grades:\n${formatGrades(insights.directGrades)}`}
                onClick={(e) => { e.stopPropagation(); if(onSelectDirect) onSelectDirect(); }}
                className={`flex items-center px-1 py-0.5 rounded transition-colors cursor-pointer ${isActive && viewMode === 'direct' ? 'bg-blue-100 text-blue-800 ring-1 ring-blue-300' : 'hover:bg-blue-50 text-slate-600'}`}
            >
                <UserCircle2 size={12} className={`mr-1 ${isActive && viewMode === 'direct' ? 'text-blue-600' : 'text-blue-500'}`}/> {insights.directCount} Direct
            </div>
            
            {insights.matrixCount > 0 && (
                <div 
                    title={`Grades:\n${formatGrades(insights.matrixGrades)}`}
                    onClick={(e) => { e.stopPropagation(); if(onSelectMatrix) onSelectMatrix(); }} 
                    className={`flex items-center px-1 py-0.5 rounded transition-colors cursor-pointer ${isActive && viewMode === 'matrix' ? 'bg-purple-100 text-purple-700 ring-1 ring-purple-300' : 'hover:bg-purple-50 text-purple-600'}`}
                >
                    <span>{insights.matrixCount} Matrix</span>
                </div>
            )}

            <div title={`Grades:\n${formatGrades(insights.teamGrades)}`} className="flex items-center cursor-help px-1 py-0.5">
                <Users size={12} className="mr-1 text-slate-500"/> {insights.totalTeam} Team
            </div>
        </div>
      </div>

      {/* Advanced Insights Tooltip anchored tightly to the card using absolute positioning */}
      {showTooltip && (
        <div 
          className={`absolute w-72 bg-white rounded-xl shadow-[0_0_40px_rgba(0,0,0,0.15)] border border-slate-200 p-0 text-sm overflow-hidden animate-scale-in z-[100] ${tooltipPos.h === 'right' ? 'left-full ml-3' : 'right-full mr-3'} ${tooltipPos.v === 'top' ? 'top-0' : 'bottom-0'}`}
          onMouseEnter={() => clearTimeout(hideTimeout.current)}
          onMouseLeave={handleMouseLeave}
        >
          <div className="bg-slate-800 text-white px-4 py-3 border-b flex items-center">
            <Info size={16} className="mr-2" />
            <span className="font-semibold">Advanced Insights</span>
          </div>
          
          <div className="p-4 space-y-4">
            {/* Organizational Context */}
            {(insights.pctOfManagerTeam !== undefined || insights.peerAvgDirects !== undefined) && (
              <div className="space-y-2">
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Organizational Context</h4>
                
                {insights.pctOfManagerTeam !== undefined && (
                  <div className="bg-blue-50 px-3 py-2 rounded text-xs text-blue-800">
                    Manages <span className="font-bold">{insights.pctOfManagerTeam}%</span> of their line manager's total team size.
                  </div>
                )}
                
                {insights.peerAvgDirects !== undefined && (
                  <div className="bg-slate-50 px-3 py-2 rounded text-xs text-slate-800 border border-slate-100">
                    Peers (under same manager) average <span className="font-bold">{insights.peerAvgDirects}</span> direct reports.
                  </div>
                )}
              </div>
            )}

            {/* Team Diversity */}
            {insights.directCount > 0 && (
              <div>
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Team Diversity (Direct)</h4>
                <div className="w-full bg-slate-200 h-2 rounded-full overflow-hidden flex">
                  {malePct > 0 && <div style={{ width: `${malePct}%` }} className="bg-blue-500 h-full"></div>}
                  {femalePct > 0 && <div style={{ width: `${femalePct}%` }} className="bg-pink-500 h-full"></div>}
                </div>
                <div className="flex justify-between text-xs mt-1 text-slate-600">
                  <span>Male: {malePct}%</span>
                  <span>Female: {femalePct}%</span>
                </div>
              </div>
            )}

            {/* Management Risk */}
            {(insights.probationCount > 0 || insights.noticeCount > 0) && (
              <div>
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Management Load & Risk</h4>
                <div className="space-y-1 text-xs">
                  {insights.probationCount > 0 && (
                    <div className="flex justify-between bg-orange-50 text-orange-700 px-2 py-1 rounded">
                      <span>On Probation</span>
                      <span className="font-bold">{insights.probationCount}</span>
                    </div>
                  )}
                  {insights.noticeCount > 0 && (
                    <div className="flex justify-between bg-red-50 text-red-700 px-2 py-1 rounded">
                      <span>Flight Risk / Notice</span>
                      <span className="font-bold">{insights.noticeCount}</span>
                    </div>
                  )}
                </div>
              </div>
            )}
            
            {insights.directCount === 0 && insights.matrixCount === 0 && (
              <p className="text-xs text-slate-500 italic text-center">Individual Contributor. No direct reports.</p>
            )}
          </div>
        </div>
      )}
    </div>
  );
}