// Regenerates orglens_sample_template.xlsx at the repo root.
// Single source of truth for the spreadsheet schema the app expects.
// Run with: npm run build:template
//
// Row 1 = column headers
// Row 2 = tier labels (Required / Recommended / Optional) — handleFileUpload
//         filters this row out at upload time
// Rows 3+ = sample data demonstrating each feature

import * as XLSX from 'xlsx';
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const outPath = resolve(here, '..', 'orglens_sample_template.xlsx');

const REQUIRED = ["Employee's Position Code", "Employee name", "Line Manager's Position Code"];
const RECOMMENDED = [
    'Line Manager Name', 'Job Title', 'Level', 'Employee Class',
    'Function 1', 'Function/Plant', 'Location Name', 'Asset', 'Cluster',
    'Gender', 'Date of Birth', 'HR Manager Name', 'HR Manager EID', 'Management Board EID',
];
const OPTIONAL = [
    'Date of Joining', 'Date in Role', 'Date Promoted', 'Manager Since',
    'Email', 'Photo URL', 'Matrix Manager EID(s)', 'Cohort Tags', 'Current Status',
];
const HEADERS = [...REQUIRED, ...RECOMMENDED, ...OPTIONAL];

const tierRow = HEADERS.map(h =>
    REQUIRED.includes(h) ? 'Required' : RECOMMENDED.includes(h) ? 'Recommended' : 'Optional'
);

// Sample rows — illustrate filled positions, MC membership, matrix, cohorts,
// vacancy ("Active"/"WIP"/"Offered"/"Vacant") and the Approved/Unapproved
// placeholder name convention.
const rows = [
    // CEO — Active, MC member
    {
        eid: 'P0001', name: 'Asha Menon', mgrEid: '',
        mgrName: '', title: 'Managing Director & CEO', level: 'MD - Managing Director',
        empClass: 'Management', function1: 'Corporate', plant: "MDs Office", location: 'Head Office',
        asset: 'Others', cluster: 'Others', gender: 'F', dob: '1965-04-12',
        hrName: 'Priya Iyer', hrEid: 'P0003', boardEid: 'P0001',
        doj: '1998-06-01', dir: '2020-01-01', prom: '2020-01-01', mgrSince: '',
        email: 'asha.menon@example.com', photo: '', matrix: '', cohorts: '',
        status: 'Active',
    },
    // CFO — Active, MC member
    {
        eid: 'P0002', name: 'Rohit Sharma', mgrEid: 'P0001',
        mgrName: 'Asha Menon', title: 'Chief Financial Officer', level: 'SVP - Sr. Vice President',
        empClass: 'Management', function1: 'Finance', plant: 'Corporate', location: 'Head Office',
        asset: 'Others', cluster: 'Others', gender: 'M', dob: '1972-08-03',
        hrName: 'Priya Iyer', hrEid: 'P0003', boardEid: 'P0002',
        doj: '2005-04-15', dir: '2022-06-01', prom: '2022-06-01', mgrSince: '2020-01-01',
        email: 'rohit.sharma@example.com', photo: '', matrix: '', cohorts: '',
        status: 'Active',
    },
    // CHRO — Active, MC member
    {
        eid: 'P0003', name: 'Priya Iyer', mgrEid: 'P0001',
        mgrName: 'Asha Menon', title: 'Chief Human Resources Officer', level: 'SVP - Sr. Vice President',
        empClass: 'Management', function1: 'HR', plant: 'Corporate HR', location: 'Head Office',
        asset: 'Others', cluster: 'Others', gender: 'F', dob: '1978-01-22',
        hrName: 'Priya Iyer', hrEid: 'P0003', boardEid: 'P0003',
        doj: '2010-09-01', dir: '2021-04-01', prom: '2021-04-01', mgrSince: '2020-01-01',
        email: 'priya.iyer@example.com', photo: '', matrix: 'P0002', cohorts: '',
        status: 'Active',
    },
    // VP Finance — currently filled, marked as a HiPo
    {
        eid: 'P0004', name: 'Karan Mehta', mgrEid: 'P0002',
        mgrName: 'Rohit Sharma', title: 'VP - Finance', level: 'VP - Vice President',
        empClass: 'Management', function1: 'Finance', plant: 'Corporate', location: 'Head Office',
        asset: 'Others', cluster: 'Others', gender: 'M', dob: '1981-11-09',
        hrName: 'Sneha Pillai', hrEid: 'P0005', boardEid: '',
        doj: '2015-03-20', dir: '2023-01-01', prom: '2023-01-01', mgrSince: '2022-06-01',
        email: 'karan.mehta@example.com', photo: '', matrix: '', cohorts: 'High Potential',
        status: 'Active',
    },
    // GM Talent Management — Active
    {
        eid: 'P0005', name: 'Sneha Pillai', mgrEid: 'P0002',
        mgrName: 'Rohit Sharma', title: 'GM - Talent Management', level: 'GM - General Manager',
        empClass: 'Management', function1: 'HR', plant: 'Talent Management', location: 'Head Office',
        asset: 'Others', cluster: 'Others', gender: 'F', dob: '1985-06-30',
        hrName: 'Priya Iyer', hrEid: 'P0003', boardEid: '',
        doj: '2018-07-10', dir: '2024-04-01', prom: '2024-04-01', mgrSince: '2022-06-01',
        email: 'sneha.pillai@example.com', photo: '', matrix: 'P0003', cohorts: 'High Potential',
        status: 'Active',
    },
    // VP Engineering — currently in WIP (interviews underway)
    {
        eid: 'P0006', name: 'Vikram Singh', mgrEid: 'P0001',
        mgrName: 'Asha Menon', title: 'VP - Engineering', level: 'VP - Vice President',
        empClass: 'Management', function1: 'Engineering', plant: 'Corporate', location: 'Hazira Plant',
        asset: 'Hot Rolled', cluster: 'West', gender: 'M', dob: '1979-03-14',
        hrName: 'Sneha Pillai', hrEid: 'P0005', boardEid: '',
        doj: '2016-08-01', dir: '2024-01-01', prom: '2024-01-01', mgrSince: '2024-01-01',
        email: 'vikram.singh@example.com', photo: '', matrix: '', cohorts: '',
        status: 'WIP',
    },
    // Sr Engineer — Offered, candidate accepted not yet onboard
    {
        eid: 'P0007', name: 'Maya Krishnan', mgrEid: 'P0006',
        mgrName: 'Vikram Singh', title: 'Senior Manager - Plant Operations', level: 'M3',
        empClass: 'Officer', function1: 'Engineering', plant: 'Hot Strip Mill', location: 'Hazira Plant',
        asset: 'Hot Rolled', cluster: 'West', gender: 'F', dob: '1988-12-05',
        hrName: 'Sneha Pillai', hrEid: 'P0005', boardEid: '',
        doj: '', dir: '', prom: '', mgrSince: '',
        email: 'maya.krishnan@example.com', photo: '', matrix: '', cohorts: '',
        status: 'Offered',
    },
    // APPROVED placeholder seat (HR has sanctioned hire, role still open)
    {
        eid: 'P0008', name: 'Approved', mgrEid: 'P0006',
        mgrName: 'Vikram Singh', title: 'Manager - Quality Assurance', level: 'M2',
        empClass: 'Officer', function1: 'Quality', plant: 'Cold Rolling Mill', location: 'Hazira Plant',
        asset: 'Cold Rolled', cluster: 'West', gender: '', dob: '',
        hrName: 'Sneha Pillai', hrEid: 'P0005', boardEid: '',
        doj: '', dir: '', prom: '', mgrSince: '',
        email: '', photo: '', matrix: '', cohorts: '',
        status: 'Vacant',
    },
    // UNAPPROVED proposed seat (not yet sanctioned)
    {
        eid: 'P0009', name: 'Unapproved', mgrEid: 'P0006',
        mgrName: 'Vikram Singh', title: 'Lead Engineer - Automation', level: 'M2',
        empClass: 'Officer', function1: 'Engineering', plant: 'Automation Lab', location: 'Hazira Plant',
        asset: 'Hot Rolled', cluster: 'West', gender: '', dob: '',
        hrName: 'Sneha Pillai', hrEid: 'P0005', boardEid: '',
        doj: '', dir: '', prom: '', mgrSince: '',
        email: '', photo: '', matrix: '', cohorts: '',
        status: 'Vacant',
    },
];

const sample = rows.map(r => ({
    "Employee's Position Code": r.eid,
    "Employee name": r.name,
    "Line Manager's Position Code": r.mgrEid,
    'Line Manager Name': r.mgrName,
    'Job Title': r.title,
    'Level': r.level,
    'Employee Class': r.empClass,
    'Function 1': r.function1,
    'Function/Plant': r.plant,
    'Location Name': r.location,
    'Asset': r.asset,
    'Cluster': r.cluster,
    'Gender': r.gender,
    'Date of Birth': r.dob,
    'HR Manager Name': r.hrName,
    'HR Manager EID': r.hrEid,
    'Management Board EID': r.boardEid,
    'Date of Joining': r.doj,
    'Date in Role': r.dir,
    'Date Promoted': r.prom,
    'Manager Since': r.mgrSince,
    'Email': r.email,
    'Photo URL': r.photo,
    'Matrix Manager EID(s)': r.matrix,
    'Cohort Tags': r.cohorts,
    'Current Status': r.status,
}));

const aoa = [HEADERS, tierRow, ...sample.map(row => HEADERS.map(h => row[h]))];

const ws = XLSX.utils.aoa_to_sheet(aoa);
ws['!cols'] = HEADERS.map(h => ({ wch: Math.max(14, h.length + 2) }));
ws['!freeze'] = { xSplit: 0, ySplit: 2 };

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Employees');

const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
writeFileSync(outPath, buf);
console.log(`wrote ${outPath} (${aoa.length} rows, ${HEADERS.length} cols)`);
