/**
 * Convert CSV/JSON to Excel format
 */

import XLSX from 'xlsx';
import fs from 'fs';

// Load master data
const masterData = JSON.parse(fs.readFileSync('MASTER_UK_TRUCK_TYRE_COMPANIES.json', 'utf8'));
const companiesHouseData = JSON.parse(fs.readFileSync('companies_house_truck_tyres.json', 'utf8'));

console.log('Converting to Excel format...');

// Create workbook
const workbook = XLSX.utils.book_new();

// Sheet 1: Master List (All Truck Tyre Companies)
const masterSheet = XLSX.utils.json_to_sheet(masterData.map(c => ({
  'Company Name': c.name,
  'Website': c.website || '',
  'Phone': c.phone || '',
  'Address': c.address || '',
  'Business Type': c.businessType || '',
  'B2B/Wholesaler': c.isB2BWholesaler || '',
  'Service Points': c.servicePoints || '',
  'Region': c.region || '',
  'Companies House #': c.companyNumber || '',
  'Status': c.status || '',
  'Date Created': c.dateCreated || '',
  'Data Source': c.source || ''
})));

// Set column widths
masterSheet['!cols'] = [
  { wch: 45 }, // Company Name
  { wch: 50 }, // Website
  { wch: 18 }, // Phone
  { wch: 60 }, // Address
  { wch: 25 }, // Business Type
  { wch: 15 }, // B2B/Wholesaler
  { wch: 20 }, // Service Points
  { wch: 15 }, // Region
  { wch: 15 }, // Companies House #
  { wch: 10 }, // Status
  { wch: 12 }, // Date Created
  { wch: 20 }  // Data Source
];

XLSX.utils.book_append_sheet(workbook, masterSheet, 'All Truck Tyre Companies');

// Sheet 2: B2B Wholesalers Only
const wholesalers = masterData.filter(c => c.isB2BWholesaler === 'Yes');
const wholesalerSheet = XLSX.utils.json_to_sheet(wholesalers.map(c => ({
  'Company Name': c.name,
  'Website': c.website || '',
  'Phone': c.phone || '',
  'Address': c.address || '',
  'Business Type': c.businessType || '',
  'Service Points': c.servicePoints || '',
  'Region': c.region || '',
  'Companies House #': c.companyNumber || '',
  'Data Source': c.source || ''
})));

wholesalerSheet['!cols'] = [
  { wch: 45 }, { wch: 50 }, { wch: 18 }, { wch: 60 },
  { wch: 25 }, { wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 20 }
];

XLSX.utils.book_append_sheet(workbook, wholesalerSheet, 'B2B Wholesalers');

// Sheet 3: Companies House Verified
const chSheet = XLSX.utils.json_to_sheet(companiesHouseData.map(c => ({
  'Company Name': c.name,
  'Company Number': c.companyNumber || '',
  'Status': c.status || '',
  'Business Type': c.businessType || '',
  'Company Type': c.type || '',
  'Date Created': c.dateCreated || '',
  'Registered Address': c.address || '',
  'SIC Codes': c.sicCodes || ''
})));

chSheet['!cols'] = [
  { wch: 50 }, { wch: 15 }, { wch: 10 }, { wch: 25 },
  { wch: 15 }, { wch: 12 }, { wch: 60 }, { wch: 15 }
];

XLSX.utils.book_append_sheet(workbook, chSheet, 'Companies House Verified');

// Sheet 4: Summary
const summary = [
  { 'Category': 'Total Companies', 'Count': masterData.length },
  { 'Category': 'With Websites', 'Count': masterData.filter(c => c.website).length },
  { 'Category': 'With Addresses', 'Count': masterData.filter(c => c.address).length },
  { 'Category': 'B2B/Wholesalers', 'Count': wholesalers.length },
  { 'Category': 'Companies House Verified', 'Count': companiesHouseData.length },
  { 'Category': '', 'Count': '' },
  { 'Category': '--- BY BUSINESS TYPE ---', 'Count': '' }
];

// Add type counts
const typeCounts = {};
masterData.forEach(c => {
  typeCounts[c.businessType] = (typeCounts[c.businessType] || 0) + 1;
});
Object.entries(typeCounts).sort((a, b) => b[1] - a[1]).forEach(([type, count]) => {
  summary.push({ 'Category': type, 'Count': count });
});

const summarySheet = XLSX.utils.json_to_sheet(summary);
summarySheet['!cols'] = [{ wch: 35 }, { wch: 10 }];
XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');

// Write Excel file
XLSX.writeFile(workbook, 'UK_TRUCK_TYRE_COMPANIES.xlsx');

console.log('\n✅ Excel file created: UK_TRUCK_TYRE_COMPANIES.xlsx');
console.log('\nSheets included:');
console.log('  1. All Truck Tyre Companies (' + masterData.length + ' companies)');
console.log('  2. B2B Wholesalers (' + wholesalers.length + ' companies)');
console.log('  3. Companies House Verified (' + companiesHouseData.length + ' companies)');
console.log('  4. Summary');
