/**
 * Combine all truck tyre data into one master file
 * ONLY TRUCK TYRES - no general tyres
 */

import fs from 'fs';
import { createObjectCsvWriter } from 'csv-writer';

// Load Companies House data
let companiesHouseData = [];
try {
  companiesHouseData = JSON.parse(fs.readFileSync('companies_house_truck_tyres.json', 'utf8'));
  console.log(`Loaded ${companiesHouseData.length} from Companies House`);
} catch (e) {
  console.log('No Companies House data found');
}

// Load industry database
let industryData = [];
try {
  industryData = JSON.parse(fs.readFileSync('uk_truck_tyre_companies.json', 'utf8'));
  console.log(`Loaded ${industryData.length} from industry database`);
} catch (e) {
  console.log('No industry data found');
}

// Filter industry data to ONLY truck tyres
const truckTerms = ['truck', 'lorry', 'hgv', 'commercial', 'fleet', 'trailer', 'artic'];
const filteredIndustry = industryData.filter(company => {
  const name = company.name.toLowerCase();
  const type = (company.businessType || '').toLowerCase();

  // Must have truck-related term in name or be explicitly truck type
  const hasTruckTerm = truckTerms.some(t => name.includes(t));
  const isTruckType = type.includes('truck') || type.includes('commercial') ||
                      type.includes('hgv') || type.includes('fleet') ||
                      type.includes('manufacturer') || type.includes('wholesaler');

  return hasTruckTerm || isTruckType;
});

console.log(`Filtered industry data to ${filteredIndustry.length} truck tyre companies`);

// Combine and deduplicate
const seen = new Set();
const allCompanies = [];

// Add Companies House data first (most reliable)
for (const company of companiesHouseData) {
  const key = company.name.toLowerCase().replace(/[^a-z0-9]/g, '');
  if (!seen.has(key)) {
    seen.add(key);
    allCompanies.push({
      name: company.name,
      companyNumber: company.companyNumber || '',
      address: company.address || '',
      phone: '',
      website: '',
      businessType: company.businessType || 'Truck Tyre Specialist',
      isB2BWholesaler: company.businessType?.includes('Wholesaler') ? 'Yes' : 'No',
      servicePoints: '',
      region: 'UK',
      status: company.status || 'Active',
      dateCreated: company.dateCreated || '',
      sicCodes: company.sicCodes || '',
      source: 'Companies House API'
    });
  }
}

// Add industry data
for (const company of filteredIndustry) {
  const key = company.name.toLowerCase().replace(/[^a-z0-9]/g, '');
  if (!seen.has(key)) {
    seen.add(key);
    allCompanies.push({
      name: company.name,
      companyNumber: company.companyNumber || '',
      address: company.address || '',
      phone: company.phone || '',
      website: company.website || '',
      businessType: company.businessType || 'Truck Tyre Specialist',
      isB2BWholesaler: (company.isB2BWholesaler === 'Yes' ||
                        company.businessType?.includes('Wholesaler') ||
                        company.businessType?.includes('B2B')) ? 'Yes' : 'No',
      servicePoints: company.servicePoints || '',
      region: company.region || 'UK',
      status: company.status || 'Active',
      dateCreated: '',
      sicCodes: '',
      source: company.source || 'Industry Database'
    });
  }
}

// Sort by business type (wholesalers first)
const typeOrder = {
  'Truck Tyre Wholesaler': 1,
  'Manufacturer/Wholesaler': 2,
  'B2B Wholesaler': 3,
  'B2B Wholesaler/Retailer': 4,
  'Retreader/Wholesaler': 5,
  'Truck Tyre Retreader': 6,
  'Truck Tyre Specialist': 7,
  'Truck Tyre Fitter': 8,
  'Retail/Fitter': 9,
  'Fleet Services': 10,
  'Mobile Truck Tyre Service': 11,
  'Mobile/Emergency Services': 12,
  'Emergency Service': 13
};

allCompanies.sort((a, b) => {
  const orderA = typeOrder[a.businessType] || 99;
  const orderB = typeOrder[b.businessType] || 99;
  if (orderA !== orderB) return orderA - orderB;
  return a.name.localeCompare(b.name);
});

// Write master CSV
const csvWriter = createObjectCsvWriter({
  path: 'MASTER_UK_TRUCK_TYRE_COMPANIES.csv',
  header: [
    { id: 'name', title: 'Company Name' },
    { id: 'website', title: 'Website' },
    { id: 'phone', title: 'Phone' },
    { id: 'address', title: 'Address' },
    { id: 'businessType', title: 'Business Type' },
    { id: 'isB2BWholesaler', title: 'B2B/Wholesaler?' },
    { id: 'servicePoints', title: 'Service Points' },
    { id: 'region', title: 'Region' },
    { id: 'companyNumber', title: 'Companies House #' },
    { id: 'status', title: 'Status' },
    { id: 'dateCreated', title: 'Date Created' },
    { id: 'source', title: 'Data Source' }
  ]
});

await csvWriter.writeRecords(allCompanies);

// Write master JSON
fs.writeFileSync('MASTER_UK_TRUCK_TYRE_COMPANIES.json', JSON.stringify(allCompanies, null, 2));

// Summary
console.log('\n' + '='.repeat(70));
console.log('MASTER FILE CREATED - TRUCK TYRES ONLY');
console.log('='.repeat(70));
console.log(`Total companies: ${allCompanies.length}`);
console.log(`With websites: ${allCompanies.filter(c => c.website).length}`);
console.log(`With addresses: ${allCompanies.filter(c => c.address).length}`);
console.log(`With Companies House #: ${allCompanies.filter(c => c.companyNumber).length}`);

// By type
const typeCounts = {};
allCompanies.forEach(c => {
  typeCounts[c.businessType] = (typeCounts[c.businessType] || 0) + 1;
});

console.log('\nBy Business Type:');
Object.entries(typeCounts).sort((a, b) => (typeOrder[a[0]] || 99) - (typeOrder[b[0]] || 99)).forEach(([type, count]) => {
  console.log(`  ${type}: ${count}`);
});

// B2B/Wholesalers
const wholesalers = allCompanies.filter(c => c.isB2BWholesaler === 'Yes');
console.log(`\nB2B/Wholesalers: ${wholesalers.length}`);

// By source
const sourceCounts = {};
allCompanies.forEach(c => {
  sourceCounts[c.source] = (sourceCounts[c.source] || 0) + 1;
});

console.log('\nBy Data Source:');
Object.entries(sourceCounts).forEach(([source, count]) => {
  console.log(`  ${source}: ${count}`);
});

console.log('\n' + '='.repeat(70));
console.log('FILES CREATED:');
console.log('  MASTER_UK_TRUCK_TYRE_COMPANIES.csv');
console.log('  MASTER_UK_TRUCK_TYRE_COMPANIES.json');
console.log('='.repeat(70));
