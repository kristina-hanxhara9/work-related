/**
 * Companies House API Scraper for UK Truck Tyre Companies
 * Uses the official Companies House REST API
 */

import axios from 'axios';
import { createObjectCsvWriter } from 'csv-writer';
import fs from 'fs';

const CONFIG = {
  API_KEY: '48d17266-ff2e-425f-9b20-7dcc9b25bb79',
  BASE_URL: 'https://api.company-information.service.gov.uk',
  OUTPUT_FILE: 'companies_house_truck_tyres.csv',
  JSON_OUTPUT_FILE: 'companies_house_truck_tyres.json',
  DELAY: 600 // Companies House allows 600 requests per 5 mins
};

let allCompanies = [];
const seen = new Set();

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function searchCompanies(query, startIndex = 0) {
  try {
    const response = await axios.get(`${CONFIG.BASE_URL}/search/companies`, {
      params: {
        q: query,
        items_per_page: 100,
        start_index: startIndex
      },
      auth: {
        username: CONFIG.API_KEY,
        password: ''
      },
      headers: {
        'Accept': 'application/json'
      },
      timeout: 30000
    });

    return response.data;
  } catch (error) {
    if (error.response) {
      console.log(`    API Error ${error.response.status}: ${error.response.statusText}`);
      if (error.response.status === 429) {
        console.log('    Rate limited - waiting 60s...');
        await delay(60000);
        return searchCompanies(query, startIndex);
      }
    } else {
      console.log(`    Error: ${error.message}`);
    }
    return { items: [], total_results: 0 };
  }
}

async function getCompanyDetails(companyNumber) {
  try {
    const response = await axios.get(`${CONFIG.BASE_URL}/company/${companyNumber}`, {
      auth: {
        username: CONFIG.API_KEY,
        password: ''
      },
      headers: {
        'Accept': 'application/json'
      },
      timeout: 15000
    });

    return response.data;
  } catch (error) {
    return null;
  }
}

function formatAddress(addr) {
  if (!addr) return '';
  const parts = [
    addr.premises,
    addr.address_line_1,
    addr.address_line_2,
    addr.locality,
    addr.region,
    addr.postal_code,
    addr.country
  ].filter(Boolean);
  return parts.join(', ');
}

function classifyCompany(name, sicCodes = []) {
  const nameLower = name.toLowerCase();

  // STRICT EXCLUSIONS - NOT truck tyres
  const excludes = [
    'agricultural', 'tractor', 'farm', 'earthmover', 'forklift',
    'bicycle', 'motorcycle', 'motorbike', 'car tyre', 'car & van',
    'car and van', 'passenger', 'pcr', 'scooter', 'quad', 'atv',
    'golf', 'lawn', 'mower', 'garden'
  ];
  if (excludes.some(ex => nameLower.includes(ex))) {
    return null;
  }

  // MUST have truck/commercial/hgv/lorry related terms
  const truckTerms = ['truck', 'lorry', 'hgv', 'commercial', 'fleet', 'trailer', 'artic', 'heavy goods'];
  const hasTruck = truckTerms.some(t => nameLower.includes(t));

  // MUST have tyre related term
  const tyreTerms = ['tyre', 'tire', 'wheel'];
  const hasTyre = tyreTerms.some(t => nameLower.includes(t));

  // ONLY include if it has BOTH truck AND tyre terms
  if (!hasTruck || !hasTyre) {
    return null;
  }

  // Classify type
  if (nameLower.includes('wholesale') || nameLower.includes('distribution') || nameLower.includes('supply') || nameLower.includes('distributor')) {
    return 'Truck Tyre Wholesaler';
  }
  if (nameLower.includes('retread') || nameLower.includes('remould') || nameLower.includes('recap')) {
    return 'Truck Tyre Retreader';
  }
  if (nameLower.includes('mobile') || nameLower.includes('breakdown') || nameLower.includes('24 hour') || nameLower.includes('emergency') || nameLower.includes('roadside')) {
    return 'Mobile Truck Tyre Service';
  }
  if (nameLower.includes('fitting') || nameLower.includes('fitter') || nameLower.includes('service')) {
    return 'Truck Tyre Fitter';
  }

  return 'Truck Tyre Specialist';
}

async function main() {
  console.log('='.repeat(70));
  console.log('COMPANIES HOUSE API - UK TRUCK TYRE COMPANIES');
  console.log('='.repeat(70));
  console.log(`API Key: ${CONFIG.API_KEY.substring(0, 8)}...`);
  console.log('');

  // STRICT search terms - ONLY truck/commercial/HGV tyres
  const searchTerms = [
    'truck tyre',
    'truck tyres',
    'truck tyre fitting',
    'truck tyre fitter',
    'truck tyre specialist',
    'truck tyre wholesale',
    'truck tyre service',
    'lorry tyre',
    'lorry tyres',
    'lorry tyre fitting',
    'hgv tyre',
    'hgv tyres',
    'hgv tyre fitting',
    'hgv tyre fitter',
    'commercial vehicle tyre',
    'commercial truck tyre',
    'fleet truck tyre',
    'trailer tyre fitting',
    'artic tyre',
    'truck tyre mobile',
    'truck tyre breakdown',
    'truck tyre 24 hour',
    'truck wheel service',
    'truck tyre retread',
    'commercial tyre fitting',
    'commercial tyre fitter',
    'heavy goods tyre'
  ];

  for (const term of searchTerms) {
    console.log(`\nSearching: "${term}"...`);

    const results = await searchCompanies(term);

    if (results.items && results.items.length > 0) {
      console.log(`  Found ${results.total_results} total results, processing ${results.items.length}...`);

      for (const item of results.items) {
        // Only active companies
        if (item.company_status !== 'active') continue;

        // Dedupe
        if (seen.has(item.company_number)) continue;

        // Classify
        const type = classifyCompany(item.title, item.sic_codes || []);
        if (!type) continue;

        seen.add(item.company_number);

        allCompanies.push({
          companyNumber: item.company_number,
          name: item.title,
          status: item.company_status,
          type: item.company_type,
          dateCreated: item.date_of_creation,
          address: item.address_snippet || '',
          businessType: type,
          sicCodes: (item.sic_codes || []).join(', '),
          source: 'Companies House API'
        });
      }

      // Get more pages if available
      if (results.total_results > 100) {
        const pagesToFetch = Math.min(Math.ceil(results.total_results / 100), 5);
        for (let page = 1; page < pagesToFetch; page++) {
          await delay(CONFIG.DELAY);
          const moreResults = await searchCompanies(term, page * 100);

          if (moreResults.items) {
            for (const item of moreResults.items) {
              if (item.company_status !== 'active') continue;
              if (seen.has(item.company_number)) continue;

              const type = classifyCompany(item.title, item.sic_codes || []);
              if (!type) continue;

              seen.add(item.company_number);

              allCompanies.push({
                companyNumber: item.company_number,
                name: item.title,
                status: item.company_status,
                type: item.company_type,
                dateCreated: item.date_of_creation,
                address: item.address_snippet || '',
                businessType: type,
                sicCodes: (item.sic_codes || []).join(', '),
                source: 'Companies House API'
              });
            }
          }
        }
      }
    } else {
      console.log(`  No results or API error`);
    }

    await delay(CONFIG.DELAY);
  }

  // Get additional details for some companies
  console.log(`\n\nEnriching top companies with full details...`);

  const toEnrich = allCompanies.slice(0, 50);
  let enriched = 0;

  for (const company of toEnrich) {
    const details = await getCompanyDetails(company.companyNumber);
    if (details) {
      company.address = formatAddress(details.registered_office_address);
      company.sicCodes = (details.sic_codes || []).join(', ');
      if (details.accounts && details.accounts.last_accounts) {
        company.lastAccounts = details.accounts.last_accounts.made_up_to;
      }
    }
    enriched++;
    if (enriched % 10 === 0) {
      console.log(`  Enriched ${enriched}/50...`);
    }
    await delay(CONFIG.DELAY);
  }

  // Sort by business type
  allCompanies.sort((a, b) => {
    const order = {
      'Wholesaler/Distributor': 1,
      'Retreader': 2,
      'Truck Tyre Specialist': 3,
      'Mobile/Emergency Service': 4,
      'Tyre Services': 5
    };
    return (order[a.businessType] || 99) - (order[b.businessType] || 99);
  });

  // Write CSV
  const csvWriter = createObjectCsvWriter({
    path: CONFIG.OUTPUT_FILE,
    header: [
      { id: 'name', title: 'Company Name' },
      { id: 'companyNumber', title: 'Company Number' },
      { id: 'status', title: 'Status' },
      { id: 'businessType', title: 'Business Type' },
      { id: 'type', title: 'Company Type' },
      { id: 'dateCreated', title: 'Date Created' },
      { id: 'address', title: 'Registered Address' },
      { id: 'sicCodes', title: 'SIC Codes' },
      { id: 'lastAccounts', title: 'Last Accounts' },
      { id: 'source', title: 'Source' }
    ]
  });

  await csvWriter.writeRecords(allCompanies);
  console.log(`\nCSV saved: ${CONFIG.OUTPUT_FILE}`);

  // Write JSON
  fs.writeFileSync(CONFIG.JSON_OUTPUT_FILE, JSON.stringify(allCompanies, null, 2));
  console.log(`JSON saved: ${CONFIG.JSON_OUTPUT_FILE}`);

  // Summary
  console.log('\n' + '='.repeat(70));
  console.log('SUMMARY');
  console.log('='.repeat(70));
  console.log(`Total companies found: ${allCompanies.length}`);

  const typeCounts = {};
  allCompanies.forEach(c => {
    typeCounts[c.businessType] = (typeCounts[c.businessType] || 0) + 1;
  });

  console.log('\nBy Business Type:');
  Object.entries(typeCounts).sort((a, b) => b[1] - a[1]).forEach(([type, count]) => {
    console.log(`  ${type}: ${count}`);
  });

  console.log('\nSample Companies:');
  allCompanies.slice(0, 15).forEach((c, i) => {
    console.log(`${i+1}. ${c.name}`);
    console.log(`   Company #: ${c.companyNumber}`);
    console.log(`   Type: ${c.businessType}`);
    console.log(`   Address: ${c.address}`);
    console.log('');
  });

  console.log('='.repeat(70));
  console.log('DONE!');
  console.log('='.repeat(70));
}

main().catch(console.error);
