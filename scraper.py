"""
UK Truck Tyre Companies Scraper - Python Version
Scrapes Companies House API for truck tyre wholesalers and fitters
"""

import requests
import json
import csv
import time
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Configuration
API_KEY = '48d17266-ff2e-425f-9b20-7dcc9b25bb79'
BASE_URL = 'https://api.company-information.service.gov.uk'
DELAY = 0.6  # 600ms between requests

all_companies = []
seen = set()

def search_companies(query, start_index=0):
    """Search Companies House API"""
    try:
        response = requests.get(
            f'{BASE_URL}/search/companies',
            params={
                'q': query,
                'items_per_page': 100,
                'start_index': start_index
            },
            auth=(API_KEY, ''),
            headers={'Accept': 'application/json'},
            timeout=30
        )

        if response.status_code == 429:
            print('    Rate limited - waiting 60s...')
            time.sleep(60)
            return search_companies(query, start_index)

        if response.status_code == 200:
            return response.json()
        else:
            print(f'    API Error {response.status_code}')
            return {'items': [], 'total_results': 0}

    except Exception as e:
        print(f'    Error: {e}')
        return {'items': [], 'total_results': 0}

def classify_company(name):
    """Classify company type - STRICT truck tyres only"""
    name_lower = name.lower()

    # STRICT EXCLUSIONS - NOT truck tyres
    excludes = [
        'agricultural', 'tractor', 'farm', 'earthmover', 'forklift',
        'bicycle', 'motorcycle', 'motorbike', 'car tyre', 'car & van',
        'car and van', 'passenger', 'pcr', 'scooter', 'quad', 'atv',
        'golf', 'lawn', 'mower', 'garden'
    ]
    if any(ex in name_lower for ex in excludes):
        return None

    # MUST have truck/commercial/hgv/lorry related terms
    truck_terms = ['truck', 'lorry', 'hgv', 'commercial', 'fleet', 'trailer', 'artic', 'heavy goods']
    has_truck = any(t in name_lower for t in truck_terms)

    # MUST have tyre related term
    tyre_terms = ['tyre', 'tire', 'wheel']
    has_tyre = any(t in name_lower for t in tyre_terms)

    # ONLY include if it has BOTH truck AND tyre terms
    if not has_truck or not has_tyre:
        return None

    # Classify type
    if any(w in name_lower for w in ['wholesale', 'distribution', 'supply', 'distributor']):
        return 'Truck Tyre Wholesaler'
    if any(w in name_lower for w in ['retread', 'remould', 'recap']):
        return 'Truck Tyre Retreader'
    if any(w in name_lower for w in ['mobile', 'breakdown', '24 hour', 'emergency', 'roadside']):
        return 'Mobile Truck Tyre Service'
    if any(w in name_lower for w in ['fitting', 'fitter', 'service']):
        return 'Truck Tyre Fitter'

    return 'Truck Tyre Specialist'

def main():
    print('=' * 70)
    print('COMPANIES HOUSE API - UK TRUCK TYRE COMPANIES (PYTHON)')
    print('=' * 70)
    print(f'API Key: {API_KEY[:8]}...')
    print()

    # STRICT search terms - ONLY truck/commercial/HGV tyres
    search_terms = [
        'truck tyre', 'truck tyres', 'truck tyre fitting', 'truck tyre fitter',
        'truck tyre specialist', 'truck tyre wholesale', 'truck tyre service',
        'lorry tyre', 'lorry tyres', 'lorry tyre fitting',
        'hgv tyre', 'hgv tyres', 'hgv tyre fitting', 'hgv tyre fitter',
        'commercial vehicle tyre', 'commercial truck tyre', 'fleet truck tyre',
        'trailer tyre fitting', 'artic tyre', 'truck tyre mobile',
        'truck tyre breakdown', 'truck tyre 24 hour', 'truck wheel service',
        'truck tyre retread', 'commercial tyre fitting', 'commercial tyre fitter',
        'heavy goods tyre'
    ]

    for term in search_terms:
        print(f'\nSearching: "{term}"...')

        results = search_companies(term)

        if results.get('items'):
            total = results.get('total_results', 0)
            print(f'  Found {total} total results, processing {len(results["items"])}...')

            for item in results['items']:
                # Only active companies
                if item.get('company_status') != 'active':
                    continue

                # Dedupe
                company_number = item.get('company_number')
                if company_number in seen:
                    continue

                # Classify
                company_type = classify_company(item.get('title', ''))
                if not company_type:
                    continue

                seen.add(company_number)

                all_companies.append({
                    'name': item.get('title', ''),
                    'companyNumber': company_number,
                    'status': item.get('company_status', ''),
                    'type': item.get('company_type', ''),
                    'dateCreated': item.get('date_of_creation', ''),
                    'address': item.get('address_snippet', ''),
                    'businessType': company_type,
                    'sicCodes': ', '.join(item.get('sic_codes', [])),
                    'source': 'Companies House API'
                })

            # Get more pages if available
            if total > 100:
                pages_to_fetch = min(total // 100 + 1, 5)
                for page in range(1, pages_to_fetch):
                    time.sleep(DELAY)
                    more_results = search_companies(term, page * 100)

                    if more_results.get('items'):
                        for item in more_results['items']:
                            if item.get('company_status') != 'active':
                                continue

                            company_number = item.get('company_number')
                            if company_number in seen:
                                continue

                            company_type = classify_company(item.get('title', ''))
                            if not company_type:
                                continue

                            seen.add(company_number)

                            all_companies.append({
                                'name': item.get('title', ''),
                                'companyNumber': company_number,
                                'status': item.get('company_status', ''),
                                'type': item.get('company_type', ''),
                                'dateCreated': item.get('date_of_creation', ''),
                                'address': item.get('address_snippet', ''),
                                'businessType': company_type,
                                'sicCodes': ', '.join(item.get('sic_codes', [])),
                                'source': 'Companies House API'
                            })
        else:
            print('  No results or API error')

        time.sleep(DELAY)

    # Sort by business type
    type_order = {
        'Truck Tyre Wholesaler': 1,
        'Truck Tyre Retreader': 2,
        'Mobile Truck Tyre Service': 3,
        'Truck Tyre Fitter': 4,
        'Truck Tyre Specialist': 5
    }
    all_companies.sort(key=lambda x: (type_order.get(x['businessType'], 99), x['name']))

    # Write CSV
    csv_file = 'UK_TRUCK_TYRE_COMPANIES.csv'
    with open(csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=[
            'name', 'companyNumber', 'status', 'businessType', 'type',
            'dateCreated', 'address', 'sicCodes', 'source'
        ])
        writer.writeheader()
        writer.writerows(all_companies)
    print(f'\nCSV saved: {csv_file}')

    # Write JSON
    json_file = 'UK_TRUCK_TYRE_COMPANIES.json'
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(all_companies, f, indent=2)
    print(f'JSON saved: {json_file}')

    # Write Excel
    excel_file = 'UK_TRUCK_TYRE_COMPANIES.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = 'All Truck Tyre Companies'

    # Headers
    headers = ['Company Name', 'Company Number', 'Status', 'Business Type',
               'Company Type', 'Date Created', 'Address', 'SIC Codes', 'Source']
    ws.append(headers)

    # Data
    for company in all_companies:
        ws.append([
            company['name'], company['companyNumber'], company['status'],
            company['businessType'], company['type'], company['dateCreated'],
            company['address'], company['sicCodes'], company['source']
        ])

    # Set column widths
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 12
    ws.column_dimensions['G'].width = 60
    ws.column_dimensions['H'].width = 15
    ws.column_dimensions['I'].width = 20

    # Add B2B Wholesalers sheet
    ws2 = wb.create_sheet('B2B Wholesalers')
    ws2.append(headers)
    wholesalers = [c for c in all_companies if 'Wholesaler' in c['businessType']]
    for company in wholesalers:
        ws2.append([
            company['name'], company['companyNumber'], company['status'],
            company['businessType'], company['type'], company['dateCreated'],
            company['address'], company['sicCodes'], company['source']
        ])

    # Add Summary sheet
    ws3 = wb.create_sheet('Summary')
    ws3.append(['Category', 'Count'])
    ws3.append(['Total Companies', len(all_companies)])
    ws3.append(['B2B/Wholesalers', len(wholesalers)])
    ws3.append(['', ''])
    ws3.append(['--- BY BUSINESS TYPE ---', ''])

    type_counts = {}
    for c in all_companies:
        type_counts[c['businessType']] = type_counts.get(c['businessType'], 0) + 1

    for btype, count in sorted(type_counts.items(), key=lambda x: -x[1]):
        ws3.append([btype, count])

    wb.save(excel_file)
    print(f'Excel saved: {excel_file}')

    # Summary
    print('\n' + '=' * 70)
    print('SUMMARY')
    print('=' * 70)
    print(f'Total companies found: {len(all_companies)}')

    print('\nBy Business Type:')
    for btype, count in sorted(type_counts.items(), key=lambda x: -x[1]):
        print(f'  {btype}: {count}')

    print('\nSample Companies:')
    for i, c in enumerate(all_companies[:10], 1):
        print(f'{i}. {c["name"]}')
        print(f'   Company #: {c["companyNumber"]}')
        print(f'   Type: {c["businessType"]}')
        print(f'   Address: {c["address"]}')
        print()

    print('=' * 70)
    print('DONE!')
    print('=' * 70)

if __name__ == '__main__':
    main()
