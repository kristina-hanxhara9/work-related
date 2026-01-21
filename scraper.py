"""
UK Truck Tyre Companies Scraper - Python Version
Loads existing data + scrapes Companies House API
Full 800+ companies
"""

import requests
import json
import csv
import time
import os
from openpyxl import Workbook

# Configuration
API_KEY = '48d17266-ff2e-425f-9b20-7dcc9b25bb79'
BASE_URL = 'https://api.company-information.service.gov.uk'
DELAY = 0.6

all_companies = []
seen = set()


def search_companies(query, start_index=0):
    """Search Companies House API"""
    try:
        response = requests.get(
            f'{BASE_URL}/search/companies',
            params={'q': query, 'items_per_page': 100, 'start_index': start_index},
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

    excludes = ['agricultural', 'tractor', 'farm', 'earthmover', 'forklift',
                'bicycle', 'motorcycle', 'motorbike', 'car tyre', 'car & van',
                'car and van', 'passenger', 'pcr', 'scooter', 'quad', 'atv',
                'golf', 'lawn', 'mower', 'garden']
    if any(ex in name_lower for ex in excludes):
        return None

    truck_terms = ['truck', 'lorry', 'hgv', 'commercial', 'fleet', 'trailer', 'artic', 'heavy goods']
    has_truck = any(t in name_lower for t in truck_terms)

    tyre_terms = ['tyre', 'tire', 'wheel']
    has_tyre = any(t in name_lower for t in tyre_terms)

    if not has_truck or not has_tyre:
        return None

    if any(w in name_lower for w in ['wholesale', 'distribution', 'supply', 'distributor']):
        return 'Truck Tyre Wholesaler'
    if any(w in name_lower for w in ['retread', 'remould', 'recap']):
        return 'Truck Tyre Retreader'
    if any(w in name_lower for w in ['mobile', 'breakdown', '24 hour', 'emergency', 'roadside']):
        return 'Mobile Truck Tyre Service'
    if any(w in name_lower for w in ['fitting', 'fitter', 'service']):
        return 'Truck Tyre Fitter'

    return 'Truck Tyre Specialist'


def load_existing_data():
    """Load existing master data if available"""
    files_to_try = ['MASTER_UK_TRUCK_TYRE_COMPANIES.json', 'uk_truck_tyre_companies.json']

    for filename in files_to_try:
        if os.path.exists(filename):
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                print(f'Loaded {len(data)} companies from {filename}')
                return data
            except:
                pass
    return []


def main():
    print('=' * 70)
    print('UK TRUCK TYRE COMPANIES SCRAPER')
    print('Companies House API + Industry Database')
    print('=' * 70)
    print()

    # Load existing data
    existing_data = load_existing_data()

    if existing_data:
        print(f'Loading {len(existing_data)} existing companies...')
        for company in existing_data:
            key = company.get('name', '').lower().replace(' ', '').replace('-', '')
            if key and key not in seen:
                seen.add(key)
                all_companies.append({
                    'name': company.get('name', ''),
                    'website': company.get('website', ''),
                    'phone': company.get('phone', ''),
                    'address': company.get('address', ''),
                    'businessType': company.get('businessType', 'Truck Tyre Specialist'),
                    'isB2BWholesaler': company.get('isB2BWholesaler', 'No'),
                    'servicePoints': company.get('servicePoints', ''),
                    'region': company.get('region', 'UK'),
                    'companyNumber': company.get('companyNumber', ''),
                    'status': company.get('status', 'Active'),
                    'dateCreated': company.get('dateCreated', ''),
                    'source': company.get('source', 'Industry Database')
                })
        print(f'  Added {len(all_companies)} companies from existing data')

    # Search Companies House API for new companies
    print()
    print('Searching Companies House API for additional companies...')
    print(f'API Key: {API_KEY[:8]}...')

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

    api_count = 0
    for term in search_terms:
        print(f'\n  Searching: "{term}"...')
        results = search_companies(term)

        if results.get('items'):
            for item in results['items']:
                if item.get('company_status') != 'active':
                    continue

                name = item.get('title', '')
                key = name.lower().replace(' ', '').replace('-', '')
                if key in seen:
                    continue

                company_type = classify_company(name)
                if not company_type:
                    continue

                seen.add(key)
                api_count += 1

                all_companies.append({
                    'name': name,
                    'website': '',
                    'phone': '',
                    'address': item.get('address_snippet', ''),
                    'businessType': company_type,
                    'isB2BWholesaler': 'Yes' if 'Wholesaler' in company_type else 'No',
                    'servicePoints': '',
                    'region': 'UK',
                    'companyNumber': item.get('company_number', ''),
                    'status': item.get('company_status', ''),
                    'dateCreated': item.get('date_of_creation', ''),
                    'source': 'Companies House API'
                })

        time.sleep(DELAY)

    print(f'\n  Added {api_count} NEW companies from Companies House API')

    # Sort
    type_order = {
        'Manufacturer/Wholesaler': 1, 'B2B Wholesaler': 2, 'B2B Wholesaler/Retailer': 3,
        'Retreader/Wholesaler': 4, 'Truck Tyre Wholesaler': 5, 'Truck Tyre Retreader': 6,
        'Truck Tyre Specialist': 7, 'Truck Tyre Fitter': 8, 'Mobile Truck Tyre Service': 9,
        'Mobile/Emergency Services': 10
    }
    all_companies.sort(key=lambda x: (type_order.get(x['businessType'], 99), x['name']))

    # Write CSV
    csv_file = 'UK_TRUCK_TYRE_COMPANIES.csv'
    with open(csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=[
            'name', 'website', 'phone', 'address', 'businessType',
            'isB2BWholesaler', 'servicePoints', 'region',
            'companyNumber', 'status', 'dateCreated', 'source'
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

    # Sheet 1: All Companies
    ws = wb.active
    ws.title = 'All Truck Tyre Companies'
    headers = ['Company Name', 'Website', 'Phone', 'Address', 'Business Type',
               'B2B/Wholesaler', 'Service Points', 'Region',
               'Companies House #', 'Status', 'Date Created', 'Source']
    ws.append(headers)
    for c in all_companies:
        ws.append([c['name'], c['website'], c['phone'], c['address'],
                   c['businessType'], c['isB2BWholesaler'], c['servicePoints'],
                   c['region'], c['companyNumber'], c['status'],
                   c['dateCreated'], c['source']])

    ws.column_dimensions['A'].width = 45
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 60
    ws.column_dimensions['E'].width = 25
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 20
    ws.column_dimensions['H'].width = 15
    ws.column_dimensions['I'].width = 15
    ws.column_dimensions['J'].width = 10
    ws.column_dimensions['K'].width = 12
    ws.column_dimensions['L'].width = 20

    # Sheet 2: B2B Wholesalers
    ws2 = wb.create_sheet('B2B Wholesalers')
    ws2.append(headers)
    wholesalers = [c for c in all_companies if c['isB2BWholesaler'] == 'Yes']
    for c in wholesalers:
        ws2.append([c['name'], c['website'], c['phone'], c['address'],
                    c['businessType'], c['isB2BWholesaler'], c['servicePoints'],
                    c['region'], c['companyNumber'], c['status'],
                    c['dateCreated'], c['source']])

    # Sheet 3: Companies House Verified
    ws3 = wb.create_sheet('Companies House Verified')
    ws3.append(headers)
    ch_companies = [c for c in all_companies if c['companyNumber']]
    for c in ch_companies:
        ws3.append([c['name'], c['website'], c['phone'], c['address'],
                    c['businessType'], c['isB2BWholesaler'], c['servicePoints'],
                    c['region'], c['companyNumber'], c['status'],
                    c['dateCreated'], c['source']])

    # Sheet 4: Summary
    ws4 = wb.create_sheet('Summary')
    ws4.append(['Category', 'Count'])
    ws4.append(['Total Companies', len(all_companies)])
    ws4.append(['With Websites', len([c for c in all_companies if c['website']])])
    ws4.append(['With Addresses', len([c for c in all_companies if c['address']])])
    ws4.append(['B2B/Wholesalers', len(wholesalers)])
    ws4.append(['Companies House Verified', len(ch_companies)])
    ws4.append(['', ''])
    ws4.append(['--- BY BUSINESS TYPE ---', ''])

    type_counts = {}
    for c in all_companies:
        type_counts[c['businessType']] = type_counts.get(c['businessType'], 0) + 1
    for btype, count in sorted(type_counts.items(), key=lambda x: type_order.get(x[0], 99)):
        ws4.append([btype, count])

    ws4.column_dimensions['A'].width = 35
    ws4.column_dimensions['B'].width = 10

    wb.save(excel_file)
    print(f'Excel saved: {excel_file}')

    # Summary
    print('\n' + '=' * 70)
    print('SUMMARY')
    print('=' * 70)
    print(f'Total companies: {len(all_companies)}')
    print(f'With websites: {len([c for c in all_companies if c["website"]])}')
    print(f'With addresses: {len([c for c in all_companies if c["address"]])}')
    print(f'B2B/Wholesalers: {len(wholesalers)}')
    print(f'Companies House verified: {len(ch_companies)}')

    print('\nBy Business Type:')
    for btype, count in sorted(type_counts.items(), key=lambda x: type_order.get(x[0], 99)):
        print(f'  {btype}: {count}')

    print('\n' + '=' * 70)
    print('FILES CREATED:')
    print(f'  {csv_file}')
    print(f'  {json_file}')
    print(f'  {excel_file}')
    print('=' * 70)


if __name__ == '__main__':
    main()
