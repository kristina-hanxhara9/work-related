#!/usr/bin/env python3
"""
UK Retailers Scraper - SIC Code Based
======================================
- Multiple SIC codes for motor vehicle/tyre businesses
- Excludes Northern Ireland
- Active companies only
- Separate filters for: truck, tyre, tire
"""

import requests
import json
import time
from datetime import datetime
import csv

# Companies House API credentials
API_KEY = "48d17266-ff2e-425f-9b20-7dcc9b25bb79"
BASE_URL = "https://api.company-information.service.gov.uk"

# Target SIC codes
TARGET_SIC_CODES = {
    "45200": "Maintenance and repair of motor vehicles",
    "45320": "Retail trade of motor vehicle parts and accessories",
    "46900": "Non-specialised wholesale trade",
    "22110": "Manufacture of rubber tyres and tubes; retreading and rebuilding of rubber tyres",
    "45310": "Wholesale trade of motor vehicle parts and accessories",
    "82990": "Other business support service activities n.e.c.",
    "29320": "Manufacture of other parts and accessories for motor vehicles",
    "30170": "Repair and maintenance of other transport equipment",
    "33190": "Repair of other equipment",
}

# Description filters (will search separately)
DESCRIPTION_FILTERS = ["truck", "tyre", "tire"]


def search_by_sic_code(sic_code):
    """
    Search Companies House for ALL active companies with a specific SIC code.
    Excludes Northern Ireland.
    """
    all_companies = []
    start_index = 0
    items_per_page = 500
    total_available = None

    while True:
        url = f"{BASE_URL}/advanced-search/companies"
        params = {
            "sic_codes": sic_code,
            "size": items_per_page,
            "start_index": start_index,
            "company_status": "active"
        }

        try:
            response = requests.get(url, params=params, auth=(API_KEY, ''))

            if response.status_code == 200:
                data = response.json()
                items = data.get('items', [])
                total_results = data.get('hits', 0)

                if total_available is None:
                    total_available = total_results
                    print(f"    Total available: {total_available:,} active companies")

                if not items:
                    break

                # Filter out Northern Ireland companies
                for company in items:
                    company_number = company.get('company_number', '')
                    address = company.get('registered_office_address', {})

                    # Skip Northern Ireland companies
                    is_ni = (
                        company_number.startswith('NI') or
                        company_number.startswith('R0') or
                        'northern ireland' in str(address).lower() or
                        address.get('country', '').lower() == 'northern ireland'
                    )

                    if not is_ni:
                        all_companies.append(company)

                print(f"    Fetched: {len(all_companies):,} (excl. NI) / {total_available:,} total")

                if start_index + items_per_page >= total_available:
                    break

                if len(items) < items_per_page:
                    break

                start_index += items_per_page
                time.sleep(0.2)

            elif response.status_code == 416:
                break
            elif response.status_code == 429:
                print("    Rate limited. Waiting 60 seconds...")
                time.sleep(60)
                continue
            else:
                print(f"    API error {response.status_code}")
                break

        except Exception as e:
            print(f"    Error: {e}")
            time.sleep(5)
            continue

    return all_companies


def main():
    print("=" * 70)
    print("UK RETAILERS SCRAPER - SIC CODE BASED")
    print("=" * 70)
    print("Source: Companies House API ONLY")
    print("Filter: ACTIVE companies only")
    print("Excludes: Northern Ireland")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)

    print("\nTarget SIC codes:")
    for code, desc in TARGET_SIC_CODES.items():
        print(f"  {code}: {desc}")

    print(f"\nDescription filters: {', '.join(DESCRIPTION_FILTERS)}")

    # Collect all companies by SIC code
    all_companies = {}
    sic_code_counts = {}

    print(f"\nSearching Companies House API...\n")

    for sic_code, description in TARGET_SIC_CODES.items():
        print(f"[SIC {sic_code}] {description}")

        companies = search_by_sic_code(sic_code)
        sic_code_counts[sic_code] = len(companies)

        print(f"    TOTAL Retrieved: {len(companies):,} active companies (excl. NI)\n")

        for company in companies:
            company_number = company.get('company_number')

            if company_number not in all_companies:
                all_companies[company_number] = {
                    'company_number': company_number,
                    'company_name': company.get('company_name', ''),
                    'status': company.get('company_status', 'active'),
                    'company_type': company.get('company_type', ''),
                    'date_of_creation': company.get('date_of_creation', ''),
                    'address': company.get('registered_office_address', {}),
                    'sic_codes': company.get('sic_codes', []),
                    'matched_sic_codes': [sic_code]
                }
            else:
                if sic_code not in all_companies[company_number]['matched_sic_codes']:
                    all_companies[company_number]['matched_sic_codes'].append(sic_code)

        time.sleep(1)

    print(f"{'=' * 70}")
    print(f"Search complete. Found {len(all_companies):,} unique ACTIVE companies (excl. NI)")
    print(f"{'=' * 70}")

    # Process and format data
    print("\nProcessing company data...")

    final_companies = []

    for company_number, company in all_companies.items():
        # Format address
        addr = company.get('address', {})
        if isinstance(addr, dict):
            address_parts = []
            for field in ['address_line_1', 'address_line_2', 'locality', 'region', 'postal_code', 'country']:
                if addr.get(field):
                    address_parts.append(addr[field])
            full_address = ', '.join(address_parts)
            postcode = addr.get('postal_code', '')
            locality = addr.get('locality', '')
            region = addr.get('region', '')
        else:
            full_address = str(addr) if addr else ''
            postcode = ''
            locality = ''
            region = ''

        # Get SIC descriptions
        sic_codes = company.get('sic_codes', [])
        sic_descriptions = []
        for code in sic_codes:
            if code in TARGET_SIC_CODES:
                sic_descriptions.append(f"{code}: {TARGET_SIC_CODES[code]}")
            else:
                sic_descriptions.append(f"{code}")

        # Check which filters match
        name_lower = company.get('company_name', '').lower()

        final_companies.append({
            'company_number': company_number,
            'company_name': company.get('company_name', ''),
            'status': 'active',
            'company_type': company.get('company_type', ''),
            'date_of_creation': company.get('date_of_creation', ''),
            'address': full_address,
            'postcode': postcode,
            'locality': locality,
            'region': region,
            'sic_codes': sic_codes,
            'sic_descriptions': sic_descriptions,
            'matched_sic_codes': company.get('matched_sic_codes', []),
            'contains_truck': 'truck' in name_lower,
            'contains_tyre': 'tyre' in name_lower,
            'contains_tire': 'tire' in name_lower,
            'source': 'Companies House API',
            'verified': True
        })

    # Separate by filters
    truck_companies = [c for c in final_companies if c['contains_truck']]
    tyre_companies = [c for c in final_companies if c['contains_tyre']]
    tire_companies = [c for c in final_companies if c['contains_tire']]

    # Print summary
    print(f"\n{'=' * 70}")
    print("FINAL RESULTS SUMMARY")
    print(f"{'=' * 70}")
    print(f"Total unique ACTIVE companies (excl. NI): {len(final_companies):,}")
    print(f"\nBy SIC code:")
    for code, count in sic_code_counts.items():
        print(f"  {code}: {count:,}")
    print(f"\nBy name filter:")
    print(f"  Contains 'truck': {len(truck_companies):,}")
    print(f"  Contains 'tyre': {len(tyre_companies):,}")
    print(f"  Contains 'tire': {len(tire_companies):,}")

    # Save results
    output_data = {
        'metadata': {
            'generated_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'source': 'Companies House API',
            'filter': 'SIC codes - ACTIVE companies - Excludes Northern Ireland',
            'target_sic_codes': TARGET_SIC_CODES,
            'description_filters': DESCRIPTION_FILTERS,
            'total_companies': len(final_companies),
            'companies_per_sic_code': sic_code_counts,
            'companies_with_truck': len(truck_companies),
            'companies_with_tyre': len(tyre_companies),
            'companies_with_tire': len(tire_companies),
        },
        'all_companies': final_companies,
        'truck_filter': truck_companies,
        'tyre_filter': tyre_companies,
        'tire_filter': tire_companies,
    }

    # Save JSON
    with open('UK_RETAILERS_SIC_COMPANIES.json', 'w') as f:
        json.dump(output_data, f, indent=2)
    print(f"\nSaved: UK_RETAILERS_SIC_COMPANIES.json")

    # Save CSV - All companies
    csv_fields = [
        'company_number', 'company_name', 'status', 'company_type',
        'date_of_creation', 'address', 'postcode', 'locality', 'region',
        'sic_codes', 'sic_descriptions', 'matched_sic_codes',
        'contains_truck', 'contains_tyre', 'contains_tire'
    ]

    with open('UK_RETAILERS_SIC_COMPANIES.csv', 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=csv_fields, extrasaction='ignore')
        writer.writeheader()
        for company in final_companies:
            row = company.copy()
            row['sic_codes'] = '; '.join(row.get('sic_codes', [])) if row.get('sic_codes') else ''
            row['sic_descriptions'] = '; '.join(row.get('sic_descriptions', [])) if row.get('sic_descriptions') else ''
            row['matched_sic_codes'] = '; '.join(row.get('matched_sic_codes', [])) if row.get('matched_sic_codes') else ''
            writer.writerow(row)
    print(f"Saved: UK_RETAILERS_SIC_COMPANIES.csv")

    # Save Excel with separate sheets
    try:
        import pandas as pd

        def prepare_df(companies_list):
            df = pd.DataFrame(companies_list)
            if not df.empty:
                df['sic_codes'] = df['sic_codes'].apply(lambda x: '; '.join(x) if x else '')
                df['sic_descriptions'] = df['sic_descriptions'].apply(lambda x: '; '.join(x) if x else '')
                df['matched_sic_codes'] = df['matched_sic_codes'].apply(lambda x: '; '.join(x) if x else '')
            return df

        with pd.ExcelWriter('UK_RETAILERS_SIC_COMPANIES.xlsx', engine='openpyxl') as writer:
            # All companies
            prepare_df(final_companies).to_excel(writer, sheet_name='All Companies', index=False)

            # Truck filter
            if truck_companies:
                prepare_df(truck_companies).to_excel(writer, sheet_name='Contains TRUCK', index=False)

            # Tyre filter
            if tyre_companies:
                prepare_df(tyre_companies).to_excel(writer, sheet_name='Contains TYRE', index=False)

            # Tire filter
            if tire_companies:
                prepare_df(tire_companies).to_excel(writer, sheet_name='Contains TIRE', index=False)

            # By SIC code
            for sic_code in TARGET_SIC_CODES.keys():
                sic_companies = [c for c in final_companies if sic_code in c.get('matched_sic_codes', [])]
                if sic_companies:
                    prepare_df(sic_companies).to_excel(writer, sheet_name=f'SIC {sic_code}', index=False)

        print(f"Saved: UK_RETAILERS_SIC_COMPANIES.xlsx")

    except ImportError:
        print("Note: pandas/openpyxl not available for Excel export")

    print(f"\n{'=' * 70}")
    print("SCRAPING COMPLETE - ALL DATA FROM COMPANIES HOUSE API")
    print(f"{'=' * 70}")

    return final_companies


if __name__ == "__main__":
    companies = main()
