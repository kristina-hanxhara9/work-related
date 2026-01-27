#!/usr/bin/env python3
"""
UK Retailers Scraper - Combined SIC Codes with Name Filters
============================================================
- All SIC codes combined
- 3 separate outputs: TRUCK, TYRE, TIRE (by name filter)
- Excludes Northern Ireland
- Active companies only
"""

import requests
import json
import time
from datetime import datetime
import csv

# Companies House API credentials
API_KEY = "48d17266-ff2e-425f-9b20-7dcc9b25bb79"
BASE_URL = "https://api.company-information.service.gov.uk"

# Target SIC codes (ALL combined)
TARGET_SIC_CODES = {
    "45200": "Maintenance and repair of motor vehicles",
    "45320": "Retail trade of motor vehicle parts and accessories",
    "46900": "Non-specialised wholesale trade",
    "22110": "Manufacture of rubber tyres and tubes; retreading and rebuilding of rubber tyres",
    "45310": "Wholesale trade of motor vehicle parts and accessories",
    "82990": "Other business support service activities n.e.c.",
    "29320": "Manufacture of other parts and accessories for motor vehicles",
    "33170": "Repair and maintenance of other transport equipment",
    "33190": "Repair of other equipment",
}


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
            elif response.status_code == 500:
                print(f"    API limit reached at {len(all_companies):,}")
                break
            else:
                print(f"    API error {response.status_code}")
                break

        except Exception as e:
            print(f"    Error: {e}")
            time.sleep(5)
            continue

    return all_companies


def format_company(company, matched_sic_codes):
    """Format company data for output."""
    addr = company.get('registered_office_address', {})
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

    return {
        'company_number': company.get('company_number', ''),
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
        'matched_sic_codes': matched_sic_codes,
        'source': 'Companies House API',
        'verified': True
    }


def main():
    print("=" * 70)
    print("UK RETAILERS SCRAPER - COMBINED SIC CODES")
    print("=" * 70)
    print("Source: Companies House API ONLY")
    print("Filter: ACTIVE companies only")
    print("Excludes: Northern Ireland")
    print("Output: 3 sheets - TRUCK, TYRE, TIRE")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)

    print("\nTarget SIC codes:")
    for code, desc in TARGET_SIC_CODES.items():
        print(f"  {code}: {desc}")

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
                all_companies[company_number] = company
                all_companies[company_number]['matched_sic_codes'] = [sic_code]
            else:
                if sic_code not in all_companies[company_number].get('matched_sic_codes', []):
                    all_companies[company_number]['matched_sic_codes'].append(sic_code)

        time.sleep(1)

    print(f"{'=' * 70}")
    print(f"Search complete. Found {len(all_companies):,} unique ACTIVE companies (excl. NI)")
    print(f"{'=' * 70}")

    # Process and filter by name
    print("\nFiltering by company name...")

    truck_companies = []
    tyre_companies = []
    tire_companies = []

    for company_number, company in all_companies.items():
        name_lower = company.get('company_name', '').lower()
        matched_sic = company.get('matched_sic_codes', [])
        formatted = format_company(company, matched_sic)

        if 'truck' in name_lower:
            truck_companies.append(formatted)
        if 'tyre' in name_lower:
            tyre_companies.append(formatted)
        if 'tire' in name_lower:
            tire_companies.append(formatted)

    # Print summary
    print(f"\n{'=' * 70}")
    print("FINAL RESULTS SUMMARY")
    print(f"{'=' * 70}")
    print(f"Total unique companies from all SIC codes: {len(all_companies):,}")
    print(f"\nBy SIC code:")
    for code, count in sic_code_counts.items():
        print(f"  {code}: {count:,}")
    print(f"\nFiltered by name:")
    print(f"  TRUCK (name contains 'truck'): {len(truck_companies):,}")
    print(f"  TYRE (name contains 'tyre'): {len(tyre_companies):,}")
    print(f"  TIRE (name contains 'tire'): {len(tire_companies):,}")

    # Save results
    output_data = {
        'metadata': {
            'generated_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'source': 'Companies House API',
            'filter': 'All SIC codes combined - ACTIVE - Excludes NI',
            'target_sic_codes': TARGET_SIC_CODES,
            'total_from_all_sic': len(all_companies),
            'companies_per_sic_code': sic_code_counts,
            'truck_count': len(truck_companies),
            'tyre_count': len(tyre_companies),
            'tire_count': len(tire_companies),
        },
        'truck_companies': truck_companies,
        'tyre_companies': tyre_companies,
        'tire_companies': tire_companies,
    }

    # Save JSON
    with open('UK_RETAILERS_TRUCK_TYRE_TIRE.json', 'w') as f:
        json.dump(output_data, f, indent=2)
    print(f"\nSaved: UK_RETAILERS_TRUCK_TYRE_TIRE.json")

    # Save CSVs - one for each filter
    csv_fields = [
        'company_number', 'company_name', 'status', 'company_type',
        'date_of_creation', 'address', 'postcode', 'locality', 'region',
        'sic_codes', 'sic_descriptions', 'matched_sic_codes'
    ]

    def save_csv(companies, filename):
        with open(filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=csv_fields, extrasaction='ignore')
            writer.writeheader()
            for company in companies:
                row = company.copy()
                row['sic_codes'] = '; '.join(row.get('sic_codes', [])) if row.get('sic_codes') else ''
                row['sic_descriptions'] = '; '.join(row.get('sic_descriptions', [])) if row.get('sic_descriptions') else ''
                row['matched_sic_codes'] = '; '.join(row.get('matched_sic_codes', [])) if row.get('matched_sic_codes') else ''
                writer.writerow(row)

    save_csv(truck_companies, 'UK_RETAILERS_TRUCK.csv')
    save_csv(tyre_companies, 'UK_RETAILERS_TYRE.csv')
    save_csv(tire_companies, 'UK_RETAILERS_TIRE.csv')
    print(f"Saved: UK_RETAILERS_TRUCK.csv ({len(truck_companies):,} companies)")
    print(f"Saved: UK_RETAILERS_TYRE.csv ({len(tyre_companies):,} companies)")
    print(f"Saved: UK_RETAILERS_TIRE.csv ({len(tire_companies):,} companies)")

    # Save Excel with 3 sheets
    try:
        import pandas as pd

        def prepare_df(companies_list):
            df = pd.DataFrame(companies_list)
            if not df.empty:
                df['sic_codes'] = df['sic_codes'].apply(lambda x: '; '.join(x) if x else '')
                df['sic_descriptions'] = df['sic_descriptions'].apply(lambda x: '; '.join(x) if x else '')
                df['matched_sic_codes'] = df['matched_sic_codes'].apply(lambda x: '; '.join(x) if x else '')
            return df

        with pd.ExcelWriter('UK_RETAILERS_TRUCK_TYRE_TIRE.xlsx', engine='openpyxl') as writer:
            # TRUCK sheet
            if truck_companies:
                prepare_df(truck_companies).to_excel(writer, sheet_name='TRUCK', index=False)

            # TYRE sheet
            if tyre_companies:
                prepare_df(tyre_companies).to_excel(writer, sheet_name='TYRE', index=False)

            # TIRE sheet
            if tire_companies:
                prepare_df(tire_companies).to_excel(writer, sheet_name='TIRE', index=False)

        print(f"Saved: UK_RETAILERS_TRUCK_TYRE_TIRE.xlsx (3 sheets)")

    except ImportError:
        print("Note: pandas/openpyxl not available for Excel export")

    print(f"\n{'=' * 70}")
    print("SCRAPING COMPLETE - ALL DATA FROM COMPANIES HOUSE API")
    print(f"{'=' * 70}")

    return truck_companies, tyre_companies, tire_companies


if __name__ == "__main__":
    truck, tyre, tire = main()
