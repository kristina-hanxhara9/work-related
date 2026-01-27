#!/usr/bin/env python3
"""
UK Companies House SIC Code Scraper
====================================
Searches for ALL companies with specific SIC codes.
100% from Companies House API - no filters, just SIC codes.
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
    "33190": "Repair of other equipment",
}


def search_by_sic_code(sic_code):
    """
    Search Companies House for ALL companies with a specific SIC code.
    Uses the advanced search endpoint with proper pagination.
    """
    all_companies = []
    start_index = 0
    items_per_page = 500  # Max allowed by API
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
                # Use 'hits' field, not 'total_results'
                total_results = data.get('hits', 0)

                if total_available is None:
                    total_available = total_results
                    print(f"    Total available: {total_available:,} active companies")

                if not items:
                    break

                all_companies.extend(items)

                # Progress update
                print(f"    Fetched: {len(all_companies):,}/{total_available:,}")

                # Check if we've got all results
                if len(all_companies) >= total_available:
                    break

                # If we got fewer items than requested, we're at the end
                if len(items) < items_per_page:
                    break

                # Move to next page
                start_index += items_per_page
                time.sleep(0.2)  # Rate limiting

            elif response.status_code == 416:
                # Range not satisfiable - we've reached the end
                break
            elif response.status_code == 429:
                print("    Rate limited. Waiting 60 seconds...")
                time.sleep(60)
                # Retry same request
                continue
            else:
                print(f"    API error {response.status_code}: {response.text[:100]}")
                break

        except Exception as e:
            print(f"    Error: {e}")
            time.sleep(5)
            continue

    return all_companies


def main():
    print("=" * 70)
    print("UK COMPANIES HOUSE - SIC CODE SCRAPER")
    print("=" * 70)
    print("Source: Companies House API ONLY")
    print("Filter: ACTIVE companies only")
    print("Getting ALL companies - no limits")
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

        print(f"    TOTAL Retrieved: {len(companies):,} active companies\n")

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
                # Company already exists, add this SIC code to matched list
                if sic_code not in all_companies[company_number]['matched_sic_codes']:
                    all_companies[company_number]['matched_sic_codes'].append(sic_code)

        time.sleep(1)  # Pause between SIC codes

    print(f"{'=' * 70}")
    print(f"Search complete. Found {len(all_companies):,} unique ACTIVE companies")
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

        final_companies.append({
            'company_number': company_number,
            'company_name': company.get('company_name', ''),
            'status': company.get('status', 'active'),
            'company_type': company.get('company_type', ''),
            'date_of_creation': company.get('date_of_creation', ''),
            'address': full_address,
            'postcode': postcode,
            'locality': locality,
            'region': region,
            'sic_codes': sic_codes,
            'sic_descriptions': sic_descriptions,
            'matched_sic_codes': company.get('matched_sic_codes', []),
            'source': 'Companies House API',
            'verified': True
        })

    # Print summary
    print(f"\n{'=' * 70}")
    print("FINAL RESULTS SUMMARY")
    print(f"{'=' * 70}")
    print(f"Total unique ACTIVE companies: {len(final_companies):,}")
    print(f"\nCompanies by SIC code:")
    for code, count in sic_code_counts.items():
        desc_short = TARGET_SIC_CODES[code][:45] + "..." if len(TARGET_SIC_CODES[code]) > 45 else TARGET_SIC_CODES[code]
        print(f"  {code}: {count:,} - {desc_short}")

    # Save results
    output_data = {
        'metadata': {
            'generated_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'source': 'Companies House API',
            'filter': 'SIC codes only - ACTIVE companies',
            'target_sic_codes': TARGET_SIC_CODES,
            'total_companies': len(final_companies),
            'companies_per_sic_code': sic_code_counts
        },
        'all_companies': final_companies
    }

    # Save JSON
    with open('UK_RETAILERS_BY_SIC_CODE.json', 'w') as f:
        json.dump(output_data, f, indent=2)
    print(f"\nSaved: UK_RETAILERS_BY_SIC_CODE.json")

    # Save CSV
    csv_fields = [
        'company_number', 'company_name', 'status', 'company_type',
        'date_of_creation', 'address', 'postcode', 'locality', 'region',
        'sic_codes', 'sic_descriptions', 'matched_sic_codes'
    ]

    with open('UK_RETAILERS_BY_SIC_CODE.csv', 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=csv_fields, extrasaction='ignore')
        writer.writeheader()
        for company in final_companies:
            row = company.copy()
            row['sic_codes'] = '; '.join(row.get('sic_codes', [])) if row.get('sic_codes') else ''
            row['sic_descriptions'] = '; '.join(row.get('sic_descriptions', [])) if row.get('sic_descriptions') else ''
            row['matched_sic_codes'] = '; '.join(row.get('matched_sic_codes', [])) if row.get('matched_sic_codes') else ''
            writer.writerow(row)
    print(f"Saved: UK_RETAILERS_BY_SIC_CODE.csv")

    # Save Excel
    try:
        import pandas as pd

        df = pd.DataFrame(final_companies)
        df['sic_codes'] = df['sic_codes'].apply(lambda x: '; '.join(x) if x else '')
        df['sic_descriptions'] = df['sic_descriptions'].apply(lambda x: '; '.join(x) if x else '')
        df['matched_sic_codes'] = df['matched_sic_codes'].apply(lambda x: '; '.join(x) if x else '')

        with pd.ExcelWriter('UK_RETAILERS_BY_SIC_CODE.xlsx', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='All Companies', index=False)

            # Create separate sheets per SIC code
            for sic_code, description in TARGET_SIC_CODES.items():
                sic_companies = [c for c in final_companies if sic_code in c.get('matched_sic_codes', [])]
                if sic_companies:
                    df_sic = pd.DataFrame(sic_companies)
                    df_sic['sic_codes'] = df_sic['sic_codes'].apply(lambda x: '; '.join(x) if x else '')
                    df_sic['sic_descriptions'] = df_sic['sic_descriptions'].apply(lambda x: '; '.join(x) if x else '')
                    df_sic['matched_sic_codes'] = df_sic['matched_sic_codes'].apply(lambda x: '; '.join(x) if x else '')
                    sheet_name = f"SIC {sic_code}"[:31]
                    df_sic.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Saved: UK_RETAILERS_BY_SIC_CODE.xlsx")

    except ImportError:
        print("Note: pandas/openpyxl not available for Excel export")

    print(f"\n{'=' * 70}")
    print("SCRAPING COMPLETE - ALL DATA FROM COMPANIES HOUSE API")
    print(f"{'=' * 70}")

    return final_companies


if __name__ == "__main__":
    companies = main()
