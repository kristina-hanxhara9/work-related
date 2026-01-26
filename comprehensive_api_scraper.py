#!/usr/bin/env python3
"""
Comprehensive UK Truck Tyre Companies Scraper
============================================
This scraper ONLY uses the Companies House API - no made-up data.
Every single company in the output is verified to exist on Companies House.

Searches for:
- Truck tyres
- Commercial tyres
- HGV tyres
- Mobile tyres
- Fleet tyres
- Tyre fitting/services
- And many more variations
"""

import requests
import json
import time
from datetime import datetime
import re

# Companies House API credentials
API_KEY = "48d17266-ff2e-425f-9b20-7dcc9b25bb79"
BASE_URL = "https://api.company-information.service.gov.uk"

# Search terms to find all truck/commercial tyre companies
SEARCH_TERMS = [
    # Core truck/commercial terms
    "truck tyre",
    "truck tyres",
    "commercial tyre",
    "commercial tyres",
    "hgv tyre",
    "hgv tyres",
    "lorry tyre",
    "lorry tyres",

    # Mobile services
    "mobile tyre",
    "mobile tyres",
    "24 hour tyre",
    "24hr tyre",
    "emergency tyre",

    # Fleet services
    "fleet tyre",
    "fleet tyres",
    "fleet service tyre",

    # Tyre fitting/services
    "tyre fitting",
    "tyre service",
    "tyre services",
    "tyre centre",
    "tyre center",

    # Retreading/remoulding
    "retread",
    "retreading",
    "remould",
    "tyre remould",

    # Wholesale/distribution
    "tyre wholesale",
    "tyre distribution",
    "tyre supplies",
    "tyre supply",

    # General tyre businesses
    "tyres ltd",
    "tyres limited",
    "tyre ltd",
    "tyre limited",

    # Variations with 'tire' spelling
    "truck tire",
    "commercial tire",
    "mobile tire",

    # Specific business types
    "tyre dealer",
    "tyre traders",
    "tyre specialist",
    "tyre depot",
    "tyre warehouse",

    # Vehicle specific
    "van tyre",
    "trailer tyre",
    "bus tyre",
    "coach tyre",
]

# Keywords to identify truck/commercial tyre relevance
TRUCK_KEYWORDS = [
    'truck', 'hgv', 'lgv', 'commercial', 'lorry', 'fleet', 'trailer',
    'bus', 'coach', 'van', 'vehicle', 'transport', 'haulage', 'freight'
]

TYRE_KEYWORDS = [
    'tyre', 'tyres', 'tire', 'tires', 'wheel', 'retread', 'remould'
]

MOBILE_KEYWORDS = [
    'mobile', '24 hour', '24hr', '24/7', 'emergency', 'roadside', 'callout'
]


def search_companies(query, max_results=500):
    """Search Companies House API for companies matching query."""
    all_results = []
    start_index = 0
    items_per_page = 100  # API maximum

    while start_index < max_results:
        url = f"{BASE_URL}/search/companies"
        params = {
            "q": query,
            "items_per_page": items_per_page,
            "start_index": start_index
        }

        try:
            response = requests.get(url, params=params, auth=(API_KEY, ''))

            if response.status_code == 200:
                data = response.json()
                items = data.get('items', [])

                if not items:
                    break

                all_results.extend(items)

                total_results = data.get('total_results', 0)
                print(f"  Query '{query}': Found {len(items)} (total available: {total_results})")

                if start_index + items_per_page >= total_results:
                    break

                start_index += items_per_page
                time.sleep(0.2)  # Rate limiting
            else:
                print(f"  Query '{query}': API error {response.status_code}")
                break

        except Exception as e:
            print(f"  Query '{query}': Error - {e}")
            break

    return all_results


def get_company_profile(company_number):
    """Get full company profile from API."""
    url = f"{BASE_URL}/company/{company_number}"

    try:
        response = requests.get(url, auth=(API_KEY, ''))
        if response.status_code == 200:
            return response.json()
    except:
        pass

    return None


def is_tyre_related(company_name, sic_codes=None):
    """Check if company is tyre related based on name and SIC codes."""
    name_lower = company_name.lower()

    # Check for tyre keywords in name
    has_tyre_keyword = any(kw in name_lower for kw in TYRE_KEYWORDS)

    if has_tyre_keyword:
        return True

    # Check SIC codes for tyre-related activities
    tyre_sic_codes = ['45310', '45320', '45400', '45200', '77110']
    if sic_codes:
        for code in sic_codes:
            if code in tyre_sic_codes:
                return True

    return False


def is_truck_commercial(company_name):
    """Check if company is specifically truck/commercial focused."""
    name_lower = company_name.lower()
    return any(kw in name_lower for kw in TRUCK_KEYWORDS)


def is_mobile_service(company_name):
    """Check if company offers mobile services."""
    name_lower = company_name.lower()
    return any(kw in name_lower for kw in MOBILE_KEYWORDS)


def categorize_company(company_name, sic_codes=None):
    """Categorize the company type."""
    name_lower = company_name.lower()

    categories = []

    # Check truck/commercial
    if is_truck_commercial(company_name):
        categories.append('truck_commercial')

    # Check mobile
    if is_mobile_service(company_name):
        categories.append('mobile')

    # Check retread
    if 'retread' in name_lower or 'remould' in name_lower:
        categories.append('retread')

    # Check wholesale/distribution
    if any(kw in name_lower for kw in ['wholesale', 'distribution', 'supplies', 'supply', 'warehouse']):
        categories.append('wholesale')

    # Check fitting/service
    if any(kw in name_lower for kw in ['fitting', 'service', 'centre', 'center', 'depot']):
        categories.append('fitting_service')

    # Check fleet
    if 'fleet' in name_lower:
        categories.append('fleet')

    # Default to general tyre
    if not categories:
        categories.append('general_tyre')

    return categories


def main():
    print("=" * 70)
    print("COMPREHENSIVE UK TRUCK TYRE COMPANIES SCRAPER")
    print("=" * 70)
    print("Source: Companies House API ONLY - 100% verified data")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)

    # Collect all search results
    all_companies = {}

    print(f"\nSearching {len(SEARCH_TERMS)} different terms...\n")

    for i, term in enumerate(SEARCH_TERMS, 1):
        print(f"[{i}/{len(SEARCH_TERMS)}] Searching: '{term}'")
        results = search_companies(term, max_results=300)

        for company in results:
            company_number = company.get('company_number')
            company_name = company.get('title', '')

            # Skip if already found or not tyre related
            if company_number in all_companies:
                continue

            # Only include if tyre-related
            if not is_tyre_related(company_name):
                continue

            all_companies[company_number] = {
                'company_number': company_number,
                'name': company_name,
                'status': company.get('company_status', 'unknown'),
                'address_snippet': company.get('address_snippet', ''),
                'date_of_creation': company.get('date_of_creation', ''),
                'company_type': company.get('company_type', ''),
            }

        # Rate limiting
        time.sleep(0.3)

    print(f"\n{'=' * 70}")
    print(f"Initial search complete. Found {len(all_companies)} unique tyre companies")
    print(f"{'=' * 70}")

    # Now enrich each company with full profile data
    print("\nEnriching company data from API profiles...")

    enriched_companies = []

    for i, (company_number, company) in enumerate(all_companies.items(), 1):
        if i % 50 == 0:
            print(f"  Progress: {i}/{len(all_companies)}")

        # Get full profile
        profile = get_company_profile(company_number)

        if profile:
            # Extract SIC codes
            sic_codes = profile.get('sic_codes', [])

            # Get registered address
            reg_address = profile.get('registered_office_address', {})
            address_parts = []
            for field in ['address_line_1', 'address_line_2', 'locality', 'region', 'postal_code']:
                if reg_address.get(field):
                    address_parts.append(reg_address[field])
            full_address = ', '.join(address_parts)

            # Categorize
            categories = categorize_company(company['name'], sic_codes)

            enriched = {
                'company_number': company_number,
                'company_name': profile.get('company_name', company['name']),
                'status': profile.get('company_status', company['status']),
                'company_type': profile.get('type', company['company_type']),
                'date_of_creation': profile.get('date_of_creation', company['date_of_creation']),
                'address': full_address,
                'postcode': reg_address.get('postal_code', ''),
                'locality': reg_address.get('locality', ''),
                'region': reg_address.get('region', ''),
                'country': reg_address.get('country', 'United Kingdom'),
                'sic_codes': sic_codes,
                'is_truck_commercial': is_truck_commercial(company['name']),
                'is_mobile': is_mobile_service(company['name']),
                'categories': categories,
                'source': 'Companies House API',
                'verified': True,
                'verified_date': datetime.now().strftime('%Y-%m-%d')
            }

            enriched_companies.append(enriched)

        time.sleep(0.15)  # Rate limiting

    print(f"\nEnrichment complete. {len(enriched_companies)} companies with full data.")

    # Separate by status
    active_companies = [c for c in enriched_companies if c['status'] == 'active']
    dissolved_companies = [c for c in enriched_companies if c['status'] != 'active']

    # Separate by type
    truck_commercial = [c for c in active_companies if c['is_truck_commercial']]
    mobile = [c for c in active_companies if c['is_mobile']]
    general_tyre = [c for c in active_companies if not c['is_truck_commercial'] and not c['is_mobile']]

    # Print summary
    print(f"\n{'=' * 70}")
    print("FINAL RESULTS SUMMARY")
    print(f"{'=' * 70}")
    print(f"Total unique tyre companies found: {len(enriched_companies)}")
    print(f"  - Active companies: {len(active_companies)}")
    print(f"  - Dissolved/Other: {len(dissolved_companies)}")
    print(f"\nActive companies breakdown:")
    print(f"  - Truck/Commercial focused: {len(truck_commercial)}")
    print(f"  - Mobile tyre services: {len(mobile)}")
    print(f"  - General tyre companies: {len(general_tyre)}")

    # Save results
    output_data = {
        'metadata': {
            'generated_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'source': 'Companies House API',
            'total_companies': len(enriched_companies),
            'active_companies': len(active_companies),
            'truck_commercial_count': len(truck_commercial),
            'mobile_count': len(mobile),
            'search_terms_used': SEARCH_TERMS
        },
        'all_companies': enriched_companies,
        'active_only': active_companies,
        'truck_commercial': truck_commercial,
        'mobile_services': mobile
    }

    # Save JSON
    with open('UK_TYRE_COMPANIES_API_ONLY.json', 'w') as f:
        json.dump(output_data, f, indent=2)
    print(f"\nSaved: UK_TYRE_COMPANIES_API_ONLY.json")

    # Save CSV for easy viewing
    import csv

    csv_fields = [
        'company_number', 'company_name', 'status', 'company_type',
        'date_of_creation', 'address', 'postcode', 'locality', 'region',
        'sic_codes', 'is_truck_commercial', 'is_mobile', 'categories'
    ]

    with open('UK_TYRE_COMPANIES_API_ONLY.csv', 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=csv_fields, extrasaction='ignore')
        writer.writeheader()
        for company in active_companies:
            row = company.copy()
            row['sic_codes'] = '; '.join(row['sic_codes']) if row['sic_codes'] else ''
            row['categories'] = '; '.join(row['categories']) if row['categories'] else ''
            writer.writerow(row)
    print(f"Saved: UK_TYRE_COMPANIES_API_ONLY.csv (active companies only)")

    # Save Excel
    try:
        import pandas as pd

        df = pd.DataFrame(active_companies)
        df['sic_codes'] = df['sic_codes'].apply(lambda x: '; '.join(x) if x else '')
        df['categories'] = df['categories'].apply(lambda x: '; '.join(x) if x else '')

        with pd.ExcelWriter('UK_TYRE_COMPANIES_API_ONLY.xlsx', engine='openpyxl') as writer:
            # All active companies
            df.to_excel(writer, sheet_name='All Active Companies', index=False)

            # Truck/Commercial only
            df_truck = pd.DataFrame(truck_commercial)
            if not df_truck.empty:
                df_truck['sic_codes'] = df_truck['sic_codes'].apply(lambda x: '; '.join(x) if x else '')
                df_truck['categories'] = df_truck['categories'].apply(lambda x: '; '.join(x) if x else '')
                df_truck.to_excel(writer, sheet_name='Truck Commercial', index=False)

            # Mobile only
            df_mobile = pd.DataFrame(mobile)
            if not df_mobile.empty:
                df_mobile['sic_codes'] = df_mobile['sic_codes'].apply(lambda x: '; '.join(x) if x else '')
                df_mobile['categories'] = df_mobile['categories'].apply(lambda x: '; '.join(x) if x else '')
                df_mobile.to_excel(writer, sheet_name='Mobile Services', index=False)

        print(f"Saved: UK_TYRE_COMPANIES_API_ONLY.xlsx")

    except ImportError:
        print("Note: pandas/openpyxl not available for Excel export")

    print(f"\n{'=' * 70}")
    print("SCRAPING COMPLETE - ALL DATA FROM COMPANIES HOUSE API")
    print(f"{'=' * 70}")

    return enriched_companies


if __name__ == "__main__":
    companies = main()
