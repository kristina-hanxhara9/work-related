#!/usr/bin/env python3
"""
Add SIC code descriptions to existing data without re-scraping.
"""

import json

# SIC code descriptions (UK standard)
SIC_DESCRIPTIONS = {
    "01110": "Growing of cereals (except rice), leguminous crops and oil seeds",
    "01610": "Support activities for crop production",
    "45110": "Sale of cars and light motor vehicles",
    "45190": "Sale of other motor vehicles",
    "45200": "Maintenance and repair of motor vehicles",
    "45310": "Wholesale trade of motor vehicle parts and accessories",
    "45320": "Retail trade of motor vehicle parts and accessories",
    "45400": "Sale, maintenance and repair of motorcycles and related parts and accessories",
    "46690": "Wholesale of other machinery and equipment",
    "46900": "Non-specialised wholesale trade",
    "47300": "Retail sale of automotive fuel in specialised stores",
    "47990": "Other retail sale not in stores, stalls or markets",
    "49410": "Freight transport by road",
    "52100": "Warehousing and storage",
    "52210": "Service activities incidental to land transportation",
    "52290": "Other transportation support activities",
    "66220": "Activities of insurance agents and brokers",
    "70100": "Activities of head offices",
    "70210": "Public relations and communication activities",
    "70229": "Management consultancy activities other than financial management",
    "71200": "Technical testing and analysis",
    "74909": "Other professional, scientific and technical activities n.e.c.",
    "77110": "Renting and leasing of cars and light motor vehicles",
    "77120": "Renting and leasing of trucks",
    "77390": "Renting and leasing of other machinery, equipment and tangible goods n.e.c.",
    "81210": "General cleaning of buildings",
    "82990": "Other business support service activities n.e.c.",
    "95120": "Repair of communication equipment",
    "96090": "Other personal service activities n.e.c.",
}

def get_sic_description(code):
    """Get description for a SIC code."""
    if code in SIC_DESCRIPTIONS:
        return SIC_DESCRIPTIONS[code]

    # Try to provide generic descriptions based on code prefix
    prefix = code[:2] if len(code) >= 2 else code

    prefix_descriptions = {
        "01": "Agriculture, forestry and fishing",
        "45": "Wholesale and retail trade; repair of motor vehicles",
        "46": "Wholesale trade, except of motor vehicles",
        "47": "Retail trade, except of motor vehicles",
        "49": "Land transport and transport via pipelines",
        "52": "Warehousing and support activities for transportation",
        "66": "Activities auxiliary to financial services",
        "70": "Activities of head offices; management consultancy",
        "71": "Architectural and engineering activities",
        "74": "Other professional, scientific and technical activities",
        "77": "Rental and leasing activities",
        "81": "Services to buildings and landscape activities",
        "82": "Office administrative and business support activities",
        "95": "Repair of computers and personal and household goods",
        "96": "Other personal service activities",
    }

    return prefix_descriptions.get(prefix, f"SIC code {code}")


def main():
    print("Adding SIC descriptions to existing data...")

    # Load existing data
    with open('UK_TYRE_COMPANIES_API_ONLY.json', 'r') as f:
        data = json.load(f)

    # Add descriptions to all companies
    for company in data['all_companies']:
        sic_codes = company.get('sic_codes', [])
        sic_descriptions = []
        for code in sic_codes:
            desc = get_sic_description(code)
            sic_descriptions.append(f"{code}: {desc}")
        company['sic_descriptions'] = sic_descriptions

    # Update active_only, truck_commercial, mobile_services too
    for company in data.get('active_only', []):
        sic_codes = company.get('sic_codes', [])
        sic_descriptions = []
        for code in sic_codes:
            desc = get_sic_description(code)
            sic_descriptions.append(f"{code}: {desc}")
        company['sic_descriptions'] = sic_descriptions

    for company in data.get('truck_commercial', []):
        sic_codes = company.get('sic_codes', [])
        sic_descriptions = []
        for code in sic_codes:
            desc = get_sic_description(code)
            sic_descriptions.append(f"{code}: {desc}")
        company['sic_descriptions'] = sic_descriptions

    for company in data.get('mobile_services', []):
        sic_codes = company.get('sic_codes', [])
        sic_descriptions = []
        for code in sic_codes:
            desc = get_sic_description(code)
            sic_descriptions.append(f"{code}: {desc}")
        company['sic_descriptions'] = sic_descriptions

    # Save updated JSON
    with open('UK_TYRE_COMPANIES_API_ONLY.json', 'w') as f:
        json.dump(data, f, indent=2)
    print("Updated: UK_TYRE_COMPANIES_API_ONLY.json")

    # Update CSV
    import csv

    active_companies = data.get('active_only', [])

    csv_fields = [
        'company_number', 'company_name', 'status', 'company_type',
        'date_of_creation', 'address', 'postcode', 'locality', 'region',
        'sic_codes', 'sic_descriptions', 'is_truck_commercial', 'is_mobile', 'categories'
    ]

    with open('UK_TYRE_COMPANIES_API_ONLY.csv', 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=csv_fields, extrasaction='ignore')
        writer.writeheader()
        for company in active_companies:
            row = company.copy()
            row['sic_codes'] = '; '.join(row.get('sic_codes', [])) if row.get('sic_codes') else ''
            row['sic_descriptions'] = '; '.join(row.get('sic_descriptions', [])) if row.get('sic_descriptions') else ''
            row['categories'] = '; '.join(row.get('categories', [])) if row.get('categories') else ''
            writer.writerow(row)
    print("Updated: UK_TYRE_COMPANIES_API_ONLY.csv")

    # Update Excel
    try:
        import pandas as pd

        df = pd.DataFrame(active_companies)
        df['sic_codes'] = df['sic_codes'].apply(lambda x: '; '.join(x) if x else '')
        df['sic_descriptions'] = df['sic_descriptions'].apply(lambda x: '; '.join(x) if x else '')
        df['categories'] = df['categories'].apply(lambda x: '; '.join(x) if x else '')

        truck_commercial = data.get('truck_commercial', [])
        mobile_services = data.get('mobile_services', [])

        with pd.ExcelWriter('UK_TYRE_COMPANIES_API_ONLY.xlsx', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='All Active Companies', index=False)

            if truck_commercial:
                df_truck = pd.DataFrame(truck_commercial)
                df_truck['sic_codes'] = df_truck['sic_codes'].apply(lambda x: '; '.join(x) if x else '')
                df_truck['sic_descriptions'] = df_truck['sic_descriptions'].apply(lambda x: '; '.join(x) if x else '')
                df_truck['categories'] = df_truck['categories'].apply(lambda x: '; '.join(x) if x else '')
                df_truck.to_excel(writer, sheet_name='Truck Commercial', index=False)

            if mobile_services:
                df_mobile = pd.DataFrame(mobile_services)
                df_mobile['sic_codes'] = df_mobile['sic_codes'].apply(lambda x: '; '.join(x) if x else '')
                df_mobile['sic_descriptions'] = df_mobile['sic_descriptions'].apply(lambda x: '; '.join(x) if x else '')
                df_mobile['categories'] = df_mobile['categories'].apply(lambda x: '; '.join(x) if x else '')
                df_mobile.to_excel(writer, sheet_name='Mobile Services', index=False)

        print("Updated: UK_TYRE_COMPANIES_API_ONLY.xlsx")

    except ImportError:
        print("Note: pandas/openpyxl not available for Excel export")

    print("\nDone! SIC descriptions added to all files.")


if __name__ == "__main__":
    main()
