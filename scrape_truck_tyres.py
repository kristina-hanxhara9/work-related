"""
UK TRUCK TYRE SCRAPER v2
========================
Scrapes real truck tyre companies from multiple sources.
"""

import requests
from bs4 import BeautifulSoup
import json
import time
import re
from urllib.parse import urljoin, urlparse

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
}

COMPANIES = []

# Known major UK truck tyre companies/networks to scrape
KNOWN_TRUCK_TYRE_SITES = [
    # Major networks
    ('https://www.247mobiletrucktyres.co.uk/', '247 Mobile Truck Tyres'),
    ('https://tyrenet.net/', 'Tyrenet'),
    ('https://www.tructyre.co.uk/', 'Tructyre ATS'),
    ('https://www.bandvulc.co.uk/', 'Bandvulc'),
    ('https://www.michelin-trucktires.co.uk/', 'Michelin Truck'),
    ('https://www.bridgestone.co.uk/truck', 'Bridgestone Truck'),
    ('https://www.continental-tyres.co.uk/truck', 'Continental Truck'),
    ('https://www.goodyear.eu/en_gb/truck.html', 'Goodyear Truck'),

    # Regional specialists
    ('https://www.mobiletyrefittinguk.co.uk/', 'Mobile Tyre Fitting UK'),
    ('https://www.tyreassist365.com/', 'Tyre Assist 365'),
    ('https://www.mobiletyrebuddy.co.uk/mobile-truck-tyre-fitting/', 'Mobile Tyre Buddy'),
    ('https://www.mcconechy.co.uk/', 'McConechy Tyres'),
    ('https://www.mts-truck-tyres.co.uk/', 'MTS Truck Tyres'),
    ('https://www.lodgewaytyres.co.uk/', 'Lodgeway Tyres'),
    ('https://www.stapletonstyreservices.co.uk/', 'Stapletons Tyre Services'),
]


def scrape_site(url, expected_name):
    """Scrape a single truck tyre company website"""
    print(f"  Scraping: {expected_name}...")

    try:
        r = requests.get(url, headers=HEADERS, timeout=15, allow_redirects=True)

        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')
            text = r.text

            # Extract phone number
            phone = None
            phone_patterns = [
                r'0800[\s\-]?\d{3}[\s\-]?\d{3,4}',  # Freephone
                r'0\d{2,4}[\s\-]?\d{3}[\s\-]?\d{3,4}',  # Landline
                r'\+44[\s\-]?\d{2,4}[\s\-]?\d{3}[\s\-]?\d{3,4}',  # International
            ]
            for pattern in phone_patterns:
                match = re.search(pattern, text)
                if match:
                    phone = match.group().strip()
                    break

            # Extract email
            email = None
            email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text)
            if email_match:
                email = email_match.group()
                if email.endswith(('.png', '.jpg', '.svg')):
                    email = None

            # Extract description/services
            description = ''
            meta_desc = soup.find('meta', {'name': 'description'})
            if meta_desc:
                description = meta_desc.get('content', '')[:200]

            # Check if it's actually truck tyres
            truck_indicators = ['truck', 'hgv', 'lorry', 'commercial', 'fleet', 'trailer']
            text_lower = text.lower()
            is_truck = any(ind in text_lower for ind in truck_indicators)

            if is_truck:
                return {
                    'name': expected_name,
                    'website': url,
                    'phone': phone,
                    'email': email,
                    'description': description,
                    'verified_truck': True,
                    'source': 'Direct website scrape'
                }
            else:
                print(f"    WARNING: {expected_name} may not be truck-specific")
                return None

        else:
            print(f"    Status: {r.status_code}")
            return None

    except Exception as e:
        print(f"    Error: {str(e)[:50]}")
        return None


def scrape_bing_search(query):
    """Scrape Bing search results for truck tyre companies"""
    print(f"\n  Searching Bing: '{query}'")

    companies = []

    try:
        url = f'https://www.bing.com/search?q={query.replace(" ", "+")}'
        r = requests.get(url, headers=HEADERS, timeout=15)

        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')

            # Find search results
            results = soup.find_all('li', class_='b_algo')

            for result in results[:10]:
                try:
                    # Get title and URL
                    link = result.find('a')
                    if link:
                        title = link.get_text(strip=True)
                        href = link.get('href', '')

                        # Filter for relevant results
                        if href and 'truck' in title.lower() or 'hgv' in title.lower() or 'commercial tyre' in title.lower():
                            companies.append({
                                'name': title[:60],
                                'website': href,
                                'source': 'Bing search'
                            })
                except:
                    continue

            print(f"    Found {len(companies)} results")

    except Exception as e:
        print(f"    Error: {e}")

    return companies


def scrape_companies_house(search_term):
    """Search Companies House for truck tyre companies"""
    print(f"\n  Searching Companies House: '{search_term}'")

    companies = []
    api_key = '48d17266-ff2e-425f-9b20-7dcc9b25bb79'

    try:
        url = f'https://api.company-information.service.gov.uk/search/companies'
        params = {'q': search_term, 'items_per_page': 50}

        r = requests.get(url, auth=(api_key, ''), params=params, timeout=15)

        if r.status_code == 200:
            data = r.json()
            items = data.get('items', [])

            for item in items:
                status = item.get('company_status', '')
                if status == 'active':
                    companies.append({
                        'name': item.get('title', ''),
                        'company_number': item.get('company_number', ''),
                        'address': item.get('address_snippet', ''),
                        'status': status,
                        'date_created': item.get('date_of_creation', ''),
                        'source': 'Companies House'
                    })

            print(f"    Found {len(companies)} active companies")

    except Exception as e:
        print(f"    Error: {e}")

    return companies


def main():
    print("=" * 70)
    print("UK TRUCK TYRE COMPANIES - WEB SCRAPER")
    print("=" * 70)

    all_companies = []

    # 1. Scrape known truck tyre websites
    print("\n[1] SCRAPING KNOWN TRUCK TYRE WEBSITES")
    print("-" * 50)

    for url, name in KNOWN_TRUCK_TYRE_SITES:
        result = scrape_site(url, name)
        if result:
            all_companies.append(result)
        time.sleep(1)

    # 2. Companies House search
    print("\n[2] SEARCHING COMPANIES HOUSE")
    print("-" * 50)

    ch_searches = [
        'truck tyre',
        'truck tyres',
        'hgv tyre',
        'lorry tyre',
        'commercial tyre'
    ]

    for term in ch_searches:
        results = scrape_companies_house(term)
        all_companies.extend(results)
        time.sleep(0.5)

    # 3. Bing search for more companies
    print("\n[3] SEARCHING BING")
    print("-" * 50)

    bing_searches = [
        'UK truck tyre companies',
        'UK HGV tyre fitters',
        'mobile truck tyre fitting UK',
        'commercial truck tyres UK',
    ]

    for query in bing_searches:
        results = scrape_bing_search(query)
        all_companies.extend(results)
        time.sleep(2)

    # Deduplicate by name
    seen = set()
    unique = []
    for c in all_companies:
        name_key = c.get('name', '').lower().strip()
        if name_key and name_key not in seen:
            seen.add(name_key)
            unique.append(c)

    print(f"\n\n{'=' * 70}")
    print(f"RESULTS SUMMARY")
    print(f"{'=' * 70}")
    print(f"Total scraped: {len(all_companies)}")
    print(f"After deduplication: {len(unique)}")

    # Save results
    with open('TRUCK_TYRE_COMPANIES_SCRAPED.json', 'w') as f:
        json.dump(unique, f, indent=2)

    print(f"\nSaved to: TRUCK_TYRE_COMPANIES_SCRAPED.json")

    # Print results
    print(f"\n\n{'=' * 70}")
    print("SCRAPED TRUCK TYRE COMPANIES")
    print(f"{'=' * 70}")

    for i, c in enumerate(unique[:50], 1):  # Show first 50
        print(f"\n{i}. {c.get('name', 'Unknown')}")
        if c.get('website'):
            print(f"   Website: {c['website']}")
        if c.get('phone'):
            print(f"   Phone: {c['phone']}")
        if c.get('email'):
            print(f"   Email: {c['email']}")
        if c.get('company_number'):
            print(f"   Company #: {c['company_number']}")
        if c.get('address'):
            print(f"   Address: {c['address']}")
        print(f"   Source: {c.get('source', 'Unknown')}")

    if len(unique) > 50:
        print(f"\n... and {len(unique) - 50} more (see JSON file)")

    return unique


if __name__ == "__main__":
    main()
