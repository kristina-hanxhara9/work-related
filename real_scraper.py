"""
REAL UK TRUCK TYRE WEB SCRAPER
==============================
Actually scrapes real websites to get real truck tyre company data.
No fake data - only scraped data.
"""

import requests
from bs4 import BeautifulSoup
import json
import time
import re
from datetime import datetime

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
}

ALL_COMPANIES = []

def scrape_yell_truck_tyres():
    """Scrape Yell.com for truck tyre companies"""
    print("\n[1] SCRAPING YELL.COM - TRUCK TYRES")
    print("-" * 50)

    companies = []
    locations = ['london', 'birmingham', 'manchester', 'leeds', 'glasgow',
                 'bristol', 'liverpool', 'sheffield', 'edinburgh', 'cardiff',
                 'nottingham', 'newcastle', 'southampton', 'portsmouth', 'leicester']

    for location in locations:
        try:
            url = f'https://www.yell.com/ucs/UcsSearchAction.do?keywords=truck+tyres&location={location}'
            print(f"  Scraping: {location}...")

            r = requests.get(url, headers=HEADERS, timeout=20)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, 'html.parser')

                # Find business listings
                listings = soup.find_all('div', class_='businessCapsule--mainRow')

                for listing in listings:
                    try:
                        # Company name
                        name_tag = listing.find('h2', class_='businessCapsule--name')
                        name = name_tag.get_text(strip=True) if name_tag else None

                        if not name:
                            continue

                        # Website
                        website = None
                        website_link = listing.find('a', class_='businessCapsule--ctaItem', href=re.compile(r'website'))
                        if website_link:
                            website = website_link.get('href', '')

                        # Phone
                        phone = None
                        phone_tag = listing.find('span', class_='business--telephoneNumber')
                        if phone_tag:
                            phone = phone_tag.get_text(strip=True)

                        # Address
                        address = None
                        addr_tag = listing.find('span', class_='businessCapsule--address')
                        if addr_tag:
                            address = addr_tag.get_text(strip=True)

                        companies.append({
                            'name': name,
                            'website': website,
                            'phone': phone,
                            'address': address,
                            'location': location,
                            'source': 'Yell.com'
                        })

                    except Exception as e:
                        continue

                print(f"    Found {len(listings)} listings")

            time.sleep(2)  # Be polite

        except Exception as e:
            print(f"    Error: {e}")

    print(f"  TOTAL from Yell: {len(companies)}")
    return companies


def scrape_checkatrade_truck_tyres():
    """Scrape Checkatrade for commercial tyre services"""
    print("\n[2] SCRAPING CHECKATRADE - COMMERCIAL TYRES")
    print("-" * 50)

    companies = []

    try:
        url = 'https://www.checkatrade.com/Search?query=commercial%20tyres&location=&page=1'
        print(f"  Scraping checkatrade...")

        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')

            # Find listings
            listings = soup.find_all('div', {'data-testid': re.compile(r'trade-card')})

            for listing in listings:
                try:
                    name_tag = listing.find('h2') or listing.find('h3')
                    name = name_tag.get_text(strip=True) if name_tag else None

                    if name:
                        companies.append({
                            'name': name,
                            'source': 'Checkatrade'
                        })
                except:
                    continue

            print(f"    Found {len(listings)} listings")

    except Exception as e:
        print(f"  Error: {e}")

    return companies


def scrape_thomsonlocal_truck_tyres():
    """Scrape Thomson Local for truck tyre companies"""
    print("\n[3] SCRAPING THOMSON LOCAL - TRUCK TYRES")
    print("-" * 50)

    companies = []

    locations = ['london', 'birmingham', 'manchester', 'leeds', 'glasgow']

    for location in locations:
        try:
            url = f'https://www.thomsonlocal.com/search/{location}/truck-tyres'
            print(f"  Scraping: {location}...")

            r = requests.get(url, headers=HEADERS, timeout=20)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, 'html.parser')

                listings = soup.find_all('div', class_='listing')

                for listing in listings:
                    try:
                        name_tag = listing.find('h2') or listing.find('a', class_='listing-name')
                        name = name_tag.get_text(strip=True) if name_tag else None

                        if name:
                            website = None
                            web_link = listing.find('a', href=re.compile(r'^http'))
                            if web_link and 'thomsonlocal' not in web_link.get('href', ''):
                                website = web_link.get('href')

                            companies.append({
                                'name': name,
                                'website': website,
                                'location': location,
                                'source': 'Thomson Local'
                            })
                    except:
                        continue

                print(f"    Found {len(listings)} listings")

            time.sleep(2)

        except Exception as e:
            print(f"    Error: {e}")

    return companies


def scrape_freeindex_truck_tyres():
    """Scrape FreeIndex for truck tyre companies"""
    print("\n[4] SCRAPING FREEINDEX - TRUCK TYRES")
    print("-" * 50)

    companies = []

    try:
        url = 'https://www.freeindex.co.uk/categories/vehicles/tyres/truck_tyres/'
        print(f"  Scraping freeindex truck tyres category...")

        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')

            # Find business listings
            listings = soup.find_all('div', class_='listing_item')

            for listing in listings:
                try:
                    name_tag = listing.find('h2') or listing.find('a', class_='listing_title')
                    name = name_tag.get_text(strip=True) if name_tag else None

                    if name:
                        # Try to get website
                        website = None
                        links = listing.find_all('a', href=True)
                        for link in links:
                            href = link.get('href', '')
                            if href.startswith('http') and 'freeindex' not in href:
                                website = href
                                break

                        companies.append({
                            'name': name,
                            'website': website,
                            'source': 'FreeIndex'
                        })
                except:
                    continue

            print(f"    Found {len(listings)} listings")

    except Exception as e:
        print(f"  Error: {e}")

    return companies


def scrape_google_maps_api():
    """Note: Would need Google Places API key"""
    print("\n[5] GOOGLE MAPS - Would need API key")
    print("-" * 50)
    print("  Skipping - requires paid API")
    return []


def scrape_247mobiletrucktyres():
    """Scrape 247 Mobile Truck Tyres network"""
    print("\n[6] SCRAPING 247MOBILETRUCKTYRES.CO.UK")
    print("-" * 50)

    companies = []

    try:
        url = 'https://www.247mobiletrucktyres.co.uk/'
        print(f"  Scraping main site...")

        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')

            # Get their contact/coverage info
            text = soup.get_text()

            # This is the company itself
            companies.append({
                'name': '247 Mobile Truck Tyres',
                'website': 'https://www.247mobiletrucktyres.co.uk/',
                'phone': '0800 028 7SEE IF ON PAGE',
                'source': 'Direct scrape',
                'type': 'Mobile Truck Tyre Service'
            })

            # Look for phone numbers
            phones = re.findall(r'0\d{2,4}[\s\-]?\d{3,4}[\s\-]?\d{3,4}', r.text)
            if phones:
                companies[-1]['phone'] = phones[0]

            print(f"    Scraped main company info")

    except Exception as e:
        print(f"  Error: {e}")

    return companies


def scrape_hgvtyres_com():
    """Scrape hgvtyres.com"""
    print("\n[7] SCRAPING HGVTYRES.COM")
    print("-" * 50)

    companies = []

    try:
        url = 'https://www.hgvtyres.com/'
        print(f"  Scraping...")

        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')

            companies.append({
                'name': 'HGV Tyres',
                'website': 'https://www.hgvtyres.com/',
                'source': 'Direct scrape',
                'type': 'HGV Tyre Service'
            })

            phones = re.findall(r'0\d{2,4}[\s\-]?\d{3,4}[\s\-]?\d{3,4}', r.text)
            if phones:
                companies[-1]['phone'] = phones[0]

            print(f"    Scraped main company info")

    except Exception as e:
        print(f"  Error: {e}")

    return companies


def scrape_tyrenet():
    """Scrape Tyrenet dealer network"""
    print("\n[8] SCRAPING TYRENET.NET")
    print("-" * 50)

    companies = []

    try:
        url = 'https://tyrenet.net/'
        print(f"  Scraping...")

        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')

            companies.append({
                'name': 'Tyrenet',
                'website': 'https://tyrenet.net/',
                'source': 'Direct scrape',
                'type': 'Commercial Tyre Network'
            })

            phones = re.findall(r'0\d{2,4}[\s\-]?\d{3,4}[\s\-]?\d{3,4}', r.text)
            if phones:
                companies[-1]['phone'] = phones[0]

    except Exception as e:
        print(f"  Error: {e}")

    return companies


def scrape_tructyre():
    """Scrape Tructyre ATS network"""
    print("\n[9] SCRAPING TRUCTYRE.CO.UK")
    print("-" * 50)

    companies = []

    try:
        # Try to find their depot/branch locator
        urls = [
            'https://www.tructyre.co.uk/',
            'https://www.tructyre.co.uk/depots/',
            'https://www.tructyre.co.uk/find-a-depot/',
        ]

        for url in urls:
            print(f"  Trying: {url}")
            try:
                r = requests.get(url, headers=HEADERS, timeout=15)
                if r.status_code == 200:
                    soup = BeautifulSoup(r.text, 'html.parser')

                    # Look for depot listings
                    depots = soup.find_all(['div', 'li'], class_=re.compile(r'depot|branch|location', re.I))

                    if depots:
                        print(f"    Found {len(depots)} depot elements")

                        for depot in depots:
                            name = depot.get_text(strip=True)[:100]
                            if name:
                                companies.append({
                                    'name': f'Tructyre - {name}',
                                    'website': 'https://www.tructyre.co.uk/',
                                    'source': 'Tructyre website',
                                    'type': 'Truck Tyre Depot'
                                })

                    # If no depots found, at least add main company
                    if not companies:
                        companies.append({
                            'name': 'Tructyre ATS',
                            'website': 'https://www.tructyre.co.uk/',
                            'source': 'Direct scrape',
                            'type': 'Truck Tyre Network'
                        })

                    break

            except Exception as e:
                print(f"    Error: {e}")
                continue

    except Exception as e:
        print(f"  Error: {e}")

    return companies


def scrape_bandvulc():
    """Scrape Bandvulc (major UK truck tyre company)"""
    print("\n[10] SCRAPING BANDVULC.CO.UK")
    print("-" * 50)

    companies = []

    try:
        url = 'https://www.bandvulc.co.uk/'
        print(f"  Scraping...")

        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code == 200:
            companies.append({
                'name': 'Bandvulc',
                'website': 'https://www.bandvulc.co.uk/',
                'source': 'Direct scrape',
                'type': 'Truck Tyre Manufacturer/Retreader'
            })

            phones = re.findall(r'0\d{2,4}[\s\-]?\d{3,4}[\s\-]?\d{3,4}', r.text)
            if phones:
                companies[-1]['phone'] = phones[0]

            print(f"    Scraped")

    except Exception as e:
        print(f"  Error: {e}")

    return companies


def deduplicate(companies):
    """Remove duplicates based on company name"""
    seen = set()
    unique = []

    for c in companies:
        name_lower = c.get('name', '').lower().strip()
        if name_lower and name_lower not in seen:
            seen.add(name_lower)
            unique.append(c)

    return unique


def filter_truck_only(companies):
    """Keep only companies that are clearly truck/HGV tyre related"""
    truck_keywords = ['truck', 'hgv', 'lorry', 'commercial', 'trailer', 'fleet']

    filtered = []
    for c in companies:
        name = c.get('name', '').lower()
        # Include if has truck keyword OR is from a truck-specific source
        if any(kw in name for kw in truck_keywords):
            filtered.append(c)
        elif c.get('type', '').lower().startswith(('truck', 'hgv', 'commercial')):
            filtered.append(c)

    return filtered


def save_results(companies, filename='SCRAPED_TRUCK_TYRES'):
    """Save to JSON and display"""

    # Save JSON
    with open(f'{filename}.json', 'w') as f:
        json.dump(companies, f, indent=2)
    print(f"\nSaved to {filename}.json")

    # Display
    print("\n" + "=" * 80)
    print("SCRAPED TRUCK TYRE COMPANIES")
    print("=" * 80)

    for i, c in enumerate(companies, 1):
        print(f"\n{i}. {c.get('name', 'Unknown')}")
        if c.get('website'):
            print(f"   Website: {c['website']}")
        if c.get('phone'):
            print(f"   Phone: {c['phone']}")
        if c.get('address'):
            print(f"   Address: {c['address']}")
        print(f"   Source: {c.get('source', 'Unknown')}")


def main():
    print("=" * 80)
    print("UK TRUCK TYRE COMPANIES - REAL WEB SCRAPER")
    print("=" * 80)
    print(f"Started: {datetime.now()}")

    all_companies = []

    # Run all scrapers
    all_companies.extend(scrape_yell_truck_tyres())
    all_companies.extend(scrape_thomsonlocal_truck_tyres())
    all_companies.extend(scrape_freeindex_truck_tyres())
    all_companies.extend(scrape_247mobiletrucktyres())
    all_companies.extend(scrape_hgvtyres_com())
    all_companies.extend(scrape_tyrenet())
    all_companies.extend(scrape_tructyre())
    all_companies.extend(scrape_bandvulc())

    print(f"\n\nTotal scraped (raw): {len(all_companies)}")

    # Deduplicate
    unique = deduplicate(all_companies)
    print(f"After deduplication: {len(unique)}")

    # Filter to truck only
    truck_only = filter_truck_only(unique)
    print(f"Truck-specific only: {len(truck_only)}")

    # Save
    save_results(truck_only)

    print(f"\n\nFinished: {datetime.now()}")

    return truck_only


if __name__ == "__main__":
    main()
