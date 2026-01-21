# UK Truck Tyre Companies Scraper (Python)

Scrapes UK truck tyre wholesalers and fitters from Companies House API with detailed company data.

## Setup

```bash
pip install requests openpyxl
```

## Available Scrapers

### 1. Full Scraper (full_scraper.py) - RECOMMENDED
**Processes ALL 846 companies** with detailed Companies House data and explicit source citations.

```bash
python full_scraper.py
```

**What it does:**
- Loads all 846 companies from scraper.py
- Fetches detailed data from Companies House API for every company with a company number (349 companies)
- Adds web research data with explicit source citations
- Takes ~15-20 minutes due to API rate limits

**Output:**
- `UK_TRUCK_TYRE_FULL_DATABASE.xlsx` - Excel with 4 sheets including source citations
- `UK_TRUCK_TYRE_FULL_DATABASE.json` - JSON format
- `UK_TRUCK_TYRE_FULL_DATABASE.csv` - CSV format

### 2. Basic Scraper (scraper.py)
Contains 846 pre-loaded companies + searches Companies House API for more.

```bash
python scraper.py
```

**Output:**
- `UK_TRUCK_TYRE_COMPANIES.xlsx` - Excel file with 3 sheets
- `UK_TRUCK_TYRE_COMPANIES.csv` - CSV format
- `UK_TRUCK_TYRE_COMPANIES.json` - JSON format

### 3. Other Scrapers
- `detailed_scraper.py` - Searches API and fetches detailed data
- `master_scraper.py` - Combines sources (but doesn't process all 846)
- `generate_research_report.py` - Creates research report for 40 major companies

## Data Sources (with citations)

### Primary Source: Companies House API
- **URL:** `https://api.company-information.service.gov.uk`
- **Data:** Company name, number, status, address, directors, PSC, charges, filings
- **Accuracy:** 100% - Official UK government data

### Secondary Sources: Web Research
All research data includes explicit source citations:

| Source | URL | Data Provided |
|--------|-----|---------------|
| UK GlobalDatabase | uk.globaldatabase.com | Revenue, employee counts |
| ZoomInfo | zoominfo.com | Company profiles |
| Growjo | growjo.com | Revenue estimates |
| Owler | owler.com | Company profiles |
| LinkedIn | linkedin.com/company/* | Employee counts |
| Tyrepress | tyrepress.com | Industry news |
| Commercial Tyre Business | commercialtyrebusiness.com | Industry awards |
| Company Websites | Various | Official company data |

## Data Included

### From Companies House API:
- Company name, number, status, type
- Registered address
- SIC codes (industry classification)
- Date of creation
- Directors/Officers names
- Persons with significant control (owners with 25%+ shares)
- Charges (loans/mortgages)
- Filing history
- Last accounts date
- Insolvency history

### From Web Research (20 major companies):
- Revenue estimates (with source)
- Employee counts (with source)
- Number of branches/depots (with source)
- Business descriptions (with source)
- Websites

## Company Types Covered

- B2B wholesalers
- Truck tyre fitters
- Retreaders
- Fleet service providers
- Manufacturers (UK operations)
- Commercial tyre specialists

## API Key

Uses Companies House API: `48d17266-ff2e-425f-9b20-7dcc9b25bb79`

Rate limited to 0.6s between requests (600 requests per 5 minutes).

## Key Market Insights

| Company | Revenue | Source |
|---------|---------|--------|
| Micheldever Group | £575M | UK GlobalDatabase |
| Halfords Commercial | £384M B2B | Halfords Annual Report |
| Kirkby Tyres | £60.4M | UK GlobalDatabase |
| Stapleton's | £200M+ | Insider Media |
| ATS Euromaster | $346-450M | Growjo |
| Bond International | $150M+ | Fast Track 100 |
| Tanvic Group | £70M | Company Website |
| Lodge Tyre | $64.6M | Growjo |
