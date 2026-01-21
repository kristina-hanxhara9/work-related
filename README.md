# UK Truck Tyre Companies Scraper (Python)

Scrapes UK truck tyre wholesalers and fitters from Companies House API with detailed company data.

## Setup

```bash
pip install requests openpyxl
```

## Available Scrapers

### 1. Basic Scraper (scraper.py)
Contains 846 pre-loaded companies + searches Companies House API for more.

```bash
python scraper.py
```

**Output:**
- `UK_TRUCK_TYRE_COMPANIES.xlsx` - Excel file with 3 sheets
- `UK_TRUCK_TYRE_COMPANIES.csv` - CSV format
- `UK_TRUCK_TYRE_COMPANIES.json` - JSON format

### 2. Detailed Scraper (detailed_scraper.py)
Fetches comprehensive data from Companies House API:
- Full company profile
- Officers/Directors
- Filing history
- Charges (mortgages/loans)
- Persons with significant control (PSC)

```bash
python detailed_scraper.py
```

**Output:**
- `UK_TRUCK_TYRE_DETAILED_REPORT.xlsx`
- `UK_TRUCK_TYRE_DETAILED_DATA.json`

### 3. Master Scraper (master_scraper.py) - RECOMMENDED
Combines all data sources:
- Industry database (846 companies)
- Companies House API search
- Detailed Companies House data
- Web research data (revenue, employees, descriptions)

```bash
python master_scraper.py
```

**Output:**
- `UK_TRUCK_TYRE_MASTER_DATABASE.xlsx`
- `UK_TRUCK_TYRE_MASTER_DATA.json`
- `UK_TRUCK_TYRE_MASTER_DATA.csv`

### 4. Research Report Generator (generate_research_report.py)
Creates Excel report with 40 major companies and market research.

```bash
python generate_research_report.py
```

**Output:**
- `UK_TRUCK_TYRE_RESEARCH_REPORT.xlsx`

## Data Included

### Companies House API Data:
- Company name, number, status
- Registered address
- SIC codes (industry classification)
- Date of creation
- Directors/Officers names
- Persons with significant control (owners)
- Charges (loans/mortgages)
- Filing history
- Insolvency history

### Research Data (for major companies):
- Revenue estimates
- Employee counts
- Number of branches/depots
- Service van counts
- Business descriptions
- Services offered
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

Rate limited to 0.6s between requests to comply with API terms.

## Key Market Insights

| Company | Revenue | Notable |
|---------|---------|---------|
| Micheldever Group | £575M | 20% UK market share |
| Halfords Commercial | £384M | UK's largest commercial provider |
| Kirkby Tyres | £60.4M | Wholesaler of Year 2024/2025 |
| Stapleton's | £200M+ | Holds 1.5M+ tyres |
| ATS Euromaster | $346-450M | 340 centres, Michelin owned |
| Bond International | $150M+ | Largest independent wholesaler |
