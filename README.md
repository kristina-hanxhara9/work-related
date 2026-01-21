# UK Truck Tyre Companies Scraper (Python)

Scrapes UK truck tyre wholesalers and fitters from Companies House API.

## Setup

```bash
pip install requests openpyxl
```

## Run the Scraper

```bash
python scraper.py
```

## Output Files

- `UK_TRUCK_TYRE_COMPANIES.xlsx` - Excel file with 3 sheets
- `UK_TRUCK_TYRE_COMPANIES.csv` - CSV format
- `UK_TRUCK_TYRE_COMPANIES.json` - JSON format

## Data Included

- UK truck tyre companies from Companies House
- Company name, number, address, business type
- B2B wholesalers, fitters, specialists, retreaders
