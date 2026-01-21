# UK Truck Tyre Companies Scraper

Scrapes UK truck tyre wholesalers and fitters from Companies House API.

## Setup

```bash
npm install
```

## Run the Scraper

```bash
# Step 1: Scrape Companies House API
node companies-house-scraper.js

# Step 2: Combine all data
node combine-all.js

# Step 3: Convert to Excel
node convert-to-excel.js
```

## Output Files

- `UK_TRUCK_TYRE_COMPANIES.xlsx` - Excel file with all data (4 sheets)
- `MASTER_UK_TRUCK_TYRE_COMPANIES.csv` - Combined CSV
- `MASTER_UK_TRUCK_TYRE_COMPANIES.json` - Combined JSON

## Data Included

- 846 UK truck tyre companies
- 57 B2B wholesalers
- 65 Companies House verified
- Company name, website, phone, address, business type, service points, region
