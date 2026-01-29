import pandas as pd
import sys

def check_names_in_excel(names_file, data_file):
    """
    Check names from a single-column Excel file against specific columns in a larger Excel file.
    Uses FAST partial matching (contains).
    Shows all matching rows with their GSNR values.
    """

    # Read the names file (single column)
    print(f"Reading names from: {names_file}")
    names_df = pd.read_excel(names_file, header=None)
    names_to_check = names_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    print(f"Found {len(names_to_check)} names to check\n")

    # Read the large data file
    print(f"Reading data from: {data_file}")
    data_df = pd.read_excel(data_file)
    print(f"Data file has {len(data_df)} rows\n")

    # Columns to search in
    search_columns = ['RETAILERNAME', 'COMMENT', 'COMPANY', 'NAME 1', 'NAME 2']

    # Check which columns actually exist in the data
    existing_columns = [col for col in search_columns if col in data_df.columns]
    missing_columns = [col for col in search_columns if col not in data_df.columns]

    if missing_columns:
        print(f"Warning: These columns were not found: {missing_columns}")
        print(f"Available columns: {list(data_df.columns)}\n")

    print(f"Searching in columns: {existing_columns}\n")
    print("=" * 100)

    # Prepare lowercase columns for fast searching
    for col in existing_columns:
        data_df[f'_{col}_lower'] = data_df[col].fillna('').astype(str).str.lower()

    # Store all results
    all_matches = []

    # Check each name using FAST vectorized pandas operations
    for name in names_to_check:
        name_lower = name.lower().strip()
        name_matches = []

        for col in existing_columns:
            col_lower = f'_{col}_lower'

            # Fast partial match using pandas str.contains
            mask = data_df[col_lower].str.contains(name_lower, case=False, na=False, regex=False)
            matches = data_df[mask]

            for idx, row in matches.iterrows():
                match_info = {
                    'Search Name': name,
                    'Found In Column': col,
                    'Matched Value': row.get(col, ''),
                    'GSNR': row.get('GSNR', 'N/A'),
                    'RETAILERNAME': row.get('RETAILERNAME', 'N/A'),
                    'COMMENT': row.get('COMMENT', 'N/A'),
                    'COMPANY': row.get('COMPANY', 'N/A'),
                    'NAME 1': row.get('NAME 1', 'N/A'),
                    'NAME 2': row.get('NAME 2', 'N/A'),
                    'STATUS': row.get('STATUS', 'N/A')
                }
                name_matches.append(match_info)

        # Add matches to results
        all_matches.extend(name_matches)

        # Print results for this name
        if not name_matches:
            print(f"No match: '{name}'")
        else:
            print(f"'{name}': {len(name_matches)} matches found")
            for match in name_matches:
                val = str(match['Matched Value'])[:40]
                print(f"  -> GSNR: {match['GSNR']} | {match['Found In Column']}: {val}...")

    # Save results to Excel
    if all_matches:
        results_df = pd.DataFrame(all_matches)
        output_file = 'name_check_results.xlsx'
        results_df.to_excel(output_file, index=False)

        print(f"\n{'='*100}")
        print(f"SUMMARY")
        print(f"{'='*100}")
        print(f"Total names checked: {len(names_to_check)}")
        print(f"Total matches found: {len(results_df)}")
        print(f"Results saved to: {output_file}")
    else:
        print(f"\n\nNo matches found for any of the {len(names_to_check)} names.")

    return all_matches


if __name__ == "__main__":
    names_file = "Calculus-list.xlsx"
    data_file = "MDM.xlsx"

    if len(sys.argv) >= 3:
        names_file = sys.argv[1]
        data_file = sys.argv[2]

    print(f"""
╔══════════════════════════════════════════════════════════════╗
║      NAME CHECKER - Fast Search Tool                         ║
╠══════════════════════════════════════════════════════════════╣
║  Checking: RETAILERNAME, COMMENT, COMPANY, NAME 1, NAME 2    ║
╚══════════════════════════════════════════════════════════════╝
    """)

    check_names_in_excel(names_file, data_file)
