import pandas as pd
import sys

try:
    from rapidfuzz import fuzz
    USE_FUZZY = True
except ImportError:
    print("Note: Install 'rapidfuzz' for fuzzy matching: pip install rapidfuzz")
    USE_FUZZY = False

def check_names_in_excel(names_file, data_file, threshold=60):
    """
    Check names from a single-column Excel file against specific columns in a larger Excel file.
    Uses fuzzy matching to find similar names in ALL specified columns.
    """

    # Read the names file (single column)
    print(f"Reading names from: {names_file}")
    names_df = pd.read_excel(names_file, header=None)
    names_to_check = names_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    print(f"Found {len(names_to_check)} names to check")
    print(f"Fuzzy threshold: {threshold}%\n")

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

    # Store all results
    all_matches = []

    # Create combined search text for each row (check all columns at once)
    print("Preparing data...")

    # Pre-process: create lowercase versions and combined text
    for col in existing_columns:
        data_df[f'_{col}_lower'] = data_df[col].fillna('').astype(str).str.lower().str.strip()

    total_names = len(names_to_check)

    # Check each name
    for i, name in enumerate(names_to_check):
        if (i + 1) % 10 == 0:
            print(f"Processing {i + 1}/{total_names}...")

        name_lower = name.lower().strip()
        name_matches = []
        matched_gsnrs = set()  # Track which GSNRs we've already matched for this name

        # Check each column
        for col in existing_columns:
            col_lower = f'_{col}_lower'

            # First: exact/contains match (fast)
            mask = data_df[col_lower].str.contains(name_lower, case=False, na=False, regex=False)
            exact_matches = data_df[mask]

            for idx, row in exact_matches.iterrows():
                gsnr = row.get('GSNR', 'N/A')
                if gsnr not in matched_gsnrs:
                    matched_gsnrs.add(gsnr)
                    match_info = {
                        'Search Name': name,
                        'Found In Column': col,
                        'Matched Value': row.get(col, ''),
                        'Similarity': '100%',
                        'GSNR': gsnr,
                        'RETAILERNAME': row.get('RETAILERNAME', 'N/A'),
                        'COMMENT': row.get('COMMENT', 'N/A'),
                        'COMPANY': row.get('COMPANY', 'N/A'),
                        'NAME 1': row.get('NAME 1', 'N/A'),
                        'NAME 2': row.get('NAME 2', 'N/A'),
                        'STATUS': row.get('STATUS', 'N/A')
                    }
                    name_matches.append(match_info)

            # Second: fuzzy match (only if rapidfuzz is available)
            if USE_FUZZY:
                non_matched = data_df[~data_df['GSNR'].isin(matched_gsnrs)]

                for idx, row in non_matched.iterrows():
                    cell_value = row.get(col_lower, '')
                    if not cell_value or len(cell_value) < 2:
                        continue

                    # Quick fuzzy check
                    score = fuzz.partial_ratio(name_lower, cell_value)

                    if score >= threshold:
                        gsnr = row.get('GSNR', 'N/A')
                        if gsnr not in matched_gsnrs:
                            matched_gsnrs.add(gsnr)
                            match_info = {
                                'Search Name': name,
                                'Found In Column': col,
                                'Matched Value': row.get(col, ''),
                                'Similarity': f'{score}%',
                                'GSNR': gsnr,
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
            print(f"'{name}': {len(name_matches)} matches")
            for match in name_matches:
                val = str(match['Matched Value'])[:35]
                print(f"  -> GSNR: {match['GSNR']} | {match['Found In Column']}: {val}... ({match['Similarity']})")

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
    threshold = 60

    if len(sys.argv) >= 3:
        names_file = sys.argv[1]
        data_file = sys.argv[2]
    if len(sys.argv) >= 4:
        threshold = int(sys.argv[3])

    print(f"""
╔══════════════════════════════════════════════════════════════╗
║      NAME CHECKER - Fuzzy Search Tool                        ║
╠══════════════════════════════════════════════════════════════╣
║  Checking: RETAILERNAME, COMMENT, COMPANY, NAME 1, NAME 2    ║
║  Fuzzy matching enabled (finds similar names)                ║
╚══════════════════════════════════════════════════════════════╝
    """)

    check_names_in_excel(names_file, data_file, threshold)
