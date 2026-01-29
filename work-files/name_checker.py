import pandas as pd
import sys

try:
    from rapidfuzz import fuzz
    USE_FUZZY = True
except ImportError:
    print("WARNING: Install 'rapidfuzz': python -m pip install rapidfuzz")
    USE_FUZZY = False

def check_names_in_excel(names_file, data_file, threshold=80):
    """
    Check names - searches for the FULL NAME only.
    No breaking into parts. Strict matching.
    """

    print(f"Reading names from: {names_file}")
    names_df = pd.read_excel(names_file, header=None)
    names_to_check = names_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    print(f"Found {len(names_to_check)} names to check")
    print(f"Similarity threshold: {threshold}%\n")

    print(f"Reading data from: {data_file}")
    data_df = pd.read_excel(data_file)
    print(f"Data file has {len(data_df)} rows\n")

    search_columns = ['RETAILERNAME', 'COMMENT', 'COMPANY', 'NAME 1', 'NAME 2']
    existing_columns = [col for col in search_columns if col in data_df.columns]

    if len(existing_columns) < len(search_columns):
        missing = [col for col in search_columns if col not in data_df.columns]
        print(f"Warning: Columns not found: {missing}\n")

    print(f"Searching in: {existing_columns}\n")
    print("=" * 100)

    for col in existing_columns:
        data_df[f'_{col}_lower'] = data_df[col].fillna('').astype(str).str.lower().str.strip()

    all_matches = []
    total_names = len(names_to_check)

    for i, name in enumerate(names_to_check):
        if (i + 1) % 5 == 0:
            print(f"Processing {i + 1}/{total_names}...")

        name_lower = name.lower().strip()

        # Clean the name (remove special chars)
        clean_name = name_lower.replace('+', ' ').replace('&', ' ').replace('-', ' ').replace(',', ' ')
        clean_name = ' '.join(clean_name.split())

        found_match = False
        match_info = None
        best_score = 0

        for col in existing_columns:
            if found_match:
                break

            col_lower = f'_{col}_lower'

            for idx, row in data_df.iterrows():
                cell_value = row.get(col_lower, '')
                if not cell_value:
                    continue

                score = 0

                # Method 1: Exact match
                if clean_name == cell_value:
                    score = 100

                # Method 2: Full name contained in cell (or vice versa)
                elif clean_name in cell_value:
                    # Score based on how much of cell is the name
                    score = int(80 * len(clean_name) / len(cell_value))
                elif cell_value in clean_name:
                    score = int(80 * len(cell_value) / len(clean_name))

                # Method 3: Fuzzy match using ONLY ratio (strict character comparison)
                elif USE_FUZZY:
                    # Only use fuzz.ratio - it's the strictest
                    score = fuzz.ratio(clean_name, cell_value)

                # Only accept if score meets threshold AND is better than previous
                if score >= threshold and score > best_score:
                    best_score = score
                    found_match = True
                    match_info = {
                        'Search Name': name,
                        'Found In Column': col,
                        'Matched Value': row.get(col, ''),
                        'Similarity': f'{score}%',
                        'GSNR': row.get('GSNR', 'N/A'),
                        'RETAILERNAME': row.get('RETAILERNAME', 'N/A'),
                        'COMMENT': row.get('COMMENT', 'N/A'),
                        'COMPANY': row.get('COMPANY', 'N/A'),
                        'NAME 1': row.get('NAME 1', 'N/A'),
                        'NAME 2': row.get('NAME 2', 'N/A'),
                        'STATUS': row.get('STATUS', 'N/A')
                    }

                    # If exact match, stop searching
                    if score == 100:
                        break

        if found_match and match_info:
            all_matches.append(match_info)
            print(f"FOUND: '{name}' -> {match_info['Matched Value'][:40]}... ({match_info['Similarity']})")
        else:
            print(f"NOT FOUND! '{name}'")

    if all_matches:
        results_df = pd.DataFrame(all_matches)
        output_file = 'name_check_results.xlsx'
        results_df.to_excel(output_file, index=False)

        print(f"\n{'='*100}")
        print(f"SUMMARY")
        print(f"{'='*100}")
        print(f"Names checked: {len(names_to_check)}")
        print(f"Matches found: {len(results_df)}")
        print(f"Not found: {len(names_to_check) - len(results_df)}")
        print(f"Results saved to: {output_file}")
    else:
        print(f"\n\nNo matches found.")

    return all_matches


if __name__ == "__main__":
    names_file = "Calculus-list.xlsx"
    data_file = "MDM.xlsx"
    threshold = 80

    if len(sys.argv) >= 3:
        names_file = sys.argv[1]
        data_file = sys.argv[2]
    if len(sys.argv) >= 4:
        threshold = int(sys.argv[3])

    print(f"""
╔══════════════════════════════════════════════════════════════╗
║      NAME CHECKER - Strict Full Name Search                  ║
╠══════════════════════════════════════════════════════════════╣
║  Searches for FULL NAME only - no breaking into parts        ║
║  Columns: RETAILERNAME, COMMENT, COMPANY, NAME 1, NAME 2     ║
╚══════════════════════════════════════════════════════════════╝
    """)

    check_names_in_excel(names_file, data_file, threshold)
