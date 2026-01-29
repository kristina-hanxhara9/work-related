import pandas as pd
import sys

try:
    from rapidfuzz import fuzz
    USE_FUZZY = True
except ImportError:
    print("WARNING: Install 'rapidfuzz' for better matching: python -m pip install rapidfuzz")
    USE_FUZZY = False

def check_names_in_excel(names_file, data_file, threshold=70):
    """
    Check names from a single-column Excel file against specific columns in a larger Excel file.
    Uses smart fuzzy matching - shows accurate confidence scores.
    """

    # Read the names file (single column)
    print(f"Reading names from: {names_file}")
    names_df = pd.read_excel(names_file, header=None)
    names_to_check = names_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    print(f"Found {len(names_to_check)} names to check")
    print(f"Minimum similarity threshold: {threshold}%\n")

    # Read the large data file
    print(f"Reading data from: {data_file}")
    data_df = pd.read_excel(data_file)
    print(f"Data file has {len(data_df)} rows\n")

    # Columns to search in
    search_columns = ['RETAILERNAME', 'COMMENT', 'COMPANY', 'NAME 1', 'NAME 2']

    # Check which columns actually exist
    existing_columns = [col for col in search_columns if col in data_df.columns]
    missing_columns = [col for col in search_columns if col not in data_df.columns]

    if missing_columns:
        print(f"Warning: Columns not found: {missing_columns}")
        print(f"Available columns: {list(data_df.columns)}\n")

    print(f"Searching in: {existing_columns}\n")
    print("=" * 100)

    # Store results
    all_matches = []

    # Pre-process columns
    print("Preparing data...")
    for col in existing_columns:
        data_df[f'_{col}_lower'] = data_df[col].fillna('').astype(str).str.lower().str.strip()

    total_names = len(names_to_check)

    # Check each name
    for i, name in enumerate(names_to_check):
        if (i + 1) % 5 == 0:
            print(f"Processing {i + 1}/{total_names}...")

        name_lower = name.lower().strip()
        name_matches = []

        # Check each row
        for idx, row in data_df.iterrows():
            best_score = 0
            best_col = None
            best_value = None

            # Check each column
            for col in existing_columns:
                col_lower = f'_{col}_lower'
                cell_value = row.get(col_lower, '')

                if not cell_value or len(cell_value) < 2:
                    continue

                # Calculate similarity score
                if USE_FUZZY:
                    # Use the best of different fuzzy methods
                    score = max(
                        fuzz.ratio(name_lower, cell_value),  # Exact similarity
                        fuzz.token_sort_ratio(name_lower, cell_value),  # Word order doesn't matter
                        fuzz.token_set_ratio(name_lower, cell_value)  # Handles extra words
                    )
                else:
                    # Fallback: exact match only
                    if name_lower == cell_value:
                        score = 100
                    elif name_lower in cell_value or cell_value in name_lower:
                        score = 80
                    else:
                        score = 0

                if score > best_score:
                    best_score = score
                    best_col = col
                    best_value = row.get(col, '')

            # Only record if above threshold
            if best_score >= threshold:
                match_info = {
                    'Search Name': name,
                    'Found In Column': best_col,
                    'Matched Value': best_value,
                    'Similarity': f'{best_score}%',
                    'GSNR': row.get('GSNR', 'N/A'),
                    'RETAILERNAME': row.get('RETAILERNAME', 'N/A'),
                    'COMMENT': row.get('COMMENT', 'N/A'),
                    'COMPANY': row.get('COMPANY', 'N/A'),
                    'NAME 1': row.get('NAME 1', 'N/A'),
                    'NAME 2': row.get('NAME 2', 'N/A'),
                    'STATUS': row.get('STATUS', 'N/A')
                }
                name_matches.append(match_info)

        all_matches.extend(name_matches)

        # Print results
        if not name_matches:
            print(f"NOT FOUND! '{name}'")
        else:
            print(f"'{name}': {len(name_matches)} matches found")

    # Save to Excel
    if all_matches:
        results_df = pd.DataFrame(all_matches)
        output_file = 'name_check_results.xlsx'
        results_df.to_excel(output_file, index=False)

        print(f"\n{'='*100}")
        print(f"SUMMARY")
        print(f"{'='*100}")
        print(f"Names checked: {len(names_to_check)}")
        print(f"Matches found: {len(results_df)}")
        print(f"Results saved to: {output_file}")
    else:
        print(f"\n\nNo matches found.")

    return all_matches


if __name__ == "__main__":
    names_file = "Calculus-list.xlsx"
    data_file = "MDM.xlsx"
    threshold = 70  # 70% similarity required

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
║  Shows accurate similarity scores                            ║
╚══════════════════════════════════════════════════════════════╝
    """)

    check_names_in_excel(names_file, data_file, threshold)
