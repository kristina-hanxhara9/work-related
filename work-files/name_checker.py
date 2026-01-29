import pandas as pd
import sys

try:
    from rapidfuzz import fuzz
    USE_FUZZY = True
except ImportError:
    print("WARNING: Install 'rapidfuzz': python -m pip install rapidfuzz")
    USE_FUZZY = False

def generate_search_variants(name):
    """
    Generate search variants from full name to shorter parts.
    Minimum 5 characters for any variant to avoid stupid matches like "Les"
    """
    name_lower = name.lower().strip()
    variants = []

    # Only add if long enough
    if len(name_lower) >= 5:
        variants.append(name_lower)

    # Clean version without special chars
    clean_name = name_lower.replace('+', ' ').replace('&', ' ').replace('-', ' ').replace(',', ' ')
    clean_name = ' '.join(clean_name.split())
    if clean_name != name_lower and len(clean_name) >= 5:
        variants.append(clean_name)

    # Split into words
    words = [w for w in clean_name.split() if len(w) >= 4]  # Only words with 4+ chars

    if len(words) >= 2:
        # First two words together
        combo = ' '.join(words[:2])
        if len(combo) >= 5 and combo not in variants:
            variants.append(combo)

        # First word + last word
        combo = words[0] + ' ' + words[-1]
        if len(combo) >= 5 and combo not in variants:
            variants.append(combo)

    # Single word only if it's long enough (6+ chars)
    if len(words) >= 1 and len(words[0]) >= 6:
        if words[0] not in variants:
            variants.append(words[0])

    # If no variants, use original if it's at least 4 chars
    if not variants and len(name_lower) >= 4:
        variants.append(name_lower)

    return variants

def check_names_in_excel(names_file, data_file, threshold=75):
    """
    Check names using cascading search.
    Requires minimum 5 character matches to avoid false positives.
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

        variants = generate_search_variants(name)

        if not variants:
            print(f"SKIPPED (too short): '{name}'")
            continue

        found_match = False
        match_info = None

        for variant in variants:
            if found_match:
                break

            for col in existing_columns:
                if found_match:
                    break

                col_lower = f'_{col}_lower'

                for idx, row in data_df.iterrows():
                    cell_value = row.get(col_lower, '')
                    if not cell_value or len(cell_value) < 4:
                        continue

                    score = 0

                    # Exact match
                    if variant == cell_value:
                        score = 100
                    # Contains - but only if variant is substantial part of cell
                    elif variant in cell_value and len(variant) >= len(cell_value) * 0.5:
                        score = 95
                    elif cell_value in variant and len(cell_value) >= len(variant) * 0.5:
                        score = 90
                    # Fuzzy match
                    elif USE_FUZZY:
                        score = max(
                            fuzz.ratio(variant, cell_value),
                            fuzz.token_sort_ratio(variant, cell_value),
                            fuzz.token_set_ratio(variant, cell_value)
                        )

                    if score >= threshold:
                        found_match = True
                        match_info = {
                            'Search Name': name,
                            'Matched Variant': variant,
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
    threshold = 75

    if len(sys.argv) >= 3:
        names_file = sys.argv[1]
        data_file = sys.argv[2]
    if len(sys.argv) >= 4:
        threshold = int(sys.argv[3])

    print(f"""
╔══════════════════════════════════════════════════════════════╗
║      NAME CHECKER - Smart Search                             ║
╠══════════════════════════════════════════════════════════════╣
║  Columns: RETAILERNAME, COMMENT, COMPANY, NAME 1, NAME 2     ║
║  Minimum 5 chars to match (no stupid short matches)          ║
╚══════════════════════════════════════════════════════════════╝
    """)

    check_names_in_excel(names_file, data_file, threshold)
