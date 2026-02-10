import pandas as pd
import sys

try:
    from rapidfuzz import fuzz
    USE_FUZZY = True
except ImportError:
    print("WARNING: Install 'rapidfuzz': python -m pip install rapidfuzz")
    USE_FUZZY = False


def get_key_word(name):
    """
    Extract the KEY WORD from a name - the most important word to search for.
    This word MUST be present in any match.
    """
    # Clean the name
    clean = name.lower().replace('+', ' ').replace('&', ' ').replace('-', ' ').replace(',', ' ')
    clean = ' '.join(clean.split())

    # Get words with 4+ characters (skip short words like "the", "and", "ltd")
    words = [w for w in clean.split() if len(w) >= 4]

    if words:
        # Return the first significant word
        return words[0]

    # Fallback: return the whole name if no long words
    return clean if len(clean) >= 3 else None


def check_names_in_excel(names_file, data_file, threshold=70):
    """
    Smart name matching:
    1. Extract KEY WORD from search name
    2. Only consider rows where KEY WORD exists in any column
    3. Use fuzzy matching to rank and find best match
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

    # Pre-process: create lowercase versions
    for col in existing_columns:
        data_df[f'_{col}_lower'] = data_df[col].fillna('').astype(str).str.lower().str.strip()

    all_matches = []
    total_names = len(names_to_check)

    for i, name in enumerate(names_to_check):
        if (i + 1) % 5 == 0:
            print(f"Processing {i + 1}/{total_names}...")

        # Get the key word that MUST be present
        key_word = get_key_word(name)

        if not key_word:
            print(f"SKIPPED (too short): '{name}'")
            continue

        # Clean full name for fuzzy comparison
        clean_name = name.lower().replace('+', ' ').replace('&', ' ').replace('-', ' ').replace(',', ' ')
        clean_name = ' '.join(clean_name.split())

        best_match = None
        best_score = 0

        # Step 1: Find all rows where KEY WORD exists in ANY column
        for idx, row in data_df.iterrows():
            key_word_found = False
            found_in_col = None
            found_value = None

            # Check if key word exists in any column
            for col in existing_columns:
                col_lower = f'_{col}_lower'
                cell_value = row.get(col_lower, '')

                if not cell_value:
                    continue

                # KEY WORD must be present (as whole word or part of word)
                if key_word in cell_value:
                    key_word_found = True
                    found_in_col = col
                    found_value = cell_value
                    break

            if not key_word_found:
                continue

            # Step 2: Key word found - now calculate fuzzy score for full name
            if USE_FUZZY:
                score = fuzz.ratio(clean_name, found_value)
            else:
                # Fallback: simple comparison
                if clean_name == found_value:
                    score = 100
                elif clean_name in found_value or found_value in clean_name:
                    score = 80
                else:
                    score = 50  # Key word matched but names differ

            # Keep track of best match
            if score >= threshold and score > best_score:
                best_score = score
                best_match = {
                    'Search Name': name,
                    'Key Word': key_word,
                    'Found In Column': found_in_col,
                    'Matched Value': row.get(found_in_col, ''),
                    'Similarity': f'{score}%',
                    'GSNR': row.get('GSNR', 'N/A'),
                    'RETAILERNAME': row.get('RETAILERNAME', 'N/A'),
                    'COMMENT': row.get('COMMENT', 'N/A'),
                    'COMPANY': row.get('COMPANY', 'N/A'),
                    'NAME 1': row.get('NAME 1', 'N/A'),
                    'NAME 2': row.get('NAME 2', 'N/A'),
                    'STATUS': row.get('STATUS', 'N/A')
                }

        if best_match:
            all_matches.append(best_match)
            print(f"FOUND: '{name}' -> {best_match['Matched Value'][:40]}... ({best_match['Similarity']})")
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
    threshold = 70

    if len(sys.argv) >= 3:
        names_file = sys.argv[1]
        data_file = sys.argv[2]
    if len(sys.argv) >= 4:
        threshold = int(sys.argv[3])

    print(f"""
╔══════════════════════════════════════════════════════════════╗
║      NAME CHECKER - Smart Key Word Matching                  ║
╠══════════════════════════════════════════════════════════════╣
║  Step 1: Find KEY WORD (first 4+ char word) in columns       ║
║  Step 2: Fuzzy match full name to rank results               ║
║  Columns: RETAILERNAME, COMMENT, COMPANY, NAME 1, NAME 2     ║
╚══════════════════════════════════════════════════════════════╝
    """)

    check_names_in_excel(names_file, data_file, threshold)
