import pandas as pd
import sys

try:
    from rapidfuzz import fuzz
    USE_FUZZY = True
except ImportError:
    print("WARNING: Install 'rapidfuzz': python -m pip install rapidfuzz")
    USE_FUZZY = False


def get_search_attempts(name):
    """
    Generate up to 4 search attempts from a name.
    Example: "J C Campbell Electrics" ->
      1. "j c campbell electrics" (full name)
      2. "campbell" (longest word)
      3. "electrics" (second longest)
      4. First word if different
    """
    clean = name.lower().replace('+', ' ').replace('&', ' ').replace('-', ' ').replace(',', ' ')
    clean = ' '.join(clean.split())

    attempts = []

    # Attempt 1: Full name
    if len(clean) >= 4:
        attempts.append(('Full Name', clean))

    # Get words sorted by length (longest first), minimum 3 chars
    words = [w for w in clean.split() if len(w) >= 3]
    words_by_length = sorted(words, key=len, reverse=True)

    # Attempt 2 & 3: Longest words (must be 4+ chars)
    for word in words_by_length[:2]:
        if len(word) >= 4:
            already_added = [a[1] for a in attempts]
            if word not in already_added:
                attempts.append((f'Word: {word}', word))

    # Attempt 4: First word if not already tried
    if words:
        first_word = words[0]
        already_added = [a[1] for a in attempts]
        if first_word not in already_added and len(first_word) >= 3:
            attempts.append((f'First: {first_word}', first_word))

    return attempts[:4]


def search_in_data(search_term, data_df, existing_columns, threshold=60):
    """
    Search for a term in the data. Returns best match or None.
    """
    best_match = None
    best_score = 0

    for idx, row in data_df.iterrows():
        for col in existing_columns:
            col_lower = f'_{col}_lower'
            cell_value = row.get(col_lower, '')

            if not cell_value:
                continue

            # Check if search term exists in cell
            if search_term in cell_value:
                # Calculate similarity score
                if USE_FUZZY:
                    score = fuzz.ratio(search_term, cell_value)
                else:
                    score = 100 if search_term == cell_value else 70

                if score > best_score:
                    best_score = score
                    best_match = {
                        'col': col,
                        'value': row.get(col, ''),
                        'score': score,
                        'row': row
                    }

    if best_match and best_match['score'] >= threshold:
        return best_match
    return None


def check_names_in_excel(names_file, data_file, threshold=60):
    """
    Multi-attempt name matching with Not Found tracking.
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

    # Pre-process columns
    for col in existing_columns:
        data_df[f'_{col}_lower'] = data_df[col].fillna('').astype(str).str.lower().str.strip()

    all_matches = []
    not_found = []
    total_names = len(names_to_check)

    for i, name in enumerate(names_to_check):
        if (i + 1) % 5 == 0:
            print(f"Processing {i + 1}/{total_names}...")

        # Get search attempts for this name
        attempts = get_search_attempts(name)

        if not attempts:
            not_found.append({
                'Original Name': name,
                'Attempts Tried': 'Name too short',
                'Reason': 'Could not generate search terms'
            })
            print(f"SKIPPED (too short): '{name}'")
            continue

        found = False
        attempts_tried = []

        # Try each attempt
        for attempt_name, search_term in attempts:
            attempts_tried.append(attempt_name)

            match = search_in_data(search_term, data_df, existing_columns, threshold)

            if match:
                found = True
                row = match['row']
                all_matches.append({
                    'Search Name': name,
                    'Search Attempt': attempt_name,
                    'Search Term Used': search_term,
                    'Found In Column': match['col'],
                    'Matched Value': match['value'],
                    'Confidence': f"{match['score']}%",
                    'GSNR': row.get('GSNR', 'N/A'),
                    'RETAILERNAME': row.get('RETAILERNAME', 'N/A'),
                    'COMMENT': row.get('COMMENT', 'N/A'),
                    'COMPANY': row.get('COMPANY', 'N/A'),
                    'NAME 1': row.get('NAME 1', 'N/A'),
                    'NAME 2': row.get('NAME 2', 'N/A'),
                    'STATUS': row.get('STATUS', 'N/A')
                })
                print(f"FOUND: '{name}' -> {match['value'][:40]}... ({match['score']}%) [via {attempt_name}]")
                break

        if not found:
            not_found.append({
                'Original Name': name,
                'Attempts Tried': ', '.join(attempts_tried),
                'Reason': 'No match found after all attempts'
            })
            print(f"NOT FOUND! '{name}' (tried: {', '.join(attempts_tried)})")

    # Save to Excel with two sheets
    output_file = 'MDM_Matches.xlsx'

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        if all_matches:
            pd.DataFrame(all_matches).to_excel(writer, sheet_name='Matches', index=False)
        if not_found:
            pd.DataFrame(not_found).to_excel(writer, sheet_name='Not Found', index=False)

    print(f"\n{'='*100}")
    print(f"SUMMARY")
    print(f"{'='*100}")
    print(f"Names checked: {len(names_to_check)}")
    print(f"Matches found: {len(all_matches)}")
    print(f"Not found: {len(not_found)}")
    print(f"\nResults saved to: {output_file}")
    print(f"  - Sheet 'Matches': {len(all_matches)} items")
    print(f"  - Sheet 'Not Found': {len(not_found)} items")

    return all_matches, not_found


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
║      NAME CHECKER - Multi-Attempt Search                     ║
╠══════════════════════════════════════════════════════════════╣
║  Up to 4 attempts per name (full name, longest words, etc)   ║
║  Columns: RETAILERNAME, COMMENT, COMPANY, NAME 1, NAME 2     ║
║  Output: MDM_Matches.xlsx (Matches + Not Found sheets)       ║
╚══════════════════════════════════════════════════════════════╝
    """)

    check_names_in_excel(names_file, data_file, threshold)
