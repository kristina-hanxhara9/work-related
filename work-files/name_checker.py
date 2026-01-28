import pandas as pd
import sys
from difflib import SequenceMatcher

def similarity_score(a, b):
    """Calculate similarity ratio between two strings (0 to 1)"""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def get_word_matches(search_name, text, threshold=0.6):
    """
    Check if any word in the text is similar to the search name.
    Returns the best similarity score found.
    """
    if not text or not search_name:
        return 0

    text = str(text).lower()
    search_name = str(search_name).lower().strip()

    # Direct containment check (partial match)
    if search_name in text:
        return 1.0

    # Check similarity with each word in the text
    words = text.replace(',', ' ').replace('.', ' ').replace('-', ' ').split()
    best_score = 0

    for word in words:
        word = word.strip()
        if len(word) < 2:
            continue
        score = similarity_score(search_name, word)
        if score > best_score:
            best_score = score

    # Also check similarity with the full text
    full_score = similarity_score(search_name, text)
    if full_score > best_score:
        best_score = full_score

    return best_score

def check_names_in_excel(names_file, data_file, similarity_threshold=0.6):
    """
    Check names from a single-column Excel file against specific columns in a larger Excel file.
    Uses fuzzy matching to find similar names.
    Shows all matching rows with their GSNR values.
    """

    # Read the names file (single column)
    print(f"Reading names from: {names_file}")
    names_df = pd.read_excel(names_file, header=None)
    names_to_check = names_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
    print(f"Found {len(names_to_check)} names to check")
    print(f"Similarity threshold: {similarity_threshold * 100}%\n")

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
        print(f"Warning: These columns were not found in the data file: {missing_columns}")
        print(f"Available columns: {list(data_df.columns)}\n")

    print(f"Searching in columns: {existing_columns}\n")
    print("=" * 100)

    # Store all results
    all_matches = []

    # Check each name
    for name in names_to_check:
        matches_found = False

        for idx, row in data_df.iterrows():
            for col in existing_columns:
                cell_value = row.get(col, '')
                score = get_word_matches(name, cell_value, similarity_threshold)

                if score >= similarity_threshold:
                    matches_found = True
                    match_info = {
                        'Search Name': name,
                        'Found In Column': col,
                        'Matched Value': cell_value,
                        'Similarity %': f"{score * 100:.1f}%",
                        'GSNR': row.get('GSNR', 'N/A'),
                        'RETAILERNAME': row.get('RETAILERNAME', 'N/A'),
                        'COMMENT': row.get('COMMENT', 'N/A'),
                        'COMPANY': row.get('COMPANY', 'N/A'),
                        'NAME 1': row.get('NAME 1', 'N/A'),
                        'NAME 2': row.get('NAME 2', 'N/A'),
                        'STATUS': row.get('STATUS', 'N/A')
                    }
                    all_matches.append(match_info)

                    print(f"\n{'='*80}")
                    print(f"MATCH FOUND for: '{name}'")
                    print(f"Found in column: {col}")
                    print(f"Matched value:   {cell_value}")
                    print(f"Similarity:      {score * 100:.1f}%")
                    print(f"-" * 40)
                    print(f"GSNR:          {row.get('GSNR', 'N/A')}")
                    print(f"RETAILERNAME:  {row.get('RETAILERNAME', 'N/A')}")
                    print(f"COMMENT:       {row.get('COMMENT', 'N/A')}")
                    print(f"COMPANY:       {row.get('COMPANY', 'N/A')}")
                    print(f"NAME 1:        {row.get('NAME 1', 'N/A')}")
                    print(f"NAME 2:        {row.get('NAME 2', 'N/A')}")
                    print(f"STATUS:        {row.get('STATUS', 'N/A')}")

        if not matches_found:
            print(f"\nNo match found for: '{name}'")

    # Create summary DataFrame
    if all_matches:
        results_df = pd.DataFrame(all_matches)

        # Keep ALL results including duplicates found in different columns

        # Save results to Excel
        output_file = 'name_check_results.xlsx'
        results_df.to_excel(output_file, index=False)

        print(f"\n\n{'='*100}")
        print(f"SUMMARY")
        print(f"{'='*100}")
        print(f"Total names checked: {len(names_to_check)}")
        print(f"Total matches found: {len(results_df)}")
        print(f"Results saved to: {output_file}")
    else:
        print(f"\n\nNo matches found for any of the {len(names_to_check)} names.")

    return all_matches


if __name__ == "__main__":
    # Default file paths
    names_file = "Calculus-list.xlsx"  # Your single-column file with names to check
    data_file = "MDM.xlsx"             # Your large Excel file with all the columns
    threshold = 0.6                    # 60% similarity threshold (adjust as needed)

    # Allow command line arguments
    if len(sys.argv) >= 3:
        names_file = sys.argv[1]
        data_file = sys.argv[2]
    if len(sys.argv) >= 4:
        threshold = float(sys.argv[3])

    print(f"""
╔══════════════════════════════════════════════════════════════╗
║      NAME CHECKER - Fuzzy Search Tool                        ║
╠══════════════════════════════════════════════════════════════╣
║  Checking columns: RETAILERNAME, COMMENT, COMPANY,           ║
║                    NAME 1, NAME 2                            ║
║  Uses FUZZY MATCHING to find similar names                   ║
╚══════════════════════════════════════════════════════════════╝
    """)

    check_names_in_excel(names_file, data_file, threshold)
