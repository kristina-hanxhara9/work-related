import pandas as pd
import sys

def check_names_in_excel(names_file, data_file):
    """
    Check names from a single-column Excel file against specific columns in a larger Excel file.
    Shows all matching rows with their GSNR values.
    """

    # Read the names file (single column)
    print(f"Reading names from: {names_file}")
    names_df = pd.read_excel(names_file, header=None)
    names_to_check = names_df.iloc[:, 0].dropna().astype(str).str.strip().str.lower().tolist()
    print(f"Found {len(names_to_check)} names to check\n")

    # Read the large data file
    print(f"Reading data from: {data_file}")
    data_df = pd.read_excel(data_file)
    print(f"Data file has {len(data_df)} rows\n")

    # Columns to search in
    search_columns = ['RETAILERNAME', 'COMMENT', 'COMPANYNAME 1', 'NAME 2']

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

        for col in existing_columns:
            # Convert column to string and lowercase for comparison
            col_values = data_df[col].fillna('').astype(str).str.lower()

            # Find rows where the name appears (partial match)
            mask = col_values.str.contains(name, case=False, na=False, regex=False)
            matching_rows = data_df[mask]

            if not matching_rows.empty:
                matches_found = True
                for idx, row in matching_rows.iterrows():
                    match_info = {
                        'Search Name': name,
                        'Found In Column': col,
                        'GSNR': row.get('GSNR', 'N/A'),
                        'RETAILERNAME': row.get('RETAILERNAME', 'N/A'),
                        'COMMENT': row.get('COMMENT', 'N/A'),
                        'COMPANYNAME 1': row.get('COMPANYNAME 1', 'N/A'),
                        'NAME 2': row.get('NAME 2', 'N/A'),
                        'STATUS': row.get('STATUS', 'N/A'),
                        'CITY': row.get('CITY', 'N/A'),
                        'COUNTRY': row.get('COUNTRY', 'N/A')
                    }
                    all_matches.append(match_info)

                    print(f"\n{'='*80}")
                    print(f"MATCH FOUND for: '{name}'")
                    print(f"Found in column: {col}")
                    print(f"-" * 40)
                    print(f"GSNR:          {row.get('GSNR', 'N/A')}")
                    print(f"RETAILERNAME:  {row.get('RETAILERNAME', 'N/A')}")
                    print(f"COMMENT:       {row.get('COMMENT', 'N/A')}")
                    print(f"COMPANYNAME 1: {row.get('COMPANYNAME 1', 'N/A')}")
                    print(f"NAME 2:        {row.get('NAME 2', 'N/A')}")
                    print(f"STATUS:        {row.get('STATUS', 'N/A')}")
                    print(f"CITY:          {row.get('CITY', 'N/A')}")
                    print(f"COUNTRY:       {row.get('COUNTRY', 'N/A')}")

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
    # Default file paths - UPDATE THESE
    names_file = "names.xlsx"  # Your single-column file with names to check
    data_file = "data.xlsx"    # Your large Excel file with all the columns

    # Allow command line arguments
    if len(sys.argv) >= 3:
        names_file = sys.argv[1]
        data_file = sys.argv[2]

    print(f"""
╔══════════════════════════════════════════════════════════════╗
║           NAME CHECKER - Excel Search Tool                   ║
╠══════════════════════════════════════════════════════════════╣
║  Checking columns: RETAILERNAME, COMMENT,                    ║
║                    COMPANYNAME 1, NAME 2                     ║
╚══════════════════════════════════════════════════════════════╝
    """)

    check_names_in_excel(names_file, data_file)
