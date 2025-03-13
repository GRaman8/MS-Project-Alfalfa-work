# import os
# import camelot
# import pandas as pd
# from fuzzywuzzy import fuzz

# def extract_tables_with_camelot(pdf_file, pages):
#     """
#     Extract tables from the given PDF file using Camelot.
#     Uses the 'lattice' flavor and falls back to 'stream' if needed.
#     """
#     try:
#         tables = camelot.read_pdf(pdf_file, pages=pages, flavor='lattice')
#         if len(tables) == 0:
#             print(f"No tables found on pages {pages} using lattice mode. Trying stream mode...")
#             tables = camelot.read_pdf(pdf_file, pages=pages, flavor='stream')
#         return tables
#     except Exception as e:
#         print(f"Error extracting tables from {pdf_file} on pages {pages}: {e}")
#         return None

# def read_excel_table(excel_file):
#     """
#     Read the Excel file into a pandas DataFrame.
#     """
#     try:
#         df = pd.read_excel(excel_file)
#         return df
#     except Exception as e:
#         print(f"Error reading Excel file {excel_file}: {e}")
#         return None

# def compare_tables(df_excel, df_pdf, threshold=90):
#     """
#     Compare two DataFrames cell by cell using fuzzy matching.
#     Even if the shapes differ, it will compare over the overlapping region.
#     It then prints out mismatches and warnings if one table has extra cells.
#     """
#     rows_excel, cols_excel = df_excel.shape
#     rows_pdf, cols_pdf = df_pdf.shape
#     min_rows = min(rows_excel, rows_pdf)
#     min_cols = min(cols_excel, cols_pdf)

#     mismatches = 0
#     total_compared = min_rows * min_cols

#     for i in range(min_rows):
#         for j in range(min_cols):
#             cell_excel = str(df_excel.iat[i, j]).strip()
#             cell_pdf = str(df_pdf.iat[i, j]).strip()
#             similarity = fuzz.ratio(cell_excel, cell_pdf)
#             if similarity < threshold:
#                 print(f"Mismatch at cell ({i+1}, {j+1}): Excel='{cell_excel}' vs PDF='{cell_pdf}' (similarity: {similarity}%)")
#                 mismatches += 1

#     if rows_excel != rows_pdf or cols_excel != cols_pdf:
#         print("Warning: The tables have different shapes!")
#         print("Excel table shape:", df_excel.shape)
#         print("PDF extracted table shape:", df_pdf.shape)
#         extra_rows_excel = rows_excel - min_rows
#         extra_cols_excel = cols_excel - min_cols
#         extra_rows_pdf = rows_pdf - min_rows
#         extra_cols_pdf = cols_pdf - min_cols
#         if extra_rows_excel or extra_cols_excel:
#             print("Extra cells in Excel table: {} extra rows, {} extra columns.".format(extra_rows_excel, extra_cols_excel))
#         if extra_rows_pdf or extra_cols_pdf:
#             print("Extra cells in PDF extracted table: {} extra rows, {} extra columns.".format(extra_rows_pdf, extra_cols_pdf))
#         # Optionally, you can decide to treat these extra cells as mismatches.

#     print(f"Total mismatches: {mismatches} out of {total_compared} compared cells.")
#     return mismatches

# def verify_landscape(pdf_file, base_name):
#     """
#     Verify the landscape Excel output.
#     Assumes the landscape Excel file is named as:
#       {base_name}_landscape.xlsx
#     """
#     landscape_excel = base_name + "_landscape.xlsx"
#     if not os.path.exists(landscape_excel):
#         print(f"Landscape Excel file {landscape_excel} not found.")
#         return

#     pages = input(f"Enter the page numbers corresponding to the landscape Excel file ({landscape_excel}) for verification: ").strip()
#     tables = extract_tables_with_camelot(pdf_file, pages)
#     if tables is None or len(tables) == 0:
#         print("No tables extracted using Camelot for landscape pages.")
#         return

#     # For demonstration, we compare against the first table extracted by Camelot.
#     camelot_df = tables[0].df
#     excel_df = read_excel_table(landscape_excel)
#     if excel_df is None:
#         return

#     print("\nComparing landscape table:")
#     mismatches = compare_tables(excel_df, camelot_df)
#     if mismatches == 0:
#         print("Landscape Excel table verification PASSED.")
#     else:
#         print("Landscape Excel table verification FAILED.")

# def verify_portrait(pdf_file, base_name):
#     """
#     Verify each portrait Excel output.
#     Assumes portrait Excel files are named as:
#       {base_name}_portrait_page_{p}.xlsx
#     """
#     portrait_files = [f for f in os.listdir('.') if f.startswith(base_name + "_portrait_page_") and f.endswith('.xlsx')]
#     if not portrait_files:
#         print("No portrait Excel files found for verification.")
#         return

#     for portrait_excel in portrait_files:
#         try:
#             page_str = portrait_excel.split("_portrait_page_")[1].split(".xlsx")[0]
#         except IndexError:
#             print(f"Could not parse page number from {portrait_excel}")
#             continue

#         pages = page_str  # One page per portrait file
#         tables = extract_tables_with_camelot(pdf_file, pages)
#         if tables is None or len(tables) == 0:
#             print(f"No tables extracted using Camelot for portrait page {page_str}.")
#             continue

#         camelot_df = tables[0].df
#         excel_df = read_excel_table(portrait_excel)
#         if excel_df is None:
#             continue

#         print(f"\nComparing portrait page {page_str} table:")
#         mismatches = compare_tables(excel_df, camelot_df)
#         if mismatches == 0:
#             print(f"Portrait page {page_str} Excel table verification PASSED.")
#         else:
#             print(f"Portrait page {page_str} Excel table verification FAILED.")

# def main():
#     pdf_file = input("Enter the path to the PDF file for verification: ").strip()
#     if not os.path.exists(pdf_file):
#         print("PDF file does not exist.")
#         return

#     base_name = os.path.splitext(os.path.basename(pdf_file))[0]
#     print(f"Verifying tables for PDF: {pdf_file}")
    
#     verify_landscape(pdf_file, base_name)
#     verify_portrait(pdf_file, base_name)

# if __name__ == "__main__":
#     main()

import os
import camelot
import pandas as pd
from fuzzywuzzy import fuzz

def extract_tables_with_camelot(pdf_file, pages):
    """
    Extract tables from the given PDF file using Camelot.
    Uses the 'lattice' flavor and falls back to 'stream' if needed.
    """
    try:
        tables = camelot.read_pdf(pdf_file, pages=pages, flavor='lattice')
        if len(tables) == 0:
            print(f"No tables found on pages {pages} using lattice mode. Trying stream mode...")
            tables = camelot.read_pdf(pdf_file, pages=pages, flavor='stream')
        return tables
    except Exception as e:
        print(f"Error extracting tables from {pdf_file} on pages {pages}: {e}")
        return None

def read_excel_table(excel_file):
    """
    Read the Excel file into a pandas DataFrame.
    """
    try:
        df = pd.read_excel(excel_file)
        return df
    except Exception as e:
        print(f"Error reading Excel file {excel_file}: {e}")
        return None

def compare_agricultural_data(df_excel, df_pdf, threshold=90):
    """
    Compare variety names and yield data between Excel and PDF tables.
    Focuses on the specific columns of interest for agricultural data.
    """
    # Try to identify the variety column in both dataframes
    variety_cols_excel = [col for col in df_excel.columns if 'variety' in str(col).lower()]
    variety_cols_pdf = [col for col in df_pdf.columns if 'variety' in str(col).lower()]
    
    # If variety columns weren't found, try to use the first column
    if not variety_cols_excel:
        variety_cols_excel = [df_excel.columns[0]]
    if not variety_cols_pdf:
        variety_cols_pdf = [df_pdf.columns[0]]
    
    # Find yield columns in both dataframes
    yield_cols_excel = [col for col in df_excel.columns if 'yield' in str(col).lower() 
                        or 'ton' in str(col).lower() or 'acre' in str(col).lower()]
    yield_cols_pdf = [col for col in df_pdf.columns if 'yield' in str(col).lower() 
                      or 'ton' in str(col).lower() or 'acre' in str(col).lower()]
    
    # If specific yield columns weren't found, use numeric columns (excluding the first one)
    if not yield_cols_excel:
        yield_cols_excel = [col for col in df_excel.columns[1:] 
                            if pd.api.types.is_numeric_dtype(df_excel[col])]
    if not yield_cols_pdf:
        yield_cols_pdf = [col for col in df_pdf.columns[1:] 
                          if pd.api.types.is_numeric_dtype(df_pdf[col])]
    
    # Extract variety names
    varieties_excel = df_excel[variety_cols_excel[0]].dropna().astype(str).tolist()
    varieties_pdf = df_pdf[variety_cols_pdf[0]].dropna().astype(str).tolist()
    
    # Compare variety names
    variety_matches = []
    variety_mismatches = []
    
    for variety_excel in varieties_excel:
        best_match = None
        best_score = 0
        
        for variety_pdf in varieties_pdf:
            score = fuzz.ratio(variety_excel.strip().lower(), variety_pdf.strip().lower())
            if score > best_score:
                best_score = score
                best_match = variety_pdf
        
        if best_score >= threshold:
            variety_matches.append((variety_excel, best_match, best_score))
        else:
            variety_mismatches.append((variety_excel, best_match, best_score))
    
    # Compare yield values for matched varieties
    yield_results = []
    severe_mismatches = 0
    minor_mismatches = 0
    total_yield_comparisons = 0
    
    for variety_excel, variety_pdf, _ in variety_matches:
        excel_row = df_excel[df_excel[variety_cols_excel[0]] == variety_excel]
        pdf_row = df_pdf[df_pdf[variety_cols_pdf[0]] == variety_pdf]
        
        if excel_row.empty or pdf_row.empty:
            continue
        
        for yield_col_excel in yield_cols_excel:
            for yield_col_pdf in yield_cols_pdf:
                try:
                    yield_excel = float(excel_row[yield_col_excel].values[0])
                    yield_pdf = float(pdf_row[yield_col_pdf].values[0])
                    
                    # Calculate percentage difference
                    if yield_excel == 0 and yield_pdf == 0:
                        diff_pct = 0
                    elif yield_excel == 0:
                        diff_pct = 100
                    else:
                        diff_pct = abs((yield_excel - yield_pdf) / yield_excel * 100)
                    
                    match_score = 100 - diff_pct
                    
                    if match_score < threshold:
                        if match_score < 50:
                            severe_mismatches += 1
                        else:
                            minor_mismatches += 1
                        
                        yield_results.append((
                            variety_excel, 
                            str(yield_col_excel), 
                            str(yield_col_pdf),
                            yield_excel, 
                            yield_pdf, 
                            match_score
                        ))
                    
                    total_yield_comparisons += 1
                    
                except (ValueError, IndexError):
                    continue
    
    # Sort yield results by match score (ascending)
    yield_results.sort(key=lambda x: x[5])
    
    # Keep only the 5 worst mismatches
    yield_results = yield_results[:5]
    
    # Calculate overall match percentage
    total_mismatches = severe_mismatches + minor_mismatches
    total_comparisons = len(variety_matches) + len(variety_mismatches) + total_yield_comparisons
    
    if total_comparisons > 0:
        match_percentage = 100 - (total_mismatches + len(variety_mismatches)) / total_comparisons * 100
    else:
        match_percentage = 0
    
    # Prepare results
    result = {
        "match_percentage": round(match_percentage, 2),
        "variety_matches": len(variety_matches),
        "variety_mismatches": variety_mismatches,
        "yield_comparisons": total_yield_comparisons,
        "severe_yield_mismatches": severe_mismatches,
        "minor_yield_mismatches": minor_mismatches,
        "yield_mismatch_examples": yield_results
    }
    
    return result

def print_agricultural_results(result, table_name):
    """Print agricultural verification results in a simplified format"""
    print("\n" + "="*50)
    print(f"VERIFICATION RESULTS: {table_name}")
    print("="*50)
    
    # Overall match percentage with status indicator
    match_percentage = result["match_percentage"]
    if match_percentage >= 95:
        status = "PASSED ✓"
    elif match_percentage >= 85:
        status = "PASSED WITH MINOR ISSUES ⚠"
    else:
        status = "FAILED ✗"
        
    print(f"Status: {status}")
    print(f"Overall Match Percentage: {match_percentage}%")
    
    # Variety matching summary
    print(f"\nVariety Name Matching:")
    print(f"  Correctly Matched: {result['variety_matches']}")
    
    # Variety mismatches if any
    if result["variety_mismatches"]:
        print(f"  Mismatched Varieties: {len(result['variety_mismatches'])}")
        print("\nVariety Mismatch Examples:")
        for variety_excel, variety_pdf, score in result["variety_mismatches"][:3]:  # Show up to 3 examples
            print(f"  Excel: '{variety_excel}' → PDF: '{variety_pdf}' (Similarity: {score}%)")
    
    # Yield comparison summary
    print(f"\nYield Value Comparison:")
    print(f"  Total Comparisons: {result['yield_comparisons']}")
    print(f"  Severe Mismatches: {result['severe_yield_mismatches']}")
    print(f"  Minor Mismatches: {result['minor_yield_mismatches']}")
    
    # Yield mismatch examples if any
    if result["yield_mismatch_examples"]:
        print("\nYield Mismatch Examples:")
        for variety, col_excel, col_pdf, yield_excel, yield_pdf, score in result["yield_mismatch_examples"]:
            print(f"  Variety: '{variety}'")
            print(f"    Excel ({col_excel}): {yield_excel}")
            print(f"    PDF ({col_pdf}): {yield_pdf}")
            print(f"    Match Score: {score:.1f}%")
    
    print("="*50)

def verify_landscape(pdf_file, base_name):
    """
    Verify the landscape Excel output with focus on agricultural data.
    """
    landscape_excel = base_name + "_landscape.xlsx"
    if not os.path.exists(landscape_excel):
        print(f"Landscape Excel file {landscape_excel} not found.")
        return

    pages = input(f"Enter the page numbers corresponding to the landscape Excel file ({landscape_excel}) for verification: ").strip()
    tables = extract_tables_with_camelot(pdf_file, pages)
    if tables is None or len(tables) == 0:
        print("No tables extracted using Camelot for landscape pages.")
        return

    # For demonstration, we compare against the first table extracted by Camelot.
    camelot_df = tables[0].df
    excel_df = read_excel_table(landscape_excel)
    if excel_df is None:
        return

    result = compare_agricultural_data(excel_df, camelot_df)
    print_agricultural_results(result, f"Landscape Table ({pages})")

def verify_portrait(pdf_file, base_name):
    """
    Verify each portrait Excel output with focus on agricultural data.
    """
    portrait_files = [f for f in os.listdir('.') if f.startswith(base_name + "_portrait_page_") and f.endswith('.xlsx')]
    if not portrait_files:
        print("No portrait Excel files found for verification.")
        return

    for portrait_excel in portrait_files:
        try:
            page_str = portrait_excel.split("_portrait_page_")[1].split(".xlsx")[0]
        except IndexError:
            print(f"Could not parse page number from {portrait_excel}")
            continue

        pages = page_str  # One page per portrait file
        tables = extract_tables_with_camelot(pdf_file, pages)
        if tables is None or len(tables) == 0:
            print(f"No tables extracted using Camelot for portrait page {page_str}.")
            continue

        camelot_df = tables[0].df
        excel_df = read_excel_table(portrait_excel)
        if excel_df is None:
            continue

        result = compare_agricultural_data(excel_df, camelot_df)
        print_agricultural_results(result, f"Portrait Page {page_str}")

def main():
    pdf_file = input("Enter the path to the PDF file for verification: ").strip()
    if not os.path.exists(pdf_file):
        print("PDF file does not exist.")
        return

    base_name = os.path.splitext(os.path.basename(pdf_file))[0]
    print(f"Verifying tables for PDF: {pdf_file}")
    
    verify_landscape(pdf_file, base_name)
    verify_portrait(pdf_file, base_name)

if __name__ == "__main__":
    main()
