import pandas as pd
import os
from glob import glob

def merge_excel_files():
    # Get the directory path from the user
    directory = input("Enter the directory path containing Excel files: ").strip()
    
    # Verify if directory exists
    if not os.path.exists(directory):
        print("❌ Directory does not exist!")
        return
    
    # Get all Excel files in the directory
    excel_files = glob(os.path.join(directory, "*.xlsx")) + glob(os.path.join(directory, "*.xls"))
    
    if not excel_files:
        print("❌ No Excel files found in the directory!")
        return
    
    # Display available Excel files
    print("\n📂 Available Excel files:")
    for i, file in enumerate(excel_files, 1):
        print(f"{i}. {os.path.basename(file)}")
    
    # Get user selection
    print("\n🔹 Enter the numbers of files you want to merge (comma-separated) or type 'all' for all files.")
    print("Example: 1,2,3 or type 'all' to select all files.")
    
    selection = input("Your selection: ").strip().lower()

    selected_files = []

    # Process user selection
    if selection == "all":
        selected_files = excel_files
    else:
        try:
            indices = [int(x.strip()) - 1 for x in selection.split(",") if x.strip().isdigit()]
            
            # Check if all indices are valid
            if any(i < 0 or i >= len(excel_files) for i in indices):
                raise ValueError("❌ One or more selected numbers are out of range.")

            selected_files = [excel_files[i] for i in indices]
        except ValueError:
            print("❌ Invalid selection! Please enter valid numbers.")
            return
    
    if not selected_files:
        print("❌ No files selected!")
        return
    
    # Get output filename from the user
    output_file = input("\n💾 Enter the output Excel file name (e.g., merged_file.xlsx): ").strip()
    if not output_file.endswith(('.xlsx', '.xls')):
        output_file += '.xlsx'
    
    try:
        # Create an Excel writer object
        output_path = os.path.join(directory, output_file)
        writer = pd.ExcelWriter(output_path, engine="xlsxwriter")
        
        # Keep track of sheet names to avoid duplicates
        used_sheet_names = set()
        
        # Process each selected file
        for file in selected_files:
            print(f"🔄 Processing {os.path.basename(file)}...")
            
            # Read all sheets from the Excel file
            xlsx = pd.ExcelFile(file)
            
            # Process each sheet in the file
            for sheet_name in xlsx.sheet_names:
                df = pd.read_excel(xlsx, sheet_name=sheet_name)
                
                # Preserve table format by keeping original sheet names
                new_sheet_name = sheet_name
                if len(new_sheet_name) > 31:  # Excel's sheet name length limit
                    new_sheet_name = new_sheet_name[:31]
                
                # Handle duplicate sheet names
                original_name = new_sheet_name
                counter = 1
                while new_sheet_name in used_sheet_names:
                    new_sheet_name = f"{original_name[:27]}_{counter}"
                    counter += 1
                
                used_sheet_names.add(new_sheet_name)
                
                # Write the sheet to the new workbook
                df.to_excel(writer, sheet_name=new_sheet_name, index=False)
                
                # Auto-adjust column width
                worksheet = writer.sheets[new_sheet_name]
                for i, col in enumerate(df.columns):
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    worksheet.set_column(i, i, max_length + 2)
        
        # Save the file
        writer.close()
        print(f"\n✅ Successfully created '{output_file}' in {directory}")
        print("📊 All sheets and tables have been preserved in their original format!")
        
    except Exception as e:
        print(f"❌ An error occurred: {str(e)}")

if __name__ == "__main__":
    merge_excel_files()
