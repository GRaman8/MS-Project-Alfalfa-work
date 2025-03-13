import pandas as pd

def load_excel(file_path):
    """Load the Excel file and display available sheets."""
    try:
        xl = pd.ExcelFile(file_path)
        print("\nAvailable sheets in the Excel file:")
        for i, sheet in enumerate(xl.sheet_names, 1):
            print(f"{i}. {sheet}")

        sheet_name = input("\nEnter the sheet name you want to process: ").strip()
        if sheet_name not in xl.sheet_names:
            print("Invalid sheet name. Exiting.")
            return None

        df = xl.parse(sheet_name)
        return df, sheet_name

    except Exception as e:
        print(f"Error loading file: {e}")
        return None


def select_columns(df):
    """Ask the user which columns to keep."""
    print("\nAvailable columns:")
    print(list(df.columns))

    selected_columns = input("\nEnter the columns you want to keep (comma-separated): ").strip().split(",")

    selected_columns = [col.strip() for col in selected_columns if col.strip() in df.columns]

    if not selected_columns:
        print("No valid columns selected. Exiting.")
        return None

    return df[selected_columns]


def remove_rows(df):
    """Ask the user whether they want to remove any rows."""
    print("\nFirst few rows of data:")
    print(df.head())

    remove_option = input("\nDo you want to remove any rows? (Y/N): ").strip().lower()
    if remove_option == "y":
        remove_type = input("Remove by (1) row index or (2) condition (column value)? Enter 1 or 2: ").strip()
        
        if remove_type == "1":
            indices = input("Enter row indices to remove (comma-separated): ").strip().split(",")
            indices = [int(i) for i in indices if i.isdigit()]
            df = df.drop(indices)
        
        elif remove_type == "2":
            column = input("Enter column name for condition-based removal: ").strip()
            if column in df.columns:
                value = input(f"Enter value to remove rows where '{column}' = value: ").strip()
                df = df[df[column] != value]
            else:
                print("Invalid column name.")

    return df


def save_excel(df, original_sheet_name):
    """Save the cleaned data to a new Excel file."""
    output_file = input("\nEnter the output Excel file name (e.g., cleaned_data.xlsx): ").strip()
    if not output_file.endswith(".xlsx"):
        output_file += ".xlsx"

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=original_sheet_name, index=False)

    print(f"✅ Cleaned Excel file saved as {output_file}")


def main():
    file_path = input("Enter the full path of the Excel file: ").strip()

    data = load_excel(file_path)
    if data is None:
        return

    df, sheet_name = data
    df = select_columns(df)
    if df is None:
        return

    df = remove_rows(df)
    save_excel(df, sheet_name)


if __name__ == "__main__":
    main()
