import pandas as pd

# Load the Excel file
file_path = "Lexington_(2001-2024)_Temp_Prep.xlsx"
df = pd.read_excel(file_path, sheet_name="Lexington_(2001-2024)_Temp_Prep", skiprows=6)

# Set proper headers and drop metadata rows
df.columns = df.iloc[1]
df = df[2:].reset_index(drop=True)

# Rename first column to 'Date'
df.rename(columns={df.columns[0]: 'Date'}, inplace=True)

# Remove rows where 'Date' is NaN or clearly not a date (like 'Sum:')
df = df[df['Date'].notna()]  # Remove NaN rows
df = df[~df['Date'].astype(str).str.contains("Sum", case=False)]  # Remove 'Sum:' or similar

# Convert to datetime
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

# Drop any rows where datetime conversion failed
df = df[df['Date'].notna()]

# Set date as index and transpose
df.set_index('Date', inplace=True)
df_transposed = df.transpose()

# Format dates as YYYY-MM for column headers
df_transposed.columns = [d.strftime("%Y-%m") for d in df_transposed.columns]

# Save to Excel
df_transposed.to_excel("Lexington_Transposed.xlsx")

print("✅ Cleaned, transposed file saved as 'Lexington_Transposed.xlsx'")
