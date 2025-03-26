import os
from dotenv import load_dotenv
import pandas as pd
import matplotlib.pyplot as plt
import io

def create_visualizations(df, output_path):
    """
    Create visualizations of the data and add to Excel workbook
    
    Args:
    df (pandas.DataFrame): DataFrame containing the data
    output_path (str): Path to the output Excel file
    """
    # Create a Pandas Excel writer using XlsxWriter as the engine
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        # Write the dataframe to a sheet
        df.to_excel(writer, sheet_name='Raw Data', index=False)
        workbook = writer.book
        
        # 1. Bar Chart: Total entries by Year
        plt.figure(figsize=(10, 6))
        year_counts = df['Year of Harvest'].value_counts().sort_index()
        ax = year_counts.plot(kind='bar')
        plt.title('Total Entries by Year of Harvest')
        plt.xlabel('Year')
        plt.ylabel('Number of Entries')
        plt.tight_layout()
        
        # Save bar chart to a BytesIO object
        bar_chart_buf = io.BytesIO()
        plt.savefig(bar_chart_buf, format='png')
        bar_chart_buf.seek(0)
        plt.close()
        
        # 2. Heatmap of Year and City
        plt.figure(figsize=(12, 8))
        pivot_data = df.pivot_table(index='Year of Harvest', 
                                    columns='City', 
                                    aggfunc='size', 
                                    fill_value=0)
        
        # Use seaborn to create a heatmap
        import seaborn as sns
        sns.heatmap(pivot_data, annot=True, cmap='YlGnBu', fmt='g')
        plt.title('Entries by Year and City')
        plt.tight_layout()
        
        # Save heatmap to a BytesIO object
        heatmap_buf = io.BytesIO()
        plt.savefig(heatmap_buf, format='png')
        heatmap_buf.seek(0)
        plt.close()
        
        # 3. Pie Chart: City Distribution
        plt.figure(figsize=(10, 8))
        city_counts = df['City'].value_counts()
        plt.pie(city_counts, labels=city_counts.index, autopct='%1.1f%%')
        plt.title('Distribution of Entries by City')
        
        # Save pie chart to a BytesIO object
        pie_chart_buf = io.BytesIO()
        plt.savefig(pie_chart_buf, format='png')
        pie_chart_buf.seek(0)
        plt.close()
        
        # Add images to worksheets
        # Yearly Entries Bar Chart
        worksheet_bar = workbook.add_worksheet('Yearly Entries Chart')
        worksheet_bar.insert_image('B2', 'bar_chart.png', {'image_data': bar_chart_buf})
        
        # Year-City Heatmap
        worksheet_heatmap = workbook.add_worksheet('Year-City Heatmap')
        worksheet_heatmap.insert_image('B2', 'heatmap.png', {'image_data': heatmap_buf})
        
        # City Distribution Pie Chart
        worksheet_pie = workbook.add_worksheet('City Distribution')
        worksheet_pie.insert_image('B2', 'pie_chart.png', {'image_data': pie_chart_buf})
    
    print("Visualizations created and added to the Excel file.")

def analyze_csv():
    """
    Analyze CSV file and export selected columns to Excel with visualizations
    """
    try:
        # Load environment variables
        load_dotenv()
        
        # Retrieve CSV path from environment variable
        csv_path = os.getenv('CSV_FILE_PATH')
        
        # Validate CSV path
        if not csv_path:
            raise ValueError("CSV_FILE_PATH not set in .env file")
        
        # Check if file exists
        if not os.path.exists(csv_path):
            raise FileNotFoundError(f"CSV file not found at {csv_path}")
        
        # Read the CSV file
        df = pd.read_csv(csv_path)
        
        # Validate column existence
        required_columns = ['Year of Harvest', 'City', 'Date Sown', 'Date of Cut (Last Cut)']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            raise ValueError(f"Missing columns: {', '.join(missing_columns)}")
        
        # Select only the required columns
        selected_df = df[required_columns]
        
        # Create output Excel file with visualizations
        output_path = 'selected_data_with_charts.xlsx'
        create_visualizations(selected_df, output_path)
        
        print(f"Analysis complete. Data exported to {output_path}")
        print(f"Total rows exported: {len(selected_df)}")
    
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    analyze_csv()