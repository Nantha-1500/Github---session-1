"""
Script to extract test data from MSME-Base-Test-Cases-V1.xlsx
and generate a Test Summary Report
"""

import pandas as pd
import json
from datetime import datetime

def extract_test_data(excel_file):
    """Extract all data from Excel file"""
    
    # Read Excel file
    excel_file_obj = pd.ExcelFile(excel_file)
    
    print(f"Sheet names found: {excel_file_obj.sheet_names}\n")
    
    all_data = {}
    
    # Extract data from each sheet
    for sheet in excel_file_obj.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet)
        all_data[sheet] = df
        print(f"Sheet: {sheet}")
        print(f"Rows: {len(df)}, Columns: {len(df.columns)}")
        print(f"Columns: {list(df.columns)}\n")
        print(df.head(10))
        print("\n" + "="*80 + "\n")
    
    return all_data

def generate_summary(data):
    """Generate summary statistics from test data"""
    
    summary = {
        "report_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "sheets": {}
    }
    
    for sheet_name, df in data.items():
        sheet_summary = {
            "total_rows": len(df),
            "total_columns": len(df.columns),
            "columns": list(df.columns),
            "data": df.to_dict(orient='records')
        }
        summary["sheets"][sheet_name] = sheet_summary
    
    return summary

def main():
    excel_file = 'MSME-Base-Test-Cases-V1.xlsx'
    
    try:
        print("Extracting test data from Excel file...\n")
        data = extract_test_data(excel_file)
        
        print("\nGenerating summary report...\n")
        summary = generate_summary(data)
        
        # Save summary to JSON
        with open('test_summary.json', 'w') as f:
            json.dump(summary, f, indent=2, default=str)
        
        print("✓ Test summary saved to 'test_summary.json'")
        print("✓ Data extraction complete!")
        
    except FileNotFoundError:
        print(f"Error: File '{excel_file}' not found!")
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
