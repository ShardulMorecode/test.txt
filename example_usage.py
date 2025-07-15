#!/usr/bin/env python3
"""
Example usage of the Excel to PDF converter
"""

from excel_to_pdf import ExcelToPdfConverter
import pandas as pd

def create_sample_excel():
    """Create a sample Excel file for testing"""
    # Sample data
    data = {
        'Name': ['John Doe', 'Jane Smith', 'Mike Johnson', 'Sarah Wilson'],
        'Age': [28, 35, 42, 29],
        'Department': ['Engineering', 'Marketing', 'Sales', 'HR'],
        'Salary': [75000, 65000, 58000, 62000],
        'Email': ['john.doe@company.com', 'jane.smith@company.com', 
                 'mike.johnson@company.com', 'sarah.wilson@company.com'],
        'Location': ['New York', 'San Francisco', 'Chicago', 'Boston']
    }
    
    df = pd.DataFrame(data)
    df.to_excel('sample_data.xlsx', index=False)
    print("Sample Excel file 'sample_data.xlsx' created!")

def example_usage():
    """Example of how to use the converter programmatically"""
    # Create sample Excel file
    create_sample_excel()
    
    # Method 1: Using the class directly
    print("\n" + "="*50)
    print("Method 1: Using ExcelToPdfConverter class")
    print("="*50)
    
    converter = ExcelToPdfConverter('sample_data.xlsx', 'employee_pdfs')
    converter.convert_all_rows(filename_prefix='employee')
    
    # Method 2: Converting specific sheet
    print("\n" + "="*50)
    print("Method 2: Converting specific sheet")
    print("="*50)
    
    # List all sheets in the Excel file
    sheets = converter.list_sheets()
    print(f"Available sheets: {sheets}")
    
    # Convert only the first sheet
    if sheets:
        converter.convert_all_rows(sheet_name=sheets[0], filename_prefix='record')

if __name__ == "__main__":
    example_usage()