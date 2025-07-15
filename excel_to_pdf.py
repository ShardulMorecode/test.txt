import pandas as pd
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import os
import sys
from pathlib import Path

class ExcelToPdfConverter:
    def __init__(self, excel_file_path, output_folder="pdf_output"):
        """
        Initialize the converter with Excel file path and output folder
        
        Args:
            excel_file_path (str): Path to the Excel file
            output_folder (str): Folder where PDFs will be saved
        """
        self.excel_file_path = excel_file_path
        self.output_folder = output_folder
        self.styles = getSampleStyleSheet()
        
        # Create output folder if it doesn't exist
        Path(self.output_folder).mkdir(parents=True, exist_ok=True)
    
    def read_excel_data(self, sheet_name=None):
        """
        Read data from Excel file
        
        Args:
            sheet_name (str): Name of the sheet to read (optional)
        
        Returns:
            pd.DataFrame: DataFrame containing the Excel data
        """
        try:
            if sheet_name:
                df = pd.read_excel(self.excel_file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(self.excel_file_path)
            
            # Replace NaN values with empty strings for better PDF formatting
            df = df.fillna('')
            
            print(f"Successfully read {len(df)} rows and {len(df.columns)} columns from Excel file")
            return df
        
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return None
    
    def create_pdf_from_row(self, row_data, row_index, filename_prefix="row"):
        """
        Create a PDF from a single row of data
        
        Args:
            row_data (pd.Series): Row data from DataFrame
            row_index (int): Index of the row
            filename_prefix (str): Prefix for the PDF filename
        """
        # Create filename
        pdf_filename = f"{filename_prefix}_{row_index + 1:03d}.pdf"
        pdf_path = os.path.join(self.output_folder, pdf_filename)
        
        # Create PDF document
        doc = SimpleDocTemplate(pdf_path, pagesize=A4, 
                              rightMargin=72, leftMargin=72,
                              topMargin=72, bottomMargin=18)
        
        # Container for the 'Flowable' objects
        elements = []
        
        # Add title
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            spaceAfter=30,
            alignment=TA_CENTER
        )
        
        title = Paragraph(f"Record {row_index + 1}", title_style)
        elements.append(title)
        elements.append(Spacer(1, 12))
        
        # Create table data
        table_data = []
        
        for column_name, value in row_data.items():
            # Convert value to string and handle special cases
            if pd.isna(value):
                value_str = ""
            elif isinstance(value, (int, float)):
                value_str = str(value)
            else:
                value_str = str(value)
            
            table_data.append([str(column_name), value_str])
        
        # Create table
        table = Table(table_data, colWidths=[2.5*inch, 4*inch])
        
        # Add table styling
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ('BACKGROUND', (1, 0), (1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        
        elements.append(table)
        
        # Build PDF
        try:
            doc.build(elements)
            print(f"Created PDF: {pdf_filename}")
        except Exception as e:
            print(f"Error creating PDF {pdf_filename}: {e}")
    
    def convert_all_rows(self, sheet_name=None, filename_prefix="row"):
        """
        Convert all rows in the Excel file to individual PDFs
        
        Args:
            sheet_name (str): Name of the sheet to process (optional)
            filename_prefix (str): Prefix for PDF filenames
        """
        # Read Excel data
        df = self.read_excel_data(sheet_name)
        
        if df is None:
            return
        
        print(f"Processing {len(df)} rows...")
        
        # Convert each row to PDF
        for index, row in df.iterrows():
            self.create_pdf_from_row(row, index, filename_prefix)
        
        print(f"Conversion complete! {len(df)} PDFs created in '{self.output_folder}' folder")
    
    def list_sheets(self):
        """
        List all sheet names in the Excel file
        
        Returns:
            list: List of sheet names
        """
        try:
            xls = pd.ExcelFile(self.excel_file_path)
            return xls.sheet_names
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return []

def main():
    """
    Main function to run the Excel to PDF converter
    """
    print("Excel to PDF Converter")
    print("=" * 50)
    
    # Check if Excel file path is provided as command line argument
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        excel_file = input("Enter the path to your Excel file: ")
    
    # Check if file exists
    if not os.path.exists(excel_file):
        print(f"Error: File '{excel_file}' not found!")
        return
    
    # Create converter instance
    converter = ExcelToPdfConverter(excel_file)
    
    # List available sheets
    sheets = converter.list_sheets()
    print(f"\nAvailable sheets: {sheets}")
    
    # Ask user which sheet to process
    if len(sheets) > 1:
        print("\nMultiple sheets found:")
        for i, sheet in enumerate(sheets):
            print(f"{i + 1}. {sheet}")
        
        choice = input("\nEnter sheet number (or press Enter for first sheet): ")
        
        if choice.strip():
            try:
                sheet_index = int(choice) - 1
                if 0 <= sheet_index < len(sheets):
                    selected_sheet = sheets[sheet_index]
                else:
                    print("Invalid selection, using first sheet")
                    selected_sheet = sheets[0]
            except ValueError:
                print("Invalid input, using first sheet")
                selected_sheet = sheets[0]
        else:
            selected_sheet = sheets[0]
    else:
        selected_sheet = sheets[0] if sheets else None
    
    # Ask for filename prefix
    filename_prefix = input("Enter filename prefix (default: 'row'): ").strip()
    if not filename_prefix:
        filename_prefix = "row"
    
    # Ask for output folder
    output_folder = input("Enter output folder name (default: 'pdf_output'): ").strip()
    if not output_folder:
        output_folder = "pdf_output"
    
    # Update converter with custom output folder
    converter.output_folder = output_folder
    Path(converter.output_folder).mkdir(parents=True, exist_ok=True)
    
    # Convert all rows
    converter.convert_all_rows(selected_sheet, filename_prefix)

if __name__ == "__main__":
    main()