# Excel to PDF Converter

A Python tool to convert Excel spreadsheet data into individual PDF files, with each row of data generating a separate PDF document.

## Features

- **Row-by-row conversion**: Each row in your Excel file becomes a separate PDF
- **Multi-sheet support**: Handle Excel files with multiple sheets
- **Professional formatting**: Clean, table-formatted PDFs with proper styling
- **Customizable output**: Choose filename prefixes and output folders
- **Error handling**: Robust error handling for various data types and edge cases
- **Interactive and programmatic usage**: Use via command line or as a Python library

## Installation

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

Or install manually:
```bash
pip install pandas reportlab openpyxl
```

## Usage

### Method 1: Interactive Command Line

Run the script interactively:
```bash
python excel_to_pdf.py
```

Or provide the Excel file path directly:
```bash
python excel_to_pdf.py path/to/your/excel_file.xlsx
```

The script will prompt you for:
- Sheet selection (if multiple sheets exist)
- Filename prefix for PDFs
- Output folder name

### Method 2: Programmatic Usage

```python
from excel_to_pdf import ExcelToPdfConverter

# Create converter instance
converter = ExcelToPdfConverter('your_file.xlsx', 'output_folder')

# Convert all rows to PDFs
converter.convert_all_rows(filename_prefix='document')

# Convert specific sheet
converter.convert_all_rows(sheet_name='Sheet1', filename_prefix='record')

# List available sheets
sheets = converter.list_sheets()
print(f"Available sheets: {sheets}")
```

### Method 3: Example with Sample Data

Run the example to see the converter in action:
```bash
python example_usage.py
```

This will:
1. Create a sample Excel file with employee data
2. Convert each row to a separate PDF
3. Demonstrate different usage patterns

## Output Format

Each PDF contains:
- **Title**: "Record X" where X is the row number
- **Table format**: Two columns showing field names and values
- **Professional styling**: Clean borders, alternating colors, proper fonts

Example PDF content:
```
                    Record 1
    
    ┌─────────────────┬──────────────────────────┐
    │ Name            │ John Doe                 │
    │ Age             │ 28                       │
    │ Department      │ Engineering              │
    │ Salary          │ 75000                    │
    │ Email           │ john.doe@company.com     │
    │ Location        │ New York                 │
    └─────────────────┴──────────────────────────┘
```

## File Structure

```
.
├── excel_to_pdf.py      # Main converter class and CLI script
├── example_usage.py     # Example usage with sample data
├── requirements.txt     # Python dependencies
└── README.md           # This documentation
```

## Features in Detail

### Data Handling
- **Empty cells**: Automatically handled as empty strings
- **Different data types**: Numbers, text, dates are properly formatted
- **Special characters**: Properly escaped in PDF output

### PDF Formatting
- **Page size**: A4 format with proper margins
- **Font**: Helvetica family for readability
- **Colors**: Light gray headers, white data cells
- **Layout**: Two-column table (field name | value)

### Error Handling
- **File not found**: Clear error messages
- **Invalid Excel files**: Graceful handling with error reporting
- **PDF generation errors**: Individual file failures don't stop the process

## Customization

### Changing PDF Layout
Modify the `create_pdf_from_row` method in `ExcelToPdfConverter` class:
- Change page size: `pagesize=letter` instead of `pagesize=A4`
- Adjust margins: Modify `rightMargin`, `leftMargin`, etc.
- Update styling: Change colors, fonts, and table formatting

### Custom Filename Patterns
The filename format is: `{prefix}_{row_number:03d}.pdf`
- `prefix`: Customizable string
- `row_number`: Auto-incremented (001, 002, etc.)

## Example Use Cases

1. **Employee Records**: Convert employee database to individual PDF profiles
2. **Product Catalogs**: Create separate PDFs for each product
3. **Customer Information**: Generate customer detail sheets
4. **Invoice Generation**: Create individual invoices from spreadsheet data
5. **Certificate Generation**: Bulk create certificates with personalized data

## Troubleshooting

### Common Issues

1. **ModuleNotFoundError**: Install missing dependencies
   ```bash
   pip install pandas reportlab openpyxl
   ```

2. **File not found**: Ensure the Excel file path is correct
3. **Permission errors**: Check write permissions in output folder
4. **Empty PDFs**: Verify Excel file has data in expected format

### Debug Mode

Add debug prints by modifying the converter class or run with Python's verbose mode:
```bash
python -v excel_to_pdf.py
```

## Contributing

Feel free to submit issues and enhancement requests!

## License

This project is open source and available under the MIT License.