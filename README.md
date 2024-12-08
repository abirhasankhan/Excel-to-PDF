
# Excel to PDF Converter

This project allows users to upload an Excel file, select specific sheets or export all sheets, and convert them into PDF files. Each sheet is displayed in a table-like format in landscape orientation, with wide tables split into multiple pages for better readability.

## Features

- **Upload Excel File**: Supports multi-sheet Excel files.
- **Sheet Selection**: Choose specific sheets to export or process all sheets at once.
- **Dynamic Column Splitting**: Handles wide tables by splitting columns across multiple pages.
- **Landscape Orientation**: Ensures readability for large tables.
- **Automatic PDF Download**: After PDF generation, files are available for download.

## Requirements

- Python
- Libraries:
  - `pandas`
  - `fpdf`
  - `google.colab` (for use in Google Colab)

## Installation

Install the required dependencies using `pip`:

```bash
pip install pandas fpdf
```

> Note: This project is designed to run on Google Colab.

## How to Use

1. **Upload Excel File**:
   - Run the script in Google Colab.
   - Use the file upload dialog to upload your Excel file.

2. **View Sheet Names**:
   - The script will display all available sheet names in the Excel file.

3. **Preview Data**:
   - A preview of the first few rows of each sheet will be displayed in the Colab output.

4. **Select Sheets**:
   - Input the names of the sheets to export as a comma-separated list (e.g., `Sheet1, Sheet2`).
   - Type `all` to process all sheets.

5. **Generate PDF**:
   - The script will process the selected sheets, convert them to PDFs, and save the files locally.

6. **Download PDF**:
   - PDFs will be automatically downloaded after generation.

## Example

### Input: Excel File
- Sheet1: Employee Details
- Sheet2: Sales Data

### Output:
- `Employee Details.pdf`
- `Sales Data.pdf`

Each PDF contains a table-like format of the sheet’s content.

## Code Highlights

- **Dynamic Column Splitting**:
  Handles tables with more columns than can fit on a single page:
  ```python
  column_chunks = [data.columns[i:i + max_columns_per_page] for i in range(0, total_columns, max_columns_per_page)]
  ```

- **PDF Table Generation**:
  Uses `fpdf` to generate table-like structures:
  ```python
  for row in data[col_chunk].itertuples(index=False):
      for value in row:
          pdf.cell(col_width, 8, str(value), border=1, align='C')
      pdf.ln()
  ```

## Limitations

- Currently supports `.xlsx` files.
- Designed for use in Google Colab due to the `google.colab` library for file upload/download.

## Author

Created by [Abir Khan].

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
