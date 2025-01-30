# PDF Roster Extractor

## Description
This script extracts student roster information (names, emails, and room numbers) from a PDF file and saves it to an Excel file.

## Requirements
- Python 3.x
- Required libraries:
  - `re`
  - `pandas`
  - `openpyxl`
  - `pypdf`
  - `logging`

Install dependencies using:
```bash
pip install pandas openpyxl pypdf
```
## Functions Overview
### validate_pages(pages)
- Ensures the input follows the correct page range format.
### parse_pages(pages)
- Converts page range input into a sorted list of page numbers.
### readPDFPages(page_list, file_path)
- Extracts text from the given pages of the PDF.
### parse_roster(text)
- Uses regex to extract names, emails, and room numbers while filtering out irrelevant words.
### moveToExcel(first, last, email, room, filename)
- Saves the extracted data into an Excel file with properly formatted columns.
### main()
- Handles user input and runs the extraction process.
