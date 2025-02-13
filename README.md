# Business Data Cleansing Tool

A Ruby script for processing and updating business data from Excel spreadsheets.

## Setup

1. Install dependencies:
```bash
bundle install
```

2. Create a `.env` file in the root directory with the following variables:
```bash
EXCEL_FILE_PATH=path/to/your/excel/file.xlsx
MARKETPLACES_CSV_PATH=path/to/your/marketplaces.csv
OUTPUT_FILE_PATH=business_cleansing_output.txt
START_LINE=1001
END_LINE=1501
```

## Usage

Run the script:
```bash
ruby dishwasher.rb
```

The script will:
1. Read business data from the specified Excel file
2. Process corrections for trading names, company numbers, and legal names
3. Handle online marketplace mappings
4. Generate update commands in the output file

## Output

The script generates SQL-like update commands in the specified output file. Each command updates one or more businesses with the same corrections.
