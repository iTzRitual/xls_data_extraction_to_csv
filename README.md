# XLS to CSV Converter
This Python script converts XLS files in the "import_xls" folder to CSV files in the "exported_csv" folder. The script looks for XLS files with a .xls extension and converts each sheet to a separate CSV file with the same name as the original XLS file, but with a .csv extension.

The script extracts specific columns from each sheet and writes them to the CSV file. The specific columns are "Nazwa", "Kwota", "Data płatności", "Faktura", and "Kupujący".

## Requirements
- Python 3.x
- xlrd library

## How to use
1. Place your XLS files in the "import_xls" folder.
2. Run the script.
3. The converted CSV files will be placed in the "exported_csv" folder.
4. If the conversion is successful, a success message will be displayed. Otherwise, an error message will be displayed.

## Note
- The script has been tested on Windows 10 with Python 3.8.8 and xlrd 2.0.1.
- The script will create the "import_xls" and "exported_csv" folders if they do not exist.
- If a file with the same name already exists, the script will add a number to the file name to avoid overwriting the existing file.