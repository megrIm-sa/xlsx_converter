# Excel Converter
Excel Converter is a powerful tool designed to convert Excel files (.xlsx) into two different formats: CFG (ConfigFile) and CSV (Comma-separated values).
## Features
#### Excel to CFG Conversion:

- Convert Excel files to CFG format.
- CFG sections are defined by the first column in the Excel file.
- Columns marked with (String) in their name are converted to string format.
- Boolean values (True/False) are converted to lowercase (true/false).
#### Excel to CSV Conversion:

- Convert Excel files to CSV format with a comma (,) as the separator.
- The converted CSV files are saved in UTF-8 format, ensuring compatibility and proper encoding across different platforms.

## How to Use
1. Download and Install:

- Clone this repository or download the ZIP file.
- Extract the contents to a desired location on your system.
- Run the provided .exe file to launch the application.
2. Convert XLSX to CFG:

- Click on the Manual Convert to CFG button.
- Select the Excel file (.xlsx) you want to convert.
- Choose the save location and provide a name for the output .cfg file.
- The program will process the Excel file and generate the corresponding CFG file.
3. Convert XLSX to CSV:

- Click on the Convert XLSX to CSV button.
- Select the Excel file (.xlsx) you want to convert.
- Choose the save location and provide a name for the output .csv file.
- The program will convert the Excel file into a CSV file, ensuring it's saved in UTF-8 format.
## Dependencies
This program is built using Python and relies on the following libraries:

- pandas
- openpyxl
- tkinter
These dependencies are bundled within the .exe file, so no additional installations are required.

## License
This project is licensed under the MIT License. See the LICENSE file for more details.
