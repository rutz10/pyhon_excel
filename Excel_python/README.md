# Excel Text Replacement Tool

This project provides a Python-based tool to perform text replacements in Excel workbooks based on mappings provided in a CSV file. The tool is designed to handle case-sensitive replacements and works on specified sheets and columns of an Excel workbook.

## Features
- Load replacement mappings from a CSV file.
- Perform case-sensitive replacements in Excel cells.
- Process specific sheets and columns in an Excel workbook.
- Save the modified workbook with a new filename.

## File Descriptions

### `find_replace.py`
This script contains the core functionality for text replacement in Excel files. It includes the following functions:

1. **`load_replacements(csv_file)`**:
   - Reads a CSV file containing two columns: `current` (text to find) and `replace` (text to replace with).
   - Returns a dictionary of replacements.

2. **`exact_case_replace(text, find_value, replace_value)`**:
   - Replaces the `find_value` in the given `text` with `replace_value` only if they match exactly.
   - Uses regular expressions to ensure whole-word matching.

3. **`replace_in_excel(sheet, columns, replacements)`**:
   - Iterates through specified columns in an Excel sheet.
   - Applies the replacements to text values in the cells.

4. **`process_workbook(excel_file, replacements, sheet_names, columns)`**:
   - Processes the specified sheets and columns of an Excel workbook.
   - Applies the replacements and saves the modified workbook with a new filename prefixed by `modified_`.

## Additional Script: `search_unreplaced_values.py`

This script is used to identify values in an Excel workbook that do not match any of the replacement mappings provided in a CSV file. It helps ensure that all expected replacements have been applied.

### Functions:

1. **`load_replacements(csv_file)`**:
   - Reads a CSV file containing a column `current`.
   - Returns a dictionary of values to be replaced, storing them for lookup purposes.

2. **`search_unreplaced_values(excel_file, csv_file, sheet_name, columns)`**:
   - Loads the replacement mappings from the CSV file.
   - Checks the specified sheet and columns of the Excel workbook for any values that do not match the replacement mappings.
   - Prints the coordinates and values of any unmatched cells.

## Requirements

To run this project, you need the following Python packages:
- `openpyxl`
- `csv`
- `re`

## Installation

1. Clone this repository or download the files.
2. Install the required Python packages using the following command:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Prepare a CSV file (e.g., `replacements.csv`) with the following columns:
   - `current`: The text to find.
   - `replace`: The text to replace with.

2. Prepare an Excel file (e.g., `workbook.xlsx`) with the data to process.

3. Update the script with the appropriate file paths, sheet names, and columns to process.

4. Run the script:
   ```bash
   python find_replace.py
   ```

5. The modified Excel file will be saved with a new filename prefixed by `modified_`.

### Usage:

1. Prepare a CSV file (e.g., `replacements.csv`) with the column `current` containing the values to match.
2. Prepare an Excel file (e.g., `workbook.xlsx`) with the data to check.
3. Update the script with the appropriate file paths, sheet name, and columns to process.
4. Run the script:
   ```bash
   python search_unreplaced_values.py
   ```
5. The script will output any unmatched values along with their cell coordinates, or confirm that all values matched.

## Example

### CSV File (`replacements.csv`):
| current | replace |
|---------|---------|
| old     | new     |
| test    | example |

### Excel File (`workbook.xlsx`):
| Column A | Column B |
|----------|----------|
| old      | test     |
| data     | value    |

### Result (`modified_workbook.xlsx`):
| Column A | Column B |
|----------|----------|
| new      | example  |
| data     | value    |

## Recent Updates

### Enhanced Replacement Logic
- The `exact_case_replace` function now ensures that replacements are applied only when the target word is followed by non-alphanumeric characters or the end of the string. Cases where the word is followed by numbers or other alphanumeric characters are ignored.

### Enhanced Unreplaced Value Search
- The `search_unreplaced_values` function now uses a similar logic to identify unreplaced values. It matches exact words followed by non-alphanumeric characters or the end of the string, while ignoring cases where the word is followed by numbers or other alphanumeric characters.

### Example Scenarios
#### Replacement Example:
- **Input Text:** `Apple.`, `Apple.Car`, `Apple1`
- **Replacement Rule:** Replace `Apple` with `Orange`
- **Result:** `Orange.`, `Orange.Car`, `Apple1` (unchanged because it is followed by a number)

#### Unreplaced Value Search Example:
- **Input Text in Excel:** `Apple.`, `Apple1`
- **CSV Mapping:** Contains `Apple`
- **Result:** `Apple.` is considered replaced, but `Apple1` is flagged as unreplaced.

## License
This project is licensed under the MIT License.