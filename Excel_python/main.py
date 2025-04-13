import sys
from find_replace import process_workbook, load_replacements
from search_unreplaced_values import search_unreplaced_values

def main():
    # Specify the file paths
    excel_file = 'test.xlsx'  # Excel file to process
    csv_file = 'replacements.csv'  # CSV file with current and replace columns
    sheet_names = ['Sheet1']  # List of sheet names to process
    columns = [1]  # Columns to check (e.g., column 1, 2 for columns A, B)

    # 1. Perform the find and replace operation
    print("Starting find and replace...")
    replacements = load_replacements(csv_file)
    process_workbook(excel_file, replacements, sheet_names, columns)
    print(f"Workbook saved as 'modified_{excel_file}'")

    # 2. Search for any unreplaced values
    print("\nSearching for unreplaced values...")
    for sheet_name in sheet_names:
        search_unreplaced_values(excel_file, csv_file, sheet_name, columns)

if __name__ == "__main__":
    main()