from pathlib import Path
import sys

from openpyxl import load_workbook


def main() -> int:
    file_name = input("Enter the Excel file name: ").strip()

    if not file_name:
        print("Invalid file name.")
        return 1

    try:
        workbook = load_workbook(filename=Path(file_name), read_only=True, data_only=True)
    except FileNotFoundError:
        print("File not found.")
        return 1
    except Exception as exc:
        print(f"Error opening the file: {exc}")
        return 1

    try:
        if not workbook.worksheets:
            print("The Excel file does not contain readable worksheets.")
            return 1

        sheet = workbook.worksheets[0]
        print(sheet.cell(row=1, column=1).value)
        return 0
    except Exception as exc:
        print(f"Error reading the file: {exc}")
        return 1
    finally:
        workbook.close()


if __name__ == "__main__":
    sys.exit(main())
