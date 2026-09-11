import argparse
from pathlib import Path
from python_calamine import CalamineWorkbook
from tabulate import tabulate
from xlsx2dict import xlsx2text

def main():
    parser = argparse.ArgumentParser(description="Convert xlsx file to text format.")
    parser.add_argument("xlsx_file", help="Path to the xlsx file.")
    parser.add_argument("-i", "--index", help="Index of the sheet to convert, if not provided, all sheets will be converted.")
    parser.add_argument("-o", "--output", help="Path to the output text file, if not provided, prints to console.")
    parser.add_argument("-c", "--count", action="store_true", help="Print the number of rows and columns for each sheet, if provided, conversion is skipped.")
    args = parser.parse_args()
    if args.count:
        if not Path(args.xlsx_file).is_file():
            raise FileNotFoundError(f"File '{args.xlsx_file}' does not exist.")
        workbook = CalamineWorkbook.from_path(args.xlsx_file)
        rows = []
        for sheet_index, sheet_name in enumerate(workbook.sheet_names, start=1):
            sheet = workbook.get_sheet_by_name(sheet_name)
            rows.append([sheet_index, sheet_name, sheet.height, sheet.width])
        print(tabulate(rows, headers=["Index", "SheetName", "Rows", "Columns"], tablefmt="simple", stralign="left", disable_numparse=True))
    else:
        xlsx2text(args.xlsx_file, index=args.index, output_file=args.output)

if __name__ == "__main__":
    main()