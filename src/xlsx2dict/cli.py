import sys
import argparse
from xlsx2dict import xlsx2text

def main():
    parser = argparse.ArgumentParser(description="Convert xlsx file to text format.")
    parser.add_argument("xlsx_file", help="Path to the xlsx file.")
    parser.add_argument("-i", "--index", help="Index of the sheet to convert, if not provided, all sheets will be converted.")
    parser.add_argument("-o", "--output", help="Path to the output text file, if not provided, prints to console.")
    args = parser.parse_args()
    xlsx2text(args.xlsx_file, index=args.index, output_file=args.output)

if __name__ == "__main__":
    main()