import xlsxwriter
from pathlib import Path
from collections import OrderedDict
from python_calamine import CalamineWorkbook
import tabulate


def sheets(xlsx_file):
    """Returns a list of sheet names in the given xlsx file."""
    if not Path(xlsx_file).is_file():
        raise FileNotFoundError(f"File '{xlsx_file}' does not exist.")
    workbook = CalamineWorkbook.from_path(xlsx_file)
    return workbook.sheet_names


def read(xlsx_file:str | Path, sheet_name: str | int, remove_comment: bool = True) -> list[list]:
    """Reads the specified sheet from the given xlsx file and returns its content as a list of lists."""
    if not Path(xlsx_file).is_file():
        raise FileNotFoundError(f"File '{xlsx_file}' does not exist.")
    workbook = CalamineWorkbook.from_path(xlsx_file)
    if isinstance(sheet_name, int):
        sheet_names = workbook.sheet_names
        if sheet_name < 0 or sheet_name >= len(sheet_names):
            raise ValueError(f"Sheet index {sheet_name} is out of range. Available sheets: {sheet_names}")
        sheet_name = sheet_names[sheet_name]
    if sheet_name not in workbook.sheet_names:
        raise ValueError(f"Sheet '{sheet_name}' does not exist in the workbook. Available sheets: {workbook.sheet_names}")
    data = workbook.get_sheet_by_name(sheet_name).to_python(skip_empty_area=False)
    # Remove comment rows if requested, assuming comments are indicated by a leading '#' in the first cell of the row.
    if remove_comment:
        data = [([] if len(cell_value) > 0 and cell_value[0].startswith("#") else cell_value) for cell_value in data]
    # replace '’', '‘', '“', '”' with "'" and '"'
    data = [[cell.replace('’', "'").replace('‘', "'").replace('“', '"').replace('”', '"') if isinstance(cell, str) else cell for cell in row] for row in data]
    # replace str "xx.0" with "xx" for numeric cells
    data = [[cell[:-2] if isinstance(cell, str) and cell.endswith('.0') and cell[:-2].isdigit() else cell for cell in row] for row in data]
    return data


def read_dict(xlsx_file: str | Path, sheet_name: str = None, strip: bool = True, key_map: dict = None, remove_comment: bool = True, save_list: bool = False, head_rows: int = 1) -> OrderedDict:
    """Reads the specified sheet from the given xlsx file and returns its content as an ordered dictionary."""
    data = read(xlsx_file, sheet_name, remove_comment)
    if not data or len(data) <= head_rows:
        return OrderedDict()
    data_dict = OrderedDict()
    # Process header rows to create keys for the dictionary.
    # If head_rows is 1, use the first row as keys.
    # If head_rows > 1, combine the header rows to create keys.
    if head_rows == 1:
        headers = [str(cell_value).strip() for cell_value in data[0]]
    else:
        header_rows = data[0:head_rows].copy()
        for header_row_index in range(head_rows-1, -1, -1):
            header_row = header_rows[header_row_index]
            for col_index, cell_value in enumerate(header_row):
                if col_index == 0:
                    continue
                if not cell_value:
                    if header_row_index > 0 and "".join([header_rows[y][col_index] for y in range(header_row_index)]):
                        continue
                    header_row[col_index] = header_row[col_index-1]
        headers = ["_".join(cell_value) for cell_value in zip(*header_rows)]
    # Apply key mapping if provided
    if key_map:
        headers = [key_map.get(cell_value, cell_value) for cell_value in headers]
    for row_index, row in enumerate(data[head_rows:]):
        row_number = row_index + head_rows + 1
        if not row or all(cell is None for cell in row):
            continue
        # Initialize the dictionary for the current row. If save_list is True, store the entire row under a special key "__list__".
        data_dict[row_number] = {"__list__": row} if save_list else {}
        for header, cell_value in zip(headers, row):
            if not header: # skip columns without header
                continue
            cell_value = str(cell_value).strip() if strip else str(cell_value)
            if header in data_dict[row_number]: # if the key already exists, convert it to a list and append the new value
                if isinstance(data_dict[row_number][header], str):
                    data_dict[row_number][header] = [data_dict[row_number][header]]
                data_dict[row_number][header].append(cell_value)
            else:
                data_dict[row_number][header] = cell_value
    return data_dict


def write(table: list, xlsx_file: str | Path, sheet_name: str = None):
    """Writes the given table (list of lists) to the specified sheet in the xlsx file."""
    workbook = xlsxwriter.Workbook(xlsx_file)
    if not sheet_name:
        sheet_name = "Sheet1"
    worksheet = workbook.add_worksheet(sheet_name)
    for row_index, row in enumerate(table):
        for col_index, cell in enumerate(row):
            worksheet.write(row_index, col_index, cell)
    workbook.close()


def write_dict(xlsx_file: str | Path, data: dict):
    """Writes the given dictionary to the specified xlsx file. The dictionary should be in the format returned by read_dict."""
    # data is expected to be a dictionary:
    # {
    #     sheet1_name: {[[...], ...], ...},
    #     sheet2_name: {[[...], ...], ...},
    #     ...
    #}
    if not data:
        return
    workbook = xlsxwriter.Workbook(xlsx_file)
    for sheet_name, table in data.items():
        worksheet = workbook.add_worksheet(sheet_name)
        for row_index, row in enumerate(table):
            for col_index, cell in enumerate(row):
                worksheet.write(row_index, col_index, cell)
    workbook.close()


def xlsx2text(xlsx_file: str | Path, index: int | str = None, output_file: str | Path = None):
    """Converts the specified sheet from the given xlsx file to a text file. If output_file is not provided, prints the content to the console."""
    sheets_list = sheets(xlsx_file)
    output = ""
    for sheet_index, sheet_name in enumerate(sheets_list):
        if not index or int(index) == sheet_index + 1:
            output += "="*70 + f"\n{sheet_name}\n" + "="*70 + "\n"
            table = read(xlsx_file, sheet_name, remove_comment=False)
            output += tabulate.tabulate(table, tablefmt="pretty", stralign="left")
            output += "\n"
    if output_file:
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(output)
    else:
        print(output)