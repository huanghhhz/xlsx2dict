# xlsx2dict

A Python library for reading and writing Excel files, converting between xlsx and dict or text.

## Installation

```bash
pip install xlsx2dict
```

## Usage

### Python API

```python
from xlsx2dict import sheets, read, read_dict, write, write_dict, xlsx2text

# List sheet names
sheet_names = sheets("example.xlsx")

# Read a sheet as list of lists
data = read("example.xlsx", sheet_name="Sheet1")

# Read a sheet as ordered dict (keyed by row number)
d = read_dict("example.xlsx", sheet_name="Sheet1")
d = read_dict("example.xlsx", sheet_name="Sheet1", head_rows=2)       # multi-row headers
d = read_dict("example.xlsx", sheet_name="Sheet1", save_list=True)    # include raw row as __list__
d = read_dict("example.xlsx", sheet_name="Sheet1", key_map={"Old": "New"})  # rename columns

# Write list of lists to xlsx
write([["Name", "Age"], ["Alice", 30]], "output.xlsx", sheet_name="Data")

# Write dict to multi-sheet xlsx
write_dict("output.xlsx", {
    "Sheet1": [["Name", "Age"], ["Alice", 30]],
    "Sheet2": [["City", "Country"], ["Paris", "France"]],
})

# Convert xlsx to formatted text
xlsx2text("example.xlsx")                    # prints all sheets
xlsx2text("example.xlsx", index=1)           # prints first sheet
xlsx2text("example.xlsx", output_file="out.txt")  # write to file
```

### CLI

```bash
# Convert all sheets to console
xlsx2text example.xlsx

# Convert a specific sheet
xlsx2text example.xlsx -i 1

# Save output to file
xlsx2text example.xlsx -o output.txt
```

## License

MIT
