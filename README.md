# variant_workbook_parser

## What does this app do?

This script parses the sheets in variant workbook(s) and extract fields required to submit to Variant Database.

## What are typical use cases for this app?

This script may be executed as a standalone to parse the variant workbook(s).

## What data are required for this app to run?

**Packages**

* Python packages (specified in requirements.txt)

**File inputs (required)**:

- `--input` / `--i`: variant workbook spreadsheet(s)

**Other Inputs (optional):**

`--outdir` / `--o`: dir where to save output file(s) \
`--unusual_sample_name`: boolean - default is False 

## Command line to run 
```python variant_workbook_parser.py --i <sample_1.xlsx sample_2.xlsx ... sample_n.xlsx>  --o </path/to/folder/> --unusual_sample_name```
