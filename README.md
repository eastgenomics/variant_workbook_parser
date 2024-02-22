# variant_workbook_parser.py

## What does this script do?

This script parses the sheets in variant workbook(s) and extract fields required to submit to Variant Database and Clinvar.

## What are typical use cases for this script?

This script may be executed as a standalone to parse the variant workbook(s).

## What data are required for this script to run?

**Packages**

* Python packages (specified in requirements.txt)

**File inputs (required)**:

- `--indir` / `--i`: directory for input file(s)

**Other Inputs (optional):**

- `--outdir` / `--o`: dir where the output csv files are saved  
- `--file` / `--f` : workbook if want to specify; if not specify, the script will take all xlxs file in the `--indir`. 
- `--logdir` / `--ld` : dir where the log file(s) are saved. 
- `--completed_dir` / `--cd` : dir to where the successfully parsed workbook(s) are moved. 
- `--unusual_sample_name`: boolean - default is False and the sample name in the workbook will be tested if it follows the standard naming format, and if the test fails, the workbook for that sample will not be parsed. Put this args to skip the test in samples with unusual naming format.

## What outputs are expected from this app?
- csv file containing all variants from the workbook
- csv file containing interpreted variant from the workbook
- workbooks_fail_to_parse.txt (optional) - txt file containing the file(s) that fails to be parsed by parser script and reason for fail
- workbooks_parsed_all_variants.txt (optional) - txt file containing the file(s) that are successfully parsed 
- workbooks_parsed_clinvar_variants.txt (optional) - txt file containing the file(s) that are successfully parsed for clinvar submission


## Command line to run 
```python variant_workbook_parser.py --i </path/to/folder/> --f <sample_name> --o </path/to/folder/> --ld </path/to/folder/> --cd  </path/to/folder/> --unusual_sample_name```

# get_completed_wb.py

## What does this script do?

This script searches file(s) for given sample(s) in clingen folder of Trust PC and copies these file(s) into another folder.

## What data are required for this script to run?

**File inputs (required)**:

- `--input` / `--i`: txt file containing a list of samples for verified workbooks 
- `--outdir` / `--o`: dir where to copy the verified workbooks 
- `--folder` / `--f`: dir where to search the verified workbooks
- `--logdir` / `--ld` : dir where the log file(s) is saved
## What outputs are expected from this app?
- found verified workbooks are copied into outdir
- workbooks_not_found_clingen.txt- txt file containing the samples that are not found

## Command line to run 
```python get_completed_wb.py --i <txt_file_name>  --o </path/to/folder/> --f </path/to/folder> --ld </path/to/folder>```