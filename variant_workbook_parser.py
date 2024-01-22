import argparse
from typing import Union
import re
from pathlib import Path
import numpy as np
from openpyxl import load_workbook
import pandas as pd


def get_command_line_args() -> argparse.Namespace:
    """
    Parse command line arguments

    Returns
    -------
    args : Namespace
        Namespace of command line argument inputs
    """
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input", "--i", nargs="+", help="input file name(s) to parse",
        required=True
    )
    parser.add_argument(
        "--outdir", "--o", help="dir to save output(s)", default="./"
    )
    parser.add_argument(
        "--unusual_sample_name", action="store_true",
        help="add this argument if sample name is unusual",
    )
    args = parser.parse_args()

    return args


def get_summary_fields(filename: str, unusual_sample_name: bool) \
                       -> Union[pd.DataFrame, str]:
    """
    Extract data from summary sheet of variant workbook

    Parameters
    ----------
      variant workbook file name
      boolean for unusual_sample_name

    Returns
    -------
      data frame from summary sheet
      str: Pass OR Fail
    """
    workbook = load_workbook(filename)
    sampleID = workbook["summary"]["B1"].value
    CI = workbook["summary"]["F1"].value
    panel = workbook["summary"]["F2"].value
    date = workbook["summary"]["I17"].value
    ref_genome = workbook["summary"]["B41"].value
    split_sampleID = sampleID.split("-")
    instrumentID = split_sampleID[0]
    sample_ID = split_sampleID[1]
    batchID = split_sampleID[2]
    testcode = split_sampleID[3]
    sex = split_sampleID[4]
    probesetID = split_sampleID[5]

    # checking sample naming
    if not unusual_sample_name:
        check_naming = check_sample_name(instrumentID, sample_ID,
                                         batchID, testcode, sex,
                                         probesetID, filename)
    d = {"instrumentID": instrumentID,
         "specimenID": sample_ID,
         "batchID": batchID,
         "test code": testcode,
         "probesetID": probesetID,
         "CI": CI,
         "panel": panel,
         "ref_genome": ref_genome,
         "date": date}
    df_summary = pd.DataFrame([d])
    df_summary['date'] = pd.to_datetime(df_summary['date'])
    df_summary["Organisation"] = "East Genomic Laboratory Hub"
    df_summary["Institution"] = "Cambridge University Hospitals Genomics"

    return df_summary, check_naming


def get_included_fields(filename: str) -> pd.DataFrame:
    """
    Extract data from included sheet of variant workbook

    Parameters
    ----------
      variant workbook file name

    Return
    ------
      data frame from included sheet
    """
    workbook = load_workbook(filename)
    num_variants = workbook['summary']['C34'].value
    df = pd.read_excel(filename, sheet_name="included", usecols="A:AT",
                       nrows=num_variants)
    df_included = df[["CHROM", "POS", "REF", "ALT", "HGVSc", "Consequence",
                      "Interpreted", "Comment"]].copy()
    return df_included


def get_report_fields(filename: str) -> pd.DataFrame:
    """
    Extract data from report sheet(s) of variant workbook

    Parameters
    ----------
      variant workbook file name

    Return
    ------
      data frame from report sheet(s)

    """
    workbook = load_workbook(filename)
    field_cells = [
        ("Associated disease", "C5"),
        ("Known inheritance", "C6"),
        ("Prevalence", "C7"),
        ("HGVSc", "C3"),
        ("Final Classification", "C26"),
        ("PVS1", "H9"),
        ("PVS1_evidence", "C9"),
        ("PS1", "H10"),
        ("PS1_evidence", "C10"),
        ("PS2", "H11"),
        ("PS2_evidence", "C11"),
        ("PS3", "H12"),
        ("PS3_evidence", "C12"),
        ("PS4", "H13"),
        ("PS4_evidence", "C13"),
        ("PM1", "H14"),
        ("PM1_evidence", "C14"),
        ("PM2", "H15"),
        ("PM2_evidence", "C15"),
        ("PM3", "H16"),
        ("PM3_evidence", "C16"),
        ("PM4", "H17"),
        ("PM4_evidence", "C17"),
        ("PM5", "H18"),
        ("PM5_evidence", "C18"),
        ("PM6", "H19"),
        ("PM6_evidence", "C19"),
        ("PP1", "H20"),
        ("PP1_evidence", "C20"),
        ("PP2", "H21"),
        ("PP2_evidence", "C21"),
        ("PP3", "H22"),
        ("PP3_evidence", "C22"),
        ("PP4", "H23"),
        ("PP4_evidence", "C23"),
        ("BS1", "K8"),
        ("BS1_evidence", "C8"),
        ("BS2", "K11"),
        ("BS2_evidence", "C11"),
        ("BS3", "K12"),
        ("BS3_evidence", "C12"),
        ("BA1", "K15"),
        ("BA1_evidence", "C15"),
        ("BP2", "K16"),
        ("BP2_evidence", "C16"),
        ("BP3", "K17"),
        ("BP3_evidence", "C17"),
        ("BS4", "K20"),
        ("BS4_evidence", "C20"),
        ("BP1", "K21"),
        ("BP1_evidence", "C21"),
        ("BP4", "K22"),
        ("BP4_evidence", "C22"),
        ("BP5", "K23"),
        ("BP5_evidence", "C23"),
        ("BP7", "K24"),
        ("BP7_evidence", "C24"),
    ]
    col_name = [i[0] for i in field_cells]
    df_report = pd.DataFrame(columns=col_name)
    report_sheets = [
        idx for idx in workbook.sheetnames if idx.lower().startswith("report")
    ]

    for idx, sheet in enumerate(report_sheets):
        for field, cell in field_cells:
            df_report.loc[idx, field] = workbook[sheet][cell].value

    # put strength as nan if it is 'NA'
    for row in range(df_report.shape[0]):
        for column in range(5, df_report.shape[1], 2):
            if df_report.iloc[row, column] == "NA":
                df_report.iloc[row, column] = np.nan

    # removing evidence value if no strength
    for row in range(df_report.shape[0]):
        for column in range(5, df_report.shape[1], 2):
            if df_report.isnull().iloc[row, column]:
                df_report.iloc[row, column + 1] = np.nan

    return df_report


def check_sample_name(instrumentID: str, sample_ID: str, batchID: str,
                      testcode: str, sex: str, probesetID: str,
                      filename: str) -> str:
    """
    checking if individual parts of sample name have
    expected naming format

    Parameters
    ----------
      str values for instrumentID, sample_ID, batchID, testcode,
      sex, probesetID

    Return
    ------
      str: Pass or Fail
    """
    print("Checking the naming format")
    try:
        assert re.match(r"^\d{9}$", instrumentID), (f"Unusual name for instrumentID in {filename}")
        assert re.match(r"^\d{5}[A-Z]\d{4}$", sample_ID), f"Unusual sampleID in {filename}"
        assert re.match(r"^\d{2}[A-Z]{5}\d{1,}$", batchID), f"Unusual batchID in {filename}"
        assert re.match(r"^\d{4}$", testcode), f"Unusual testcode in {filename}"
        assert re.match(r"^[A-Z]$", sex), f"Unusual sex naming in {filename}"
        assert 0 < len(probesetID) < 20, f"probesetID in {filename} is too long/short"
        assert probesetID.isalnum() and not probesetID.isalpha(), f"Unusual probesetID in {filename}"
        check_naming = "Pass"
    except AssertionError as msg:
        check_naming = "Fail"
        print(msg)

    return check_naming


def checking_sheets(filename) -> str:
    """
    check if extra row(s)/col(s) are added in the sheets

    Parameters
    ----------
      variant workbook file name

    Return
    ------
      str Pass OR Fail
    """
    workbook = load_workbook(filename)
    summary = workbook['summary']
    included = workbook['included']
    reports = [idx for idx in workbook.sheetnames if idx.lower().startswith("report")]
    try:
        assert summary["A51"].value == "Report Job ID:", f"extra row(s) added in summary of {filename}"
        assert summary["I16"].value == "Date", f"extra col(s) added in summary of {filename}"
        assert included['AT1'].value == "Interpreted", f"extra col(s) added in included of {filename}"
        assert included.max_row == summary["C34"].value+1, f"extra row(s) added in included of {filename}"
        for sheet in reports:
            report = workbook[sheet]
            assert report["B26"].value == "Final Classification", f"extra row(s) added in {report.title} of {filename}"
            assert report["L4"].value == "B_POINTS", f"extra col(s) added in {report.title} of {filename}"
        check_sheets = "Pass"
    except AssertionError as msg:
        check_sheets = "Fail"
        print(msg)

    return check_sheets


def write_txt_file(filename):
    """
    write txt file output

    Parameters
    ----------
      variant workbook file name
    """
    with open('workbooks_fail_to_parse.txt', 'a') as file:
        file.write(filename+"\n")
        file.close()


def main():
    arguments = get_command_line_args()
    input_file = arguments.input
    unusual_sample_name = arguments.unusual_sample_name

    # extract fields from variant workbooks as df and merged
    for filename in input_file:
        check_sheets = checking_sheets(filename)
        if check_sheets == "Pass":
            df_summary, checking_name = get_summary_fields(filename,
                                                           unusual_sample_name)
            if checking_name == "Pass":
                df_included = get_included_fields(filename)
                df_report = get_report_fields(filename)
                df_merged = pd.merge(df_included, df_summary, how="cross")
                df_final = pd.merge(df_merged, df_report, on="HGVSc",
                                    how="left")
                df_final.to_csv(arguments.outdir + Path(filename).stem +
                                ".csv", index=False)
            else:
                write_txt_file(filename)
        else:
            write_txt_file(filename)
    print("Done")


if __name__ == "__main__":
    main()
