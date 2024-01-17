import argparse
from openpyxl import load_workbook
import pandas as pd
from pathlib import Path
import numpy as np


def get_input_file() -> argparse.Namespace:
    """
    Parse command line arguments

    Returns
    -------
      input file(s) name
    """
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input", "--i", nargs="+", help="filename(s) to parse", required=True
    )
    args = parser.parse_args()

    return args.input


def extract_fields(filename) -> pd.DataFrame:
    """
    Extract data from different sheets of variant workbook

    Parameters
    ----------
      variant workbook file name

    Return
    ------
      data frames from different sheets

    """
    workbook = load_workbook(filename)

    # extract from summary sheet
    sampleID = workbook["summary"]["B1"].value
    CI = workbook["summary"]["F1"].value
    panel = workbook["summary"]["F2"].value
    date = workbook["summary"]["I17"].value
    split_sampleID = sampleID.split("-")
    instrumentID = split_sampleID[0]
    sample_ID = split_sampleID[1]
    batchID = split_sampleID[2]
    testcode = split_sampleID[3]
    sex = split_sampleID[4]
    probesetID = split_sampleID[5]
    ref_genome = workbook["summary"]["B41"].value

    # extract from included sheet
    df = pd.read_excel(filename, sheet_name="included", usecols="A:AT")
    df_included = df[
        ["CHROM", "POS", "REF", "ALT", "HGVSc", "Consequence", "Interpreted", "Comment"]
    ].copy()
    df_included["ref_genome"] = ref_genome
    df_included["instrumentID"] = instrumentID
    df_included["batchID"] = batchID
    df_included["specimenID"] = sample_ID
    df_included["test code"] = testcode
    df_included["sex"] = sex
    df_included["probesetID"] = probesetID
    df_included["CI"] = CI
    df_included["panel"] = panel
    df_included["date"] = date
    df_included["date"] = pd.to_datetime(df_included["date"])
    df_included["Organisation"] = "East Genomic Laboratory Hub"
    df_included["Institution"] = "Cambridge University Hospitals Genomics"

        previous_cell = chr(ord(cell_col[0])-1)
    # put strength as null if it is 'NA'
    for row in range(df_report.shape[0]):
        for column in range(5, df_report.shape[1], 2):
            if df_report.iloc[row, column] == "NA":
                df_report.iloc[row, column] = np.nan

    # removing evidence value if no strength
    for row in range(df_report.shape[0]):
        for column in range(5, df_report.shape[1], 2):
            if df_report.isnull().iloc[row, column]:
                df_report.iloc[row, column + 1] = np.nan

    return df_report, df_included


def main():
    input_file = get_input_file()
    # extract fields from variant workbooks as df
    for filename in input_file:
        df_report, df_included = extract_fields(filename)
        df_final = pd.merge(df_included, df_report, on="HGVSc", how="left")
        df_final.to_csv(Path(filename).stem + ".csv", index=False)
    print("Done")


if __name__ == "__main__":
    main()
