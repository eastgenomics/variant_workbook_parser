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

    # extract from report sheet
    df_report = pd.DataFrame(
        columns=[
            "Associated_disease",
            "Known_inheritance",
            "Prevalence",
            "HGVSc",
            "Final_classification",
            "PVS1",
            "PVS1_evidence",
            "PS1",
            "PS1_evidence",
            "PS2",
            "PS2_evidence",
            "PS3",
            "PS3_evidence",
            "PS4",
            "PS4_evidence",
            "PM1",
            "PM1_evidence",
            "PM2",
            "PM2_evidence",
            "PM3",
            "PM3_evidence",
            "PM4",
            "PM4_evidence",
            "PM5",
            "PM5_evidence",
            "PM6",
            "PM6_evidence",
            "PP1",
            "PP1_evidence",
            "PP2",
            "PP2_evidence",
            "PP3",
            "PP3_evidence",
            "PP4",
            "PP4_evidence",
            "PP5",
            "PP5_evidence",
            "BS1",
            "BS1_evidence",
            "BS2",
            "BS2_evidence",
            "BS3",
            "BS3_evidence",
            "BA1",
            "BA1_evidence",
            "BP2",
            "BP2_evidence",
            "BP3",
            "BP3_evidence",
            "BS4",
            "BS4_evidence",
            "BP1",
            "BP1_evidence",
            "BP4",
            "BP4_evidence",
            "BP5",
            "BP5_evidence",
            "BP6",
            "BP6_evidence",
            "BP7",
            "BP7_evidence",
        ]
    )
    for sheet in range(1, 4):  # hard coded for 3 sheets, can put arg if wanted
        df_report.loc[sheet, "Associated_disease"] = workbook[f"report_{sheet}"][
            "C5"
        ].value
        df_report.loc[sheet, "Known_inheritance"] = workbook[f"report_{sheet}"][
            "C6"
        ].value
        df_report.loc[sheet, "Prevalence"] = workbook[f"report_{sheet}"]["C7"].value
        df_report.loc[sheet, "HGVSc"] = workbook[f"report_{sheet}"]["C3"].value
        df_report.loc[sheet, "Final_classification"] = workbook[f"report_{sheet}"][
            "C27"
        ].value
        df_report.loc[sheet, "PVS1"] = workbook[f"report_{sheet}"]["H9"].value
        df_report.loc[sheet, "PVS1_evidence"] = workbook[f"report_{sheet}"]["C9"].value
        df_report.loc[sheet, "PS1"] = workbook[f"report_{sheet}"]["H10"].value
        df_report.loc[sheet, "PS1_evidence"] = workbook[f"report_{sheet}"]["C10"].value
        df_report.loc[sheet, "PS2"] = workbook[f"report_{sheet}"]["H11"].value
        df_report.loc[sheet, "PS2_evidence"] = workbook[f"report_{sheet}"]["C11"].value
        df_report.loc[sheet, "PS3"] = workbook[f"report_{sheet}"]["H12"].value
        df_report.loc[sheet, "PS3_evidence"] = workbook[f"report_{sheet}"]["C12"].value
        df_report.loc[sheet, "PS4"] = workbook[f"report_{sheet}"]["H13"].value
        df_report.loc[sheet, "PS4_evidence"] = workbook[f"report_{sheet}"]["C13"].value
        df_report.loc[sheet, "PM1"] = workbook[f"report_{sheet}"]["H14"].value
        df_report.loc[sheet, "PM1_evidence"] = workbook[f"report_{sheet}"]["C14"].value
        df_report.loc[sheet, "PM2"] = workbook[f"report_{sheet}"]["H15"].value
        df_report.loc[sheet, "PM2_evidence"] = workbook[f"report_{sheet}"]["C15"].value
        df_report.loc[sheet, "PM3"] = workbook[f"report_{sheet}"]["H16"].value
        df_report.loc[sheet, "PM3_evidence"] = workbook[f"report_{sheet}"]["C16"].value
        df_report.loc[sheet, "PM4"] = workbook[f"report_{sheet}"]["H17"].value
        df_report.loc[sheet, "PM4_evidence"] = workbook[f"report_{sheet}"]["C17"].value
        df_report.loc[sheet, "PM5"] = workbook[f"report_{sheet}"]["H18"].value
        df_report.loc[sheet, "PM5_evidence"] = workbook[f"report_{sheet}"]["C18"].value
        df_report.loc[sheet, "PM6"] = workbook[f"report_{sheet}"]["H19"].value
        df_report.loc[sheet, "PM6_evidence"] = workbook[f"report_{sheet}"]["C19"].value
        df_report.loc[sheet, "PP1"] = workbook[f"report_{sheet}"]["H20"].value
        df_report.loc[sheet, "PP1_evidence"] = workbook[f"report_{sheet}"]["C20"].value
        df_report.loc[sheet, "PP2"] = workbook[f"report_{sheet}"]["H21"].value
        df_report.loc[sheet, "PP2_evidence"] = workbook[f"report_{sheet}"]["C21"].value
        df_report.loc[sheet, "PP3"] = workbook[f"report_{sheet}"]["H22"].value
        df_report.loc[sheet, "PP3_evidence"] = workbook[f"report_{sheet}"]["C22"].value
        df_report.loc[sheet, "PP4"] = workbook[f"report_{sheet}"]["H23"].value
        df_report.loc[sheet, "PP4_evidence"] = workbook[f"report_{sheet}"]["C23"].value
        df_report.loc[sheet, "PP5"] = workbook[f"report_{sheet}"]["H24"].value
        df_report.loc[sheet, "PP5_evidence"] = workbook[f"report_{sheet}"]["C24"].value
        df_report.loc[sheet, "BS1"] = workbook[f"report_{sheet}"]["J8"].value
        df_report.loc[sheet, "BS1_evidence"] = workbook[f"report_{sheet}"]["C8"].value
        df_report.loc[sheet, "BS2"] = workbook[f"report_{sheet}"]["J11"].value
        df_report.loc[sheet, "BS2_evidence"] = workbook[f"report_{sheet}"]["C11"].value
        df_report.loc[sheet, "BS3"] = workbook[f"report_{sheet}"]["J12"].value
        df_report.loc[sheet, "BS3_evidence"] = workbook[f"report_{sheet}"]["C12"].value
        df_report.loc[sheet, "BA1"] = workbook[f"report_{sheet}"]["J15"].value
        df_report.loc[sheet, "BA1_evidence"] = workbook[f"report_{sheet}"]["C15"].value
        df_report.loc[sheet, "BP2"] = workbook[f"report_{sheet}"]["J16"].value
        df_report.loc[sheet, "BP2_evidence"] = workbook[f"report_{sheet}"]["C16"].value
        df_report.loc[sheet, "BP3"] = workbook[f"report_{sheet}"]["J17"].value
        df_report.loc[sheet, "BP3_evidence"] = workbook[f"report_{sheet}"]["C17"].value
        df_report.loc[sheet, "BS4"] = workbook[f"report_{sheet}"]["J20"].value
        df_report.loc[sheet, "BS4_evidence"] = workbook[f"report_{sheet}"]["C20"].value
        df_report.loc[sheet, "BP1"] = workbook[f"report_{sheet}"]["J21"].value
        df_report.loc[sheet, "BP1_evidence"] = workbook[f"report_{sheet}"]["C21"].value
        df_report.loc[sheet, "BP4"] = workbook[f"report_{sheet}"]["J22"].value
        df_report.loc[sheet, "BP4_evidence"] = workbook[f"report_{sheet}"]["C22"].value
        df_report.loc[sheet, "BP5"] = workbook[f"report_{sheet}"]["J23"].value
        df_report.loc[sheet, "BP5_evidence"] = workbook[f"report_{sheet}"]["C23"].value
        df_report.loc[sheet, "BP6"] = workbook[f"report_{sheet}"]["J24"].value
        df_report.loc[sheet, "BP6_evidence"] = workbook[f"report_{sheet}"]["C24"].value
        df_report.loc[sheet, "BP7"] = workbook[f"report_{sheet}"]["J25"].value
        df_report.loc[sheet, "BP7_evidence"] = workbook[f"report_{sheet}"]["C25"].value

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
