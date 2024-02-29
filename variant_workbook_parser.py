import argparse
import re
import os
import sys
import glob
import shutil
from pathlib import Path
import time
from datetime import datetime
import uuid
import json
import numpy as np
from openpyxl import load_workbook
import pandas as pd

PARSED_FILE = "workbooks_parsed_all_variants.txt"
CLINVAR_FILE = "workbooks_parsed_clinvar_variants.txt"
FAILED_FILE = "workbooks_fail_to_parse.txt"


def get_command_line_args(arguments) -> argparse.Namespace:
    """
    Parse command line arguments

    Returns
    -------
    args : Namespace
        Namespace of command line argument inputs
    """
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--indir", "--i", help="input dir of file(s) to parse",
        required=True
    )
    parser.add_argument(
        "--file", "--f", nargs='+', help="input file(s) to parse if want to specify"
    )
    parser.add_argument(
        "--outdir", "--o", help="dir to save the output csv files", default="./"
    )
    parser.add_argument(
        "--logdir", "--ld", help="dir to save log txt files", default="./"
    )
    parser.add_argument(
        "--completed_dir", "--cd", help="dir to move the successfully parsed workbooks", default="./"
    )
    parser.add_argument(
        "--unusual_sample_name", action="store_true",
        help="add this argument if sample name is unusual",
    )
    args = parser.parse_args(arguments)

    return args


def get_summary_fields(filename: str, config_variable: dict,
                       unusual_sample_name: bool) \
                       -> tuple[pd.DataFrame, str]:
    """
    Extract data from summary sheet of variant workbook

    Parameters
    ----------
      variant workbook file name
      dict from config file
      boolean for unusual_sample_name

    Returns
    -------
      data frame from summary sheet
      str for error message
    """
    workbook = load_workbook(filename)
    sampleID = workbook["summary"]["B1"].value
    CI = workbook["summary"]["F1"].value
    if ";" in CI:
        split_CI = CI.split(";")
        indication = []
        for each in split_CI:
            remove_R = each.split("_")[1]
            indication.append(remove_R)
        new_CI = ";".join(indication)
    else:
        new_CI = CI.split("_")[1]
    panel = workbook["summary"]["F2"].value
    date = workbook["summary"]["I17"].value
    split_sampleID = sampleID.split("-")
    instrumentID = split_sampleID[0]
    sample_ID = split_sampleID[1]
    batchID = split_sampleID[2]
    testcode = split_sampleID[3]
    probesetID = split_sampleID[5]
    ref_genome = "not_defined"
    for cell in workbook["summary"]["A"]:
        if cell.value == "Reference:":
            ref_genome = workbook["summary"][f"B{cell.row}"].value

    # checking sample naming
    error_msg = None
    if not unusual_sample_name:
        error_msg = check_sample_name(instrumentID, sample_ID,
                                      batchID, testcode,
                                      probesetID)
    d = {"Instrument ID": instrumentID,
         "Specimen ID": sample_ID,
         "Batch ID": batchID,
         "Test code": testcode,
         "Probeset ID": probesetID,
         "Preferred condition name": new_CI,
         "Panel": panel,
         "Ref genome": ref_genome,
         "Date last evaluated": date}
    df_summary = pd.DataFrame([d])
    df_summary['Date last evaluated'] = pd.to_datetime(df_summary
                                                       ['Date last evaluated'])
    df_summary["Organisation"] = config_variable["info"]["Organisation"]
    df_summary["Institution"] = config_variable["info"]["Institution"]
    df_summary["Collection method"] = config_variable["info"] \
                                      ["Collection method"]
    df_summary["Allele origin"] = config_variable["info"]["Allele origin"]
    df_summary["Affected status"] = config_variable["info"]["Affected status"]

    # getting the folder name of workbook
    # the folder name should return designated folder for either CUH or NUH
    folder_name = get_folder(filename)
    if folder_name == config_variable["info"]["CUH folder"]:
        df_summary["Organisation ID"] = config_variable["info"]["CUH org ID"]
    elif folder_name == config_variable["info"]["NUH folder"]:
        df_summary["Organisation ID"] = config_variable["info"]["NUH org ID"]
    else:
        print("Running for the wrong folder")
        sys.exit(1)

    return df_summary, error_msg


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
    num_variants = workbook['summary']['C34'].value #TO DO: change to 28
    interpreted_col = get_col_letter(workbook["included"], "Interpreted")
    df = pd.read_excel(filename, sheet_name="included",
                       usecols=f"A:{interpreted_col}",
                       nrows=num_variants)
    df_included = df[["CHROM", "POS", "REF", "ALT", "SYMBOL", "HGVSc",
                      "Consequence", "Interpreted", "Comment"]].copy()
    if len(df_included["Interpreted"].value_counts()) > 0:
        df_included["Interpreted"] = df_included["Interpreted"].str.lower()
    df_included.rename(columns={"CHROM": "Chromosome", "SYMBOL": "Gene symbol",
                                "POS": "Start", "REF": "Reference allele",
                                "ALT": "Alternate allele"},
                       inplace=True)
    df_included['Local ID'] = ""
    for row in range(df_included.shape[0]):
        unique_id = uuid.uuid1()
        df_included.loc[row, "Local ID"] = f"uid_{unique_id.time}"
        time.sleep(0.5)
    df_included["Linking ID"] = df_included["Local ID"]

    return df_included


def get_report_fields(filename: str, df_included: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    """
    Extract data from interpret sheet(s) of variant workbook

    Parameters
    ----------
      variant workbook file name
      data frame from included sheet

    Return
    ------
      data frame from interpret sheet(s)
      str for error message

    """
    workbook = load_workbook(filename)
    field_cells = [
        ("Associated disease", "C4"),
        ("Known inheritance", "C5"),
        ("Prevalence", "C6"),
        ("HGVSc", "C3"),
        ("Germline classification", "C26"),
        ("PVS1", "H10"),
        ("PVS1_evidence", "C10"),
        ("PS1", "H11"),
        ("PS1_evidence", "C11"),
        ("PS2", "H12"),
        ("PS2_evidence", "C12"),
        ("PS3", "H13"),
        ("PS3_evidence", "C13"),
        ("PS4", "H14"),
        ("PS4_evidence", "C14"),
        ("PM1", "H15"),
        ("PM1_evidence", "C15"),
        ("PM2", "H16"),
        ("PM2_evidence", "C16"),
        ("PM3", "H17"),
        ("PM3_evidence", "C17"),
        ("PM4", "H18"),
        ("PM4_evidence", "C18"),
        ("PM5", "H19"),
        ("PM5_evidence", "C19"),
        ("PM6", "H20"),
        ("PM6_evidence", "C20"),
        ("PP1", "H21"),
        ("PP1_evidence", "C21"),
        ("PP2", "H22"),
        ("PP2_evidence", "C22"),
        ("PP3", "H23"),
        ("PP3_evidence", "C23"),
        ("PP4", "H24"),
        ("PP4_evidence", "C24"),
        ("BS1", "K9"),
        ("BS1_evidence", "C9"),
        ("BS2", "K12"),
        ("BS2_evidence", "C12"),
        ("BS3", "K13"),
        ("BS3_evidence", "C13"),
        ("BA1", "K16"),
        ("BA1_evidence", "C16"),
        ("BP2", "K17"),
        ("BP2_evidence", "C17"),
        ("BP3", "K18"),
        ("BP3_evidence", "C18"),
        ("BS4", "K21"),
        ("BS4_evidence", "C21"),
        ("BP1", "K22"),
        ("BP1_evidence", "C22"),
        ("BP4", "K23"),
        ("BP4_evidence", "C23"),
        ("BP5", "K24"),
        ("BP5_evidence", "C24"),
        ("BP7", "K25"),
        ("BP7_evidence", "C25"),
    ]
    col_name = [i[0] for i in field_cells]
    df_report = pd.DataFrame(columns=col_name)
    report_sheets = [
        idx for idx in workbook.sheetnames if idx.lower().startswith("interpret")
    ]

    for idx, sheet in enumerate(report_sheets):
        for field, cell in field_cells:
            if workbook[sheet][cell].value is not None:
                df_report.loc[idx, field] = workbook[sheet][cell].value
    df_report.reset_index(drop=True, inplace=True)
    error_msg = None
    if not df_report.empty:
        error_msg = check_interpret_table(df_report, df_included)
    if not error_msg:
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

        # getting comment on classification for clinvar submission
        matched_strength = [("PVS", "Very Strong"),
                            ("PS", "Strong"),
                            ("PM", "Moderate"),
                            ("PP", "Supporting"),
                            ("BS", "Supporting"),
                            ("BA", "Stand-Alone"),
                            ("BP", "Supporting")
                            ]
        df_report["Comment on classification"] = ""
        for row in range(df_report.shape[0]):
            evidence = []
            for column in range(5, df_report.shape[1]-1, 2):
                if not df_report.isnull().iloc[row, column]:
                    evidence.append([df_report.columns[column],
                                     df_report.iloc[row, column]])
            for index, value in enumerate(evidence):
                for st1, st2 in matched_strength:
                    if st1 in value[0] and st2 == value[1]:
                        evidence[index][1] = ""
            evidence_pair = []
            for e in evidence:
                evidence_pair.append('_'.join(e).rstrip("_"))
            comment_on_classification = ','.join(evidence_pair)
            df_report.iloc[row, df_report.columns.get_loc('Comment on classification')] = comment_on_classification

    return df_report, error_msg


def check_sample_name(instrumentID: str, sample_ID: str, batchID: str,
                      testcode: str, probesetID: str) -> str:
    """
    checking if individual parts of sample name have
    expected naming format

    Parameters
    ----------
      str values for instrumentID, sample_ID, batchID, testcode,
      probesetID

    Return
    ------
      str for error message
    """
    try:
        assert re.match(r"^\d{9}$", instrumentID), \
        "Unusual name for instrumentID"
        assert re.match(r"^\d{5}[A-Z]\d{4}$", sample_ID), "Unusual sampleID"
        assert re.match(r"^\d{2}[A-Z]{5}\d{1,}$", batchID), "Unusual batchID"
        assert re.match(r"^\d{4}$", testcode), "Unusual testcode"
        assert 0 < len(probesetID) < 20, "probesetID is too long/short"
        assert probesetID.isalnum() and not probesetID.isalpha(), \
        "Unusual probesetID"
        error_msg = None
    except AssertionError as msg:
        error_msg = str(msg)
        print(msg)

    return error_msg


def checking_sheets(filename: str) -> str:
    """
    check if extra row(s)/col(s) are added in the sheets

    Parameters
    ----------
      variant workbook file name

    Return
    ------
      str for error message
    """
    workbook = load_workbook(filename)
    summary = workbook["summary"]
    reports = [idx for idx in workbook.sheetnames if idx.lower().startswith("interpret")]
    try:
        assert summary["I16"].value == "Date", \
        "extra col(s) added or change(s) done in summary sheet"
        for sheet in reports:
            report = workbook[sheet]
            assert report["B26"].value == "FINAL ACMG CLASSIFICATION", \
            f"extra row(s) or col(s) added or change(s) done in interpret sheet"
            assert report["L8"].value == "B_POINTS", \
            f"extra row(s) or col(s) added or change(s) done in interpret sheet"
        error_msg = None
    except AssertionError as msg:
        error_msg = str(msg)
        print(msg)

    return error_msg


def get_col_letter(worksheet: object, col_name: str) -> str:
    """
    Getting the column letter with specific col name

    Parameters
    ----------
    openpyxl object of current sheet
    str for name of column to get col letter

    Return
    -------
        str for column letter for specific column name
    """
    col_letter = None
    for column_cell in worksheet.iter_cols(1, worksheet.max_column):
        if column_cell[0].value == col_name:
            col_letter = column_cell[0].column_letter

    return col_letter


def write_txt_file(txt_file_name: str, output_dir: str, filename: str, msg: str) -> None:
    """
    write txt file output

    Parameters
    ----------
      str for output txt file name
      str for output dir
      variant workbook file name
      str for error message
    """
    with open(output_dir + txt_file_name, 'a') as file:
        dt = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        file.write(dt+" "+filename+" "+msg+"\n")
        file.close()


def check_interpret_table(df_report: pd.DataFrame, df_included: pd.DataFrame) -> str:
    """
    check if ACMG classification and HGVSc are correctly
    filled in in the interpret table(s)

    Parameters
    ----------
      df from interpret sheet(s)
      df from included sheet

    Return
    ------
      str for error message
    """
    row_index = df_report[df_report.isnull()].index.tolist()
    error_msg = []
    for row in row_index:
        try:
            assert df_report.loc[row, "Germline classification"] is not np.nan, \
            f"empty ACMG classification in interpret table"
            assert df_report.loc[row, "HGVSc"] is not np.nan, \
            f"empty HGVSc in interpret table"
            assert df_report.loc[row, "HGVSc"] in list(df_included['HGVSc']), \
            f"HGVSc in interpret table does not match with that in included sheet"

        except AssertionError as msg:
            error_msg.append(str(msg))
            print(msg)
    error_msg = "".join(error_msg)

    return error_msg


def check_interpreted_col(df: pd.DataFrame) -> str:
    """
    check if interpreted col in included sheet
    is correctly filled in

    Parameters
    ----------
    merged df

    Return
    ------
      str for error message
    """
    row_yes = df[df['Interpreted'] == 'yes'].index.tolist()
    error_msg = []
    for row in range(df.shape[0]):
        if row in row_yes:
            try:
                assert df.loc[row, "Germline classification"] is not np.nan, \
                f"Wrong interpreted column in row {row+1} of included sheet"
            except AssertionError as msg:
                error_msg.append(str(msg))
                print(msg)
        else:
            try:
                assert df.loc[row, 'Interpreted'] == "no", \
                f"wrong entry in row {row+1} of included sheet"
                assert df.loc[row, "Germline classification"] is np.nan, \
                f"Wrong interpreted column in row {row+1} of included sheet"
            except AssertionError as msg:
                error_msg.append(str(msg))
                print(msg)
    error_msg = " ".join(error_msg)

    return error_msg


def get_folder(filename: str) -> str:
    """
    get the folder of input file
    Parameters:
    ----------
      str for input finame 

    Return:
      str for folder name
    """
    folder = os.path.basename(os.path.normpath(os.path.dirname(filename)))
    print(folder)
    return folder


def get_parsed_list(file: str) -> list:
    """
    getting the list of previously parsed workbook

    Parameters
    ----------
    str for ref file that records previously parsed workbook

    Return
    ------
    a list of previously parsed workbook
    """
    f = open(file, "r")
    lines = f.readlines()
    parsed_list = []
    for x in lines:
        parsed_list.append(x.split(' ')[2].split('/')[-1])
    f.close()

    return parsed_list


def check_and_create(dir: str) -> None:
    """
    check if a dir exists and create
    Parameters
    ----------
    str for directory
    """
    if not os.path.exists(dir):
        os.makedirs(dir)


def main():
    arguments = get_command_line_args()
    input_dir = arguments.indir
    if arguments.file:
        input_file = []
        for idx, file in enumerate(arguments.file):
            input_file.append(glob.glob(input_dir+file)[0])
    else:
        input_file = glob.glob(input_dir+"*.xlsx")
    if len(input_file) == 0:
        print("Input file(s) not exist")
    check_and_create(arguments.outdir)
    check_and_create(arguments.completed_dir)
    check_and_create(arguments.logdir)
    if not os.path.isfile(arguments.logdir+PARSED_FILE):
        with open(arguments.logdir+PARSED_FILE, 'w') as file:
            file.close()
    unusual_sample_name = arguments.unusual_sample_name
    with open('parser_config.json') as f:
        config_variable = json.load(f)
    parsed_list = get_parsed_list(arguments.logdir+PARSED_FILE)
    # extract fields from variant workbooks as df and merged
    for filename in input_file:
        print("Running", filename)
        if not filename.split('/')[-1] in parsed_list:
            error_msg_sheet = checking_sheets(filename)
            if not error_msg_sheet:
                df_summary, error_msg_name = get_summary_fields(filename, config_variable,
                                                            unusual_sample_name)
                if not error_msg_name:
                    df_included = get_included_fields(filename)
                    if df_included["Interpreted"].isna().sum() == 0:
                        df_report, error_msg_table = get_report_fields(filename, df_included)
                        if not error_msg_table:
                            if not df_included.empty:
                                df_merged = pd.merge(df_included, df_summary, how="cross")
                                empty_workbook = False
                            else:
                                df_merged = pd.concat([df_summary, df_included], axis=1)
                                empty_workbook = True
                            df_final = pd.merge(df_merged, df_report, on="HGVSc",
                                                 how="left")
                            error_msg_interpreted = check_interpreted_col(df_final)
                            if not error_msg_interpreted:
                                df_final = df_final[['Local ID', 'Linking ID', 'Organisation ID', 'Gene symbol',
                                                     'Chromosome', 'Start', 'Reference allele', 'Alternate allele',
                                                     'Preferred condition name', 'Germline classification', 'Date last evaluated',
                                                     'Comment on classification', 'Collection method', 'Allele origin', 'Affected status',
                                                     'HGVSc', 'Consequence', 'Interpreted', 'Comment', 'Instrument ID', 'Specimen ID',
                                                     'Batch ID', 'Test code', 'Probeset ID', 'Panel', 'Ref genome', 'Organisation',
                                                     'Institution', 'Associated disease', 'Known inheritance', 'Prevalence', 'PVS1',
                                                     'PVS1_evidence', 'PS1', 'PS1_evidence', 'PS2', 'PS2_evidence', 'PS3', 'PS3_evidence',
                                                     'PS4', 'PS4_evidence', 'PM1', 'PM1_evidence', 'PM2', 'PM2_evidence', 'PM3',
                                                     'PM3_evidence', 'PM4', 'PM4_evidence', 'PM5', 'PM5_evidence', 'PM6', 'PM6_evidence',
                                                     'PP1', 'PP1_evidence', 'PP2', 'PP2_evidence', 'PP3', 'PP3_evidence', 'PP4',
                                                     'PP4_evidence', 'BS1', 'BS1_evidence', 'BS2', 'BS2_evidence', 'BS3', 'BS3_evidence',
                                                     'BA1', 'BA1_evidence', 'BP2', 'BP2_evidence', 'BP3', 'BP3_evidence', 'BS4',
                                                     'BS4_evidence', 'BP1', 'BP1_evidence', 'BP4', 'BP4_evidence', 'BP5', 'BP5_evidence',
                                                     'BP7', 'BP7_evidence']]
                                if empty_workbook:
                                    df_final.fillna('zero_variant', inplace=True)
                                else:
                                    if (df_final.Interpreted == 'yes').sum() > 0:
                                        df_clinvar = df_final[df_final["Interpreted"] == 'yes']
                                        df_clinvar = df_clinvar[['Local ID', 'Linking ID',  'Organisation ID', 'Gene symbol', 'Chromosome', 'Start',
                                                                 'Reference allele', 'Alternate allele', 'Preferred condition name',
                                                                 'Germline classification', 'Date last evaluated', 'Comment on classification',
                                                                 'Collection method', 'Allele origin', 'Affected status', 'Ref genome',
                                                                 'HGVSc', 'Consequence', 'Interpreted', 'Instrument ID', 'Specimen ID']]
                                        df_clinvar.to_csv(arguments.outdir + Path(filename).stem + "_clinvar_variants.csv", index=False)
                                        write_txt_file(CLINVAR_FILE, arguments.logdir, filename, "")
                                df_final.to_csv(arguments.outdir + Path(filename).stem + "_all_variants.csv", index=False)
                                write_txt_file(PARSED_FILE, arguments.logdir, filename, "")
                                print("Successfully parsed", filename)
                                shutil.move(filename, arguments.completed_dir)

                            else:
                                write_txt_file(FAILED_FILE, arguments.logdir, filename, error_msg_interpreted)
                        else:
                            write_txt_file(FAILED_FILE, arguments.logdir, filename, error_msg_table)
                    else:
                        print("Interpreted column in included sheet needs to be fixed")
                        write_txt_file(FAILED_FILE, arguments.logdir, filename, "Interpreted column in included sheet needs to be fixed")
                else:
                    write_txt_file(FAILED_FILE, arguments.logdir, filename, error_msg_name)
            else:
                write_txt_file(FAILED_FILE, arguments.logdir, filename, error_msg_sheet)
        else:
            print(filename, "is already parsed")
    print("Done")


if __name__ == "__main__":
    main()
