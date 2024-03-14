import json
import os
import sys
import unittest
import pandas as pd
from openpyxl import load_workbook
from mock import patch

sys.path.insert(1, "../")
from variant_workbook_parser import *
from tests import TEST_DATA_DIR

excel_data_NUH = f"{TEST_DATA_DIR}/NUH/cen_snv_test4.xlsx"
excel_data_CUH = f"{TEST_DATA_DIR}/CUH/cen_snv_test2.xlsx"
parsed_data = f"{TEST_DATA_DIR}/workbooks_parsed_all_variants.txt"
excel_data_wrong_HGVSc = f"{TEST_DATA_DIR}/CUH/cen_snv_test2_wrong_HGVSc.xlsx"
excel_data_empty_HGVSc = f"{TEST_DATA_DIR}/CUH/cen_snv_test2_empty_HGVSc.xlsx"
excel_data_empty_ACMG = f"{TEST_DATA_DIR}/CUH/cen_snv_test2_empty_ACMG.xlsx"
excel_data_wrong_summary = (
    f"{TEST_DATA_DIR}/NUH/cen_snv_test4_wrong_summary.xlsx"
)
excel_data_wrong_interpret_row = (
    f"{TEST_DATA_DIR}/NUH/cen_snv_test4_wrong_interpret_row.xlsx"
)
excel_data_wrong_interpret_col = (
    f"{TEST_DATA_DIR}/NUH/cen_snv_test4_wrong_interpret_col.xlsx"
)
excel_data_wrong_interpreted = (
    f"{TEST_DATA_DIR}/CUH/cen_snv_test2_wrong_interpreted.xlsx"
)

with open(f"{TEST_DATA_DIR}/test_parser_config.json") as f:
    config_variable = json.load(f)


class TestParserScript(unittest.TestCase):
    """
    Tests to ensure that all functions in variant_workbook_parser.py
    works as expected
    """

    def test_get_parsed_list(self):
        """
        Test "get_parsed_list" generates a list containing file name
        recorded in workbooks_parsed_all_variants.txt
        """
        parsed_list = get_parsed_list(parsed_data)
        self.assertTrue(len(parsed_list) == 4)
        self.assertTrue(
            parsed_list
            == [
                "cen_snv_test2.xlsx",
                "cen_snv_test4.xlsx",
                "cen_snv_test3.xlsx",
                "cen_snv_test5.xlsx",
            ]
        )

    def test_get_folder(self):
        """
        Test "test_get_folder" generates the correct folder
        where the workbook exists
        """
        NUH_folder = get_folder(excel_data_NUH)
        self.assertTrue(NUH_folder == "NUH")
        CUH_folder = get_folder(excel_data_CUH)
        self.assertTrue(CUH_folder == "CUH")

    def test_get_included_fields(self):
        """
        Test "get_included_fields" generates df with expected shape
        (2 rows and 11 columns in test case)
        Test generated df has expected columns
        Test generated df has expected Interpreted column values
        in lowercase (no,yes in test case)
        Test random contents of df and see if they are as expected
        (Start and HGVSc in test case)
        """
        df = get_included_fields(excel_data_CUH)
        self.assertTrue(df.shape[0] == 2)
        self.assertTrue(df.shape[1] == 11)
        self.assertTrue(
            list(df.columns)
            == [
                "Chromosome",
                "Start",
                "Reference allele",
                "Alternate allele",
                "Gene symbol",
                "HGVSc",
                "Consequence",
                "Interpreted",
                "Comment",
                "Local ID",
                "Linking ID",
            ]
        )
        self.assertTrue(list(df["Interpreted"]) == ["no", "yes"])
        self.assertTrue(df["Start"][0] == 135773000)
        self.assertTrue(df["HGVSc"][1] == "NM_000548.5:c.4255C>T")

    def test_get_report_fields(self):
        """
        Test "get_report_fields" generates df with expected shape
        (1 row and 58 columns in test case)
        Test generated df has expected columns
        Test generated df has the correct values for "HGVSc" and
        "Germline classification"
        Test generated df has the correct values for the strength
        and evidence columns (some are expected to be np.nan)
        Test there is no error message
        """
        df_included = get_included_fields(excel_data_CUH)
        df_report, msg = get_report_fields(excel_data_CUH, df_included)
        self.assertTrue(df_report.shape[0] == 1)
        self.assertTrue(df_report.shape[1] == 58)
        self.assertTrue(
            list(df_report.columns)
            == [
                "Associated disease",
                "Known inheritance",
                "Prevalence",
                "HGVSc",
                "Germline classification",
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
                "BP7",
                "BP7_evidence",
                "Comment on classification",
            ]
        )
        self.assertTrue(df_report["HGVSc"][0] == "NM_000548.5:c.4255C>T")
        self.assertTrue(
            df_report["Germline classification"][0] == "Pathogenic"
        )
        self.assertTrue(
            df_report["Comment on classification"][0] == "PVS1,PS4_Moderate"
        )
        self.assertTrue(df_report["PVS1"][0] == "Very Strong")
        self.assertTrue(df_report["PS4"][0] == "Moderate")
        self.assertTrue(
            df_report["PVS1_evidence"][0]
            == (
                "Exon present in all transcripts on gnomAD. "
                "LOF known mechanism of disease. "
                "Predicted to undergo nonsense-mediated decay."
            )
        )
        self.assertTrue(
            df_report["PS4_evidence"][0]
            == (
                "PMID: 10205261 (1 case, Roach et al criteria), "
                "35870981 (1 case, 2012 International TS Complex "
                "Consensus Conference criteria), 12111193 "
                "(1 case, Roach et al criteria), 28065512 "
                "(1 case, 2012 International TS Complex "
                "Consensus Conference criteria)."
            )
        )
        self.assertTrue(df_report["PP1"][0] is np.nan)
        self.assertTrue(df_report["PP1_evidence"][0] is np.nan)
        self.assertTrue(df_report["PP3"][0] is np.nan)
        self.assertTrue(df_report["PP3_evidence"][0] is np.nan)
        self.assertTrue(df_report["BP4_evidence"][0] is np.nan)
        self.assertTrue(msg == "")

    def test_check_interpret_table_correct_wb(self):
        """
        Test df_report has expected HGVSc and Germline classification
        so the error msg is empty
        """
        df_included = get_included_fields(excel_data_CUH)
        df_report, msg = get_report_fields(excel_data_CUH, df_included)
        error_msg = check_interpret_table(df_report, df_included)
        self.assertTrue(error_msg == "")

    def test_check_interpret_table_wrong_HGVSc(self):
        """
        Test if the wrong entry of HGVSc in df_report
        is captured as error
        """
        df_included = get_included_fields(excel_data_wrong_HGVSc)
        df_report, msg = get_report_fields(excel_data_wrong_HGVSc, df_included)
        error_msg = check_interpret_table(df_report, df_included)
        self.assertTrue(
            error_msg
            == (
                "HGVSc in interpret table does "
                "not match with that in included sheet"
            )
        )

    def test_check_interpret_table_empty_HGVSc(self):
        """
        Test if no entry of HGVSc in df_report
        is captured as error
        """
        df_included = get_included_fields(excel_data_empty_HGVSc)
        df_report, msg = get_report_fields(excel_data_empty_HGVSc, df_included)
        error_msg = check_interpret_table(df_report, df_included)
        self.assertTrue(error_msg == ("empty HGVSc in interpret table"))

    def test_check_interpret_table_empty_ACMG(self):
        """
        Test if no entry of ACMG classification in df_report
        is captured as error
        """
        df_included = get_included_fields(excel_data_empty_ACMG)
        df_report, msg = get_report_fields(excel_data_empty_ACMG, df_included)
        error_msg = check_interpret_table(df_report, df_included)
        self.assertTrue(
            error_msg == ("empty ACMG classification in interpret table")
        )

    def test_checking_sheet_wrong_summary(self):
        """
        Test if change done in columns of summary sheet is captured as error
        """
        msg = checking_sheets(excel_data_wrong_summary)
        self.assertTrue(
            msg == "extra col(s) added or change(s) done in summary sheet"
        )

    def test_checking_sheet_wrong_interpret_row(self):
        """
        Test if change done in rows of interpret sheet is captured as error
        """
        msg = checking_sheets(excel_data_wrong_interpret_row)
        self.assertTrue(
            msg == "extra row(s) or col(s) added or change(s) done in "
            "interpret sheet"
        )

    def test_checking_sheet_wrong_interpret_col(self):
        """
        Test if change done in cols of interpret sheet is captured as error
        """
        msg = checking_sheets(excel_data_wrong_interpret_col)
        self.assertTrue(
            msg
            == "extra row(s) or col(s) added or change(s) done in interpret "
            "sheet"
        )

    def test_get_summary_fields(self):
        """
        Test "get_summary_fields" generates df with expected shape
        (1 row and 15 columns in test case)
        Test the "Preferred condition name" is split as expected
        (Tuberous sclerosis in test case)
        Test if the "Ref genome" is correctly extracted
        (GRCh37.p13 in test case)
        Test if the "Date last evaluated" is pd date time format
        """
        df, msg = get_summary_fields(excel_data_NUH, config_variable, False)
        self.assertTrue(df.shape[0] == 1)
        self.assertTrue(df.shape[1] == 16)
        self.assertTrue(
            df["Preferred condition name"][0]
            == (
                "Inherited breast cancer and ovarian cancer;Inherited "
                "breast cancer and ovarian cancer"
            )
        )
        self.assertTrue(df["Ref genome"][0] == "GRCh37.p13")
        self.assertTrue(
            type(df["Date last evaluated"][0])
            == pd._libs.tslibs.timestamps.Timestamp
        )

    def test_check_sample_name_no_error(self):
        """
        Test "check_sample_name" correctly check the naming
        No error message is expected is this test
        """
        instrumentID = "124256019"
        sample_ID = "23201R0067"
        batchID = "23NGCEN32"
        testcode = "9527"
        probesetID = "99347387"

        msg = check_sample_name(
            instrumentID, sample_ID, batchID, testcode, probesetID
        )
        self.assertTrue(msg is None)

    def test_check_instrument_name_error(self):
        """
        Test if abnormal instrumentId is captured
        Normal instrument ID has 9 digits
        """
        instrumentID = "12425609x"
        sample_ID = "23201R0067"
        batchID = "23NGCEN32"
        testcode = "9527"
        probesetID = "99347387"

        msg = check_sample_name(
            instrumentID, sample_ID, batchID, testcode, probesetID
        )
        self.assertTrue(msg == "Unusual name for instrumentID")

    def test_check_sample_name_error(self):
        """
        Test if abnormal sample Id is captured
        Normal sample ID has 5digits,1alphabet,4digits
        """
        instrumentID = "124256019"
        sample_ID = "2320100067"
        batchID = "23NGCEN32"
        testcode = "9527"
        probesetID = "99347387"

        msg = check_sample_name(
            instrumentID, sample_ID, batchID, testcode, probesetID
        )
        self.assertTrue(msg == "Unusual sampleID")

    def test_check_batch_ID_error(self):
        """
        Test if abnormal batch Id is captured
        Normal batch ID has 2digits, 5alphabets,1 or more digits
        """
        instrumentID = "124256019"
        sample_ID = "23201R0067"
        batchID = "23NG2EN32"
        testcode = "9527"
        probesetID = "99347387"

        msg = check_sample_name(
            instrumentID, sample_ID, batchID, testcode, probesetID
        )
        self.assertTrue(msg == "Unusual batchID")

    def test_check_test_code_error(self):
        """
        Test if abnormal test code is captured
        Normal test code has 4 digits
        """
        instrumentID = "124256019"
        sample_ID = "23201R0067"
        batchID = "23NGCEN32"
        testcode = "9527A"
        probesetID = "99347387"

        msg = check_sample_name(
            instrumentID, sample_ID, batchID, testcode, probesetID
        )
        self.assertTrue(msg == "Unusual testcode")

    def test_check_probeset_ID_too_long(self):
        """
        Test if abnormal probeset ID is captured
        Normal probeset ID has length between 0 to 20
        can be all numbers or mixed of number and alphabet
        """
        instrumentID = "124256019"
        sample_ID = "23201R0067"
        batchID = "23NGCEN32"
        testcode = "9527"
        probesetID = "99347387344522344678976555"

        msg = check_sample_name(
            instrumentID, sample_ID, batchID, testcode, probesetID
        )
        self.assertTrue(msg == "probesetID is too long/short")

    def test_check_probeset_ID_wrong_format(self):
        """
        Test if abnormal probeset ID is captured
        Normal probeset ID has length between 0 to 20
        can be all numbers or mixed of number and alphabet
        """
        instrumentID = "124256019"
        sample_ID = "23201R0067"
        batchID = "23NGCEN32"
        testcode = "9527"
        probesetID = "abdfsettd"

        msg = check_sample_name(
            instrumentID, sample_ID, batchID, testcode, probesetID
        )
        self.assertTrue(msg == "Unusual probesetID")

    def test_get_col_letter(self):
        """
        Test the correct col letter is retrieved
        Randomly checked a few columns
        (AU for Interpreted, AT for Comment and B for POS)
        """
        workbook = load_workbook(excel_data_CUH)
        interpreted_col_letter = get_col_letter(
            workbook["included"], "Interpreted"
        )
        comment_col_letter = get_col_letter(workbook["included"], "Comment")
        pos_col_letter = get_col_letter(workbook["excluded"], "POS")

        self.assertTrue(interpreted_col_letter == "AU")
        self.assertTrue(comment_col_letter == "AT")
        self.assertTrue(pos_col_letter == "B")

    def test_interpreted_col_correct(self):
        """
        Test if interpreted col (yes/no) is correctly filled in
        """
        unusual_sample_name = False
        df_summary, error_msg_name = get_summary_fields(
            excel_data_CUH, config_variable, unusual_sample_name
        )
        df_included = get_included_fields(excel_data_CUH)
        df_report, error_msg_table = get_report_fields(
            excel_data_CUH, df_included
        )
        df_merged = pd.merge(df_included, df_summary, how="cross")
        df_final = pd.merge(df_merged, df_report, on="HGVSc", how="left")
        msg = check_interpreted_col(df_final)
        self.assertTrue(msg == "")

    def test_interpreted_col_wrong(self):
        """
        Test if interpreted col (yes/no) is correctly filled in
        Expected to throw error for both row 1 and row 2 in test case
        """
        unusual_sample_name = False
        df_summary, error_msg_name = get_summary_fields(
            excel_data_wrong_interpreted, config_variable, unusual_sample_name
        )
        df_included = get_included_fields(excel_data_wrong_interpreted)
        df_report, error_msg_table = get_report_fields(
            excel_data_wrong_interpreted, df_included
        )
        df_merged = pd.merge(df_included, df_summary, how="cross")
        df_final = pd.merge(df_merged, df_report, on="HGVSc", how="left")
        msg = check_interpreted_col(df_final)
        self.assertTrue(
            msg
            == (
                "Wrong interpreted column in row 1 of "
                "included sheet Wrong interpreted column "
                "in row 2 of included sheet"
            )
        )

    def test_get_command_line_args(self):
        """
        Test if parser args are correctly read in
        """
        parser_args = get_command_line_args(
            [
                "--i",
                "/test_data/CUH/",
                "--f",
                "cen_snv_test2.xlsx",
                "--o",
                "/test_data/output/",
                "--pf",
                "/test_data/output/log/workbooks_parsed_all_variants.txt",
                "--cf",
                "/test_data/output/log/workbooks_parsed_clinvar_variants.txt",
                "--ff",
                "/test_data/output/log/workbooks_fail_to_parse.txt",
                "--cd",
                "/test_data/output/completed_wb/",
                "--unusual_sample_name",
                "--tk",
                "abcdefgh",
            ]
        )
        self.assertTrue(parser_args.indir == "/test_data/CUH/")
        self.assertTrue(parser_args.file == ["cen_snv_test2.xlsx"])
        self.assertTrue(parser_args.outdir == "/test_data/output/")
        self.assertTrue(
            parser_args.parsed_file
            == "/test_data/output/log/workbooks_parsed_all_variants.txt"
        )
        self.assertTrue(
            parser_args.clinvar_file
            == "/test_data/output/log/workbooks_parsed_clinvar_variants.txt"
        )
        self.assertTrue(
            parser_args.failed_file
            == "/test_data/output/log/workbooks_fail_to_parse.txt"
        )
        self.assertTrue(
            parser_args.completed_dir == "/test_data/output/completed_wb/"
        )
        self.assertTrue(parser_args.unusual_sample_name is True)
        self.assertTrue(parser_args.token == "abcdefgh")

    def test_write_txt_file(self):
        """
        Test if "write_txt_file" writes the log file as expected
        "abc.xlsx" is expected filename in the log file
        "testing_msg" is expected msg in the log file
        """
        outfile_path = "./test_log_file.txt"
        write_txt_file("test_log_file.txt", "abc.xlsx", "testing_msg")
        contents = open(outfile_path).read()
        os.remove(outfile_path)
        self.assertEqual(contents.split(" ")[2], "abc.xlsx")
        self.assertEqual(contents.split(" ")[3], "testing_msg\n")

    @patch("os.path.exists")
    @patch("os.makedirs")
    def test_check_and_create_folder(self, patch_makedirs, patch_exists):
        """
        Test if "check_and_create" is called if folder does not exists
        """
        patch_exists.return_value = False
        check_and_create_folder("./new_folder_created")
        assert patch_makedirs.called is True


if __name__ == "__main__":
    unittest.main()
