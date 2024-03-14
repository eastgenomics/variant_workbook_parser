import os
from datetime import datetime
import argparse
import shutil


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
        "--input",
        "--i",
        help="input file that contains the list of verified workbook(s)",
        required=True,
    )
    parser.add_argument(
        "--outdir",
        "--o",
        help="dir to where the workbooks are copied into",
        required=True,
    )
    parser.add_argument(
        "--folder",
        "--f",
        help="folder to check for verified workbooks",
        required=True,
    )
    parser.add_argument(
        "--file_not_found",
        "--fnf",
        help="log file to record files not found",
        default="./workbooks_not_found_clingen.txt",
    )
    args = parser.parse_args()

    return args


def write_txt_file(txt_file_name: str, filename: str) -> None:
    """
    write txt file output

    Parameters
    ----------
      str for output txt file name
      variant workbook file name
    """
    with open(txt_file_name, "a") as file:
        dt = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        file.write(dt + " " + filename + "\n")
        file.close()


def main():
    arguments = get_command_line_args()
    input_file = open(arguments.input, "r")
    lines = input_file.read().splitlines()
    for line in lines:
        found = False
        for root, dirs, files in os.walk(arguments.folder):
            if line in files:
                shutil.copy(
                    os.path.abspath(root + "/" + line), arguments.outdir
                )
                print("found", line, "in", os.path.abspath(root))
                found = True
        if not found:
            write_txt_file(arguments.file_not_found, line)


if __name__ == "__main__":
    main()
