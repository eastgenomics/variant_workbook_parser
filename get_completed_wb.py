import os
from datetime import datetime
import argparse
import shutil

FILE_NOT_FOUND = "workbooks_not_found_clingen.txt"
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
        "--input", "--i", help="input file name",
        required=True
    )
    parser.add_argument(
        "--outdir", "--o", help="dir to where the workbooks are copied into", required=True
    )
    parser.add_argument(
        "--folder", "--f", help="folder to check for completed workbooks", required=True
    )
    parser.add_argument(
        "--logdir", "--ld", help="dir to log txt files", default="./"
    )
    args = parser.parse_args()

    return args

def write_txt_file(txt_file_name: str, output_dir: str, filename: str) -> None:
    """
    write txt file output

    Parameters
    ----------
      str for output txt file name
      str for output dir
      variant workbook file name
    """
    with open(output_dir + txt_file_name, 'a') as file:
        dt = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        file.write(dt+" "+filename + "\n")
        file.close()


def main():
    arguments = get_command_line_args()
    file = open(arguments.input, 'r')
    lines = file.read().splitlines()
    for root, dirs, files in os.walk(arguments.folder):
        for line in lines:
            found = False
            for file in files:
                if line in file:
                    shutil.copy(os.path.abspath(root + '/' + file), arguments.outdir)
                    print("found in", os.path.abspath(root))
                    found = True
            if not found:
                write_txt_file(FILE_NOT_FOUND, arguments.logdir, line)


if __name__ == "__main__":
    main()
