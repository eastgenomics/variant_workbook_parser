import os
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
        "--input", "--i", help="input file name",
        required=True
    )
    parser.add_argument(
        "--outdir", "--o", help="dir to save output(s)", default="./"
    )
    parser.add_argument(
        "--folder", "--f", help="folder to check", required=True
    )
    args = parser.parse_args()

    return args


def main():
    arguments = get_command_line_args()
    file = open(arguments.input, 'r')
    lines = file.read().splitlines()
    for root, dirs, files in os.walk(arguments.folder):#('C:\\'):
        for line in lines:
            for file in files:
                if line in file:
                    # If we find it, notify us about it and copy it it to C:\NewPath\
                    shutil.copy(os.path.abspath(root + '/' + file), "new_path")


if __name__ == "__main__":
    main()
