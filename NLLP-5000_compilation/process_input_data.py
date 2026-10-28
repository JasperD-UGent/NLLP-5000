"""
This module processes Tiberius and Schoonheim's A Frequency Dictionary of Dutch and returns a spreadsheet in which
inflected forms are automatically populated. These forms were checked manually, after which the final NLLP-5000 resource
was created using the `create_nllp5000.py` script.
"""


from lib.data_spreadsheet import convert_raw_input_into_dic, convert_d_headwords_orig_into_unverified_spreadsheet
from lib.logging_setup import get_logger
import argparse
import time


def main(path_input_data: str) -> None:
    """Entry point for the script. Processes the original Routledge dictionary and creates a spreadsheet that serves as
    input to create the final NLLP-5000 resource. The spreadsheet contains automatically populated inflected forms that
    need to be checked manually by an expert.
    
    :param path_input_data: Path where the file containing the original Routledge dictionary in TXT format is located.
    :return: `None`
    """
    # DEFINE TIME STRING OF CURRENT JOB
    timestr = time.strftime("%Y%m%d-%H%M%S")

    # GET LOGGER
    logger = get_logger(f"process_input_data_{timestr}")

    # CREATE SPREADSHEET
    d_headwords_orig = convert_raw_input_into_dic(path_input_data, logger)
    convert_d_headwords_orig_into_unverified_spreadsheet(d_headwords_orig)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument(
        "path_input_data", type=str,
        help="Path where the file containing the original Routledge dictionary in TXT format is located."
    )
    main(**vars(parser.parse_args()))
