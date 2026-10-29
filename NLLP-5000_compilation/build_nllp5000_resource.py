"""
This module builds the final NLLP-5000 resource based on the verified and completed XLSX file returned by the
`process_input_data.py` script.
"""


from lib.data_spreadsheet import convert_verified_spreadsheet_into_nllp5000
from lib.logging_setup import get_logger
import argparse
import time


def main(path_xlsx_filled_in: str) -> None:
    """Entry point for the script. Builds the final NLLP-5000 resource.

    :param path_xlsx_filled_in: Path where the XLSX file containing the verified and completed inflected forms is
        located.
    :return: `None`
    """
    # DEFINE TIME STRING OF CURRENT JOB
    timestr = time.strftime("%Y%m%d-%H%M%S")

    # GET LOGGER
    logger = get_logger(f"build_nllp5000_resource_{timestr}")

    # BUILD NLLP-5000 RESOURCE
    convert_verified_spreadsheet_into_nllp5000(path_xlsx_filled_in, logger)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument(
        "path_xlsx_filled_in", type=str,
        help="Path where the XLSX file containing the verified and completed inflected forms is located."
    )
    main(**vars(parser.parse_args()))
