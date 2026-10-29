from pathlib import Path


# (PATHS TO) DIRECTORIES
DIREC_ROOT = Path(__file__).resolve().parent.parent  # project root, independent of the current working directory

DIREC_INP = DIREC_ROOT / "input"
DIREC_SPREADSHEET_HEADERS = DIREC_INP / "spreadsheet" / "headers"

DIREC_OUTP = DIREC_ROOT / "output"
DIREC_INP_DATA_PROCSD = DIREC_OUTP / "inputData_procsd"
DIREC_FINAL_RESOURCE = DIREC_OUTP / "finalResource"

DIREC_LOGS = DIREC_ROOT / "logs"

DIREC_TEMP = DIREC_ROOT / "temp"
