from lib import config
import logging
import os
import sys


def get_logger(script_id: str) -> logging.Logger:
    """Get a dedicated file logger for the script, independent of the root logger.

    :param script_id: ID of the script to be used in the name of the log file.
    :return: The logger.
    """
    path_direc_log = config.DIREC_LOGS
    os.makedirs(path_direc_log, exist_ok=True)

    logger = logging.getLogger(script_id)
    logger.setLevel(logging.INFO)
    logger.propagate = False  # do not forward to the root logger (and its handlers)

    job_id = os.environ.get("SLURM_JOB_ID", "local")
    path_log = os.path.join(path_direc_log, f"{job_id}_{script_id}.log")
    open(path_log, "w", encoding="utf-8").close()  # start each run with an empty log file

    if not logger.handlers:
        handler = logging.FileHandler(path_log, mode="a", encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(levelname)s - %(message)s"))
        logger.addHandler(handler)

    return logger
