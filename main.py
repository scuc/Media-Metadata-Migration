import logging
import logging.config
import os
from logging.handlers import TimedRotatingFileHandler
from time import localtime, strftime
from typing import Tuple

import yaml

import csv_clean
import csv_parse
import csv_clean_final
import create_db


logger = logging.getLogger(__name__)


def set_logger():
    """Setup logging configuration"""
    path = "logging.yaml"

    with open(path, "rt") as f:
        config = yaml.safe_load(f.read())
        logger = logging.config.dictConfig(config)

    return logger


def main(csv_file: str) -> None:
    """Main function to process the CSV file."""
    date = str(strftime("%Y%m%d%H%M", localtime()))
    start_msg = f"\n\
    ================================================================\n \
                Gorilla-Diva Asset Migration Script\n\
                Version: 1.0.0\n\
                Date: {date}\n\
    ================================================================\n\
    \n"
    logger.info(start_msg)
    logger.error(start_msg)

    parsed_csv = csv_parse.parse_csv(date, csv_file)
    clean_csv = csv_clean.csv_clean(date, parsed_csv)
    clean_csv_final = csv_clean_final.csv_clean_final(date, clean_csv)
    create_db.file_to_sqlite(clean_csv_final[1], "gorilla.db", "assets")

    complete_msg = f"\n\n{'='*25}  SCRIPT COMPLETE  {'='*25}"
    logger.info(complete_msg)


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 2:
        print("Usage: python script.py <csv_file>")
        sys.exit(1)

    set_logger()
    csv_file = sys.argv[1]
    main(csv_file)
