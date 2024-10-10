#! /usr/bin/env python3

import logging
import os
import re
from typing import Optional

import pandas as pd

import config as cfg

# Configure logger
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

# Constants for regex patterns
include_patterns = [
    "VM",
    "EM",
    "AVP",
    "PPRO",
    "FCP",
    "PTS",
    "GRFX",
    "GFX",
    "UHD",
    "XDCAM",
    "XDCAMHD",
    "WAV",
    "WAVS",
    "OUTGOING",
]
exclude_patterns = ["PGS", "SVM", "SDM", "CEM", "CVM", "PROMO"]

# include_pattern = r"(?:[-_])(VM|EM|AVP|PPRO|FCP|PTS|GRFX|GFX|UHD|XDCAM|XDCAMHD|WAV|WAVS|OUTGOING)(?:[-_]|$)"
exclude_pattern = r"(?:[-_])(PGS|SVM|SDM|CEM|PROMO|TACHYON)(?:[-_]|$)"

# pattern = f"^(?!.*{exclude_pattern}).*{include_pattern}.*$"
pattern = f"^(?!.*{exclude_pattern}).*$"
compiled_pattern = re.compile(pattern, re.IGNORECASE)

# Constants
MAX_INDEX_COUNT = 200000

config = cfg.get_config()
root_path = config["paths"]["root_path"]
csv_path = config["paths"]["csv_path"]


def parse_csv(date: str, merged_csv: str):
    index_count = 0
    parsed_count = 0

    os.chdir(csv_path)
    parsed_csv = f"{date}_gor_diva_merged_parsed.csv"

    with open(merged_csv, mode="r", encoding="utf-8-sig") as m_csv, open(
        parsed_csv, "w+", newline=""
    ) as p_csv:
        try:
            logger.info("START GORILLA-DIVA DB PARSE")
            pd_reader = pd.read_csv(m_csv, header=0)
            df = pd.DataFrame(pd_reader)

            # write header to parsed csv
            write_row(p_csv, df.iloc[0], parsed_count)
            parsed_count += 1

            for index, row in df.iterrows():
                name = str(row["NAME"]).upper()
                # logger.info(f"Index {index_count}: {name}")

                if index_count <= MAX_INDEX_COUNT:
                    if should_parse_row(name):
                        parsed_count += 1
                    else:
                        write_row(p_csv, row, parsed_count)

                    index_count += 1

            logger.info("GORILLA-DIVA DB PARSE COMPLETE")
            return parsed_csv

        except Exception as e:
            logger.exception("Exception raised during the Gor-Diva DB Parse.")
            return None


def should_parse_row(name: str) -> bool:
    """
    Determine if the row should be parsed based on the name.
    """
    em_check = compiled_pattern.search(name)

    print(f"\nEM CHECK: {em_check}\n")

    if em_check is not None:
        return False
    else:
        logger.info(f"\nParsing out: {name}\n")
        return True


def write_row(p_csv, row: pd.Series, parsed_count: int):
    """
    Write the row to the parsed csv file.
    """
    dft = pd.DataFrame(row).transpose()
    header = parsed_count == 0
    dft.to_csv(p_csv, mode="a", index=False, header=header)


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("Usage: python script.py <date> <merged_csv>")
        sys.exit(1)

    date = sys.argv[1]
    merged_csv = sys.argv[2]
    parsed_csv = parse_csv(merged_csv)
