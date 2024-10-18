import logging
import os
from typing import Optional, Tuple

import pandas as pd

import config as cfg
import get_mediainfo as gmi
import re

# Configure logger
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)


def csv_clean_final(date: str, cleaned_csv: Optional[str] = None) -> Tuple[str, str]:
    """
    Final Pass on the cleaned data to fix bad rows or values, and remove any unwanted data.
    """
    config = cfg.get_config()

    root_path, db_path, csv_path = (
        config["paths"]["root_path"],
        config["paths"]["db_path"],
        config["paths"]["csv_path"],
    )

    os.chdir(csv_path)

    clean_csv_final = f"{date}_gor_diva_merged_cleaned_final.csv"
    logger.info("START FINAL PASS of GORILLA-DIVA DB CLEAN")

    df = pd.read_csv(cleaned_csv, header=0)
    df.index.name = "ROWID"

    for index, row in df.iterrows():
        if row["TITLETYPE"] != "video":
            continue

        else:
            logger.info(f"Final Check for row {index} for {row['NAME']}")
            original_row = row.copy()
            if pd.isnull(row["CODEC"]):
                row = check_codec(row)
            if pd.isnull(row["FRAMERATE"]):
                row = check_framerate(row, df)
            if pd.isnull(row["V_WIDTH"]) or pd.isnull(row["V_HEIGHT"]):
                row = check_resolution(row)

            # Log changes
            log_changes(original_row, row, index)

            # Update specific columns in the DataFrame
            df.loc[index, ["CODEC", "FRAMERATE", "V_WIDTH", "V_HEIGHT"]] = row[
                ["CODEC", "FRAMERATE", "V_WIDTH", "V_HEIGHT"]
            ]

    df.to_csv(clean_csv_final, index=False)

    logger.info("FINAL PASS COMPLETE")
    return clean_csv_final, os.path.join(csv_path, clean_csv_final)


def clean_content_type(row):
    content_type_mappings = {
        "PROMOSELECTS": "PSEL",
        "PROMO": "PSEL",
        "PROMOS": "PSEL",
        "SELECTS": "PSEL",
        "CLEANCOVERS": "CCOV",
        "CLEAN": "CCOV",
        "COVERS": "CCOV",
        "COVER": "CCOV",
        "BEHINDTHESCENES": "BTS",
        "BEHINDTHESCENE": "BTS",
        "BEHIND THE SCENES": "BTS",
        "BEHIND THE SCENE": "BTS",
        "PATCH": "PATCH",
        "FIX": "PATCH",
        "FIXES": "PATCH",
        "CREDIT PATCH": "CREDITS",
        "CREDITS": "CREDITS",
        "CREDP": "CREDITS",
        "DELETED": "DSCN",
        "DELETED SCENES": "DSCN",
        "DELETEDSCENES": "DSCN",
        "DELETED SCENE": "DSCN",
        "DELETEDSCENE": "DSCN",
    }

    updated_ctype = []
    for ctype in row["CONTENT_TYPE"].split(","):
        ctype_upper = ctype.strip().upper()
        updated_ctype.append(content_type_mappings.get(ctype_upper, ctype_upper))

    return ",".join(updated_ctype)


def log_changes(original_row, updated_row, index):
    changes = []
    for col in original_row.index:
        if original_row[col] != updated_row[col]:
            changes.append(f"{col}: {original_row[col]} -> {updated_row[col]}")
    if changes:
        logger.info(f"Changes for row {index}:\n" + "\n".join(changes))


def check_codec(row):
    # logger.info(f"Checking codec for {row['NAME']}")
    codec_list = [
        "XAVC",
        "UHD",
        "H264",
        "MPEG2",
        "MPEG4",
        "DV",
        "DNXHD",
        "PRORES",
        "XDCAM",
        "AVC",
    ]

    if row["CODEC"] != "NULL" and not pd.isnull(row["CODEC"]):
        return row

    codec = gmi.get_codec(row)
    if codec[1] != "NULL":
        logger.info(f"Found codec {codec[1]} for {row['NAME']}")
        row["CODEC"] = codec
    elif "VM" in str(row["CONTENT_TYPE"]) or "EM" in str(row["CONTENT_TYPE"]):
        logger.info("FILENAME indicates a VM or EM file, setting codec to 'PRORES'")
        row["CODEC"] = "PRORES"
    else:
        logger.error(f"Could not find codec for {row['NAME']}")

    return row


def check_framerate(row, df):
    # logger.info(f"Checking framerate for {row['NAME']}")
    framerate_list = ["23.976", "23.98", "25", "29.97", "59.94"]

    if row["TITLETYPE"] != "video":
        pass
    if str(row["FRAMERATE"]) in framerate_list:
        logger.info(
            f"Skipping framerate check for {row['NAME']}, framerate is {row['FRAMERATE']}"
        )
        return row

    if (
        row["FRAMERATE"] == "NULL"
        or row["FRAMERATE"] == "00"
        or pd.isnull(row["FRAMERATE"])
        or len(str(row["FRAMERATE"])) == 0
    ):
        logger.info(f"No framerate for {row['OBJECTNM']}, attempting best estimate.")
        # #split the trafficode value to eliminate the suffix
        org_traffic_code = re.split(r"[_-]", row["TRAFFIC_CODE"])
        df_traffic_code = df["TRAFFIC_CODE"].str.split(r"[_-]", expand=True)
        matching_rows = df[df_traffic_code[0] == org_traffic_code[0]]

        em_matching_rows = matching_rows[
            (matching_rows["CONTENT_TYPE"].str.contains("EM", na=False))
            & (matching_rows["TITLETYPE"] == "video")
        ]

        if len(em_matching_rows) == 0:
            logger.info(
                f"No matching rows for TRAFFIC_CODE {org_traffic_code[0]} with CONTENT_TYPE containing 'EM'"
            )
            return row

        else:
            logger.info(
                f"Matching rows for TRAFFIC_CODE {org_traffic_code[0]} with CONTENT_TYPE containing 'EM': { em_matching_rows['OBJECTNM']}"
            )
            for _, matching_row in em_matching_rows.iterrows():
                logger.info(matching_row.to_dict())
                framerate = matching_row["FRAMERATE"]
                if framerate != "NULL":
                    logger.info(
                        f"Found framerate {framerate} for {matching_row['OBJECTNM']}"
                    )
                    row["FRAMERATE"] = framerate
                else:
                    logger.error(
                        f"Could not find framerate for {matching_row['OBJECTNM']}"
                    )

    else:  # Framerate is not NULL
        logger.info(
            f"Final Check of Framerate for {row['NAME']} - no changes made: {row['FRAMERATE']}"
        )

    return row


def check_resolution(row):
    # logger.info(f"Checking resolution for {row['NAME']}")

    if (
        row["V_WIDTH"] != "NULL"
        and row["V_HEIGHT"] != "NULL"
        and not pd.isnull(row["V_WIDTH"])
        and not pd.isnull(row["V_HEIGHT"])
        and len(str(row["V_WIDTH"])) != 0
        and len(str(row["V_HEIGHT"])) != 0
    ):
        return row

    else:
        logger.info(f"No resolution for {row['NAME']}, attempting best estimate.")

        if (
            "UHD" in str(row["CONTENT_TYPE"])
            or "HDR" in str(row["NAME"])
            or "XAVC" in str(row["CODEC"])
        ):
            row["V_WIDTH"], row["V_HEIGHT"] = "3840", "2160"
            logger.info(
                f"Setting (estimated) resolution to '3840 x 2160' for {row['NAME']}"
            )
        elif (
            "VM" in str(row["CONTENT_TYPE"])
            or "EM" in str(row["CONTENT_TYPE"])
            and "UHD" not in str(row["CONTENT_TYPE"])
        ):
            row["V_WIDTH"], row["V_HEIGHT"] = "1920", "1080"

        elif (
            "_xdcam_" in row["FILENAME"].lower()
            or "_xdcamhd_" in row["FILENAME"].lower()
        ):
            row["V_WIDTH"], row["V_HEIGHT"] = "1920", "1080"

        else:
            logger.info(f"Could not estimate resolution for {row['NAME']}")
            return row

        logger.info(
            f"Setting (estimated) resolution to '1920 x 1080' for {row['NAME']}"
        )

    return row


if __name__ == "__main__":
    date = "202408071645"
    cleaned_csv = "202408071600_gor_diva_merged_cleaned.csv"
    csv_clean_final(date, cleaned_csv)
