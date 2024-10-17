#! /usr/bin/env python3

import logging
import os
import re
from typing import Optional, Tuple

import pandas as pd

import config as cfg
import database as db
import get_mediainfo as gmi

# Configure logger
logger = logging.getLogger(__name__)

# Constants for regex patterns
VIDEO_PATTERN_1 = r"(?<![0-9]|[A-Z])(?<=[-_])(VM|EM|UHD)(?=(-|_|[1-5])?)(?![A-Z])"
VIDEO_PATTERN_2 = (
    r"(?<![0-9]|[A-Z])(?<=[-_])(SMLS|TXTLS|TXTD|CTC|HDR)(?=(-|_|[1-5])?)(?![A-Z])"
)
VIDEO_PATTERN_3 = r"(?<=[_-])(PATCH|MXF|MOV)(?=(-|_|[1-5])?)(?![A-Z])"
VIDEO_PATTERN_4 = r"((?<=(-|_| ))(PRORES(HD)?|XDCAM(HD)?|DNX(HD)?)(?=(-|_|[1-5]|HD)?))"
VIDEO_PATTERN_5 = r"(?<![0-9A-Z])(?<=(-|_))(DV100|IMX50|CEM|CVM|SVM|PGS|DOLBY|PROMOSELECTS|CLEANCOVERS|CREDITPATCH|DELETEDSCENES)(?=(-|_|[1-5])?)"
# VIDEO_PATTERN_6 = r"(?<![0-9A-Z])(?<=(-|_))(23(?:\.?98|\.?976|\.?97|\.?94)?|25|29(?:\.?97)?|59(?:\.?94)?|NTSC|PAL|24P|720P|1080)(?=[IPip]?)(?=(-|_)?)"
MISC_VIDEO_PATTERN = r"(CCOV|DSCN|PSEL|CREDP|DELETEDSCENE(?:S)?|PROMOSELECT(?:S)?|PROMO|SELECT(S)?|CLEAN(?:\ )?COVER(?:S)?|CLEAN|CREDIT(?:\ )?PATCH(?:ES)?|PATCH(?: ES)?|CREDIT(?:S)?|DELETED(?:\ )?SCENE(?:S)?|DELETED|FIX(?:ES)?)|TEXTLESS|TEXTED|TXLS|TXTD|BEHIND(?:\ )?THE(?:\ )?SCENE(?:\S)?"

ARCHIVE_PATTERN = r"((?<=(-|_)))?(AVP|PPRO|FCP|PTS|AVP|GRFX|GFX|WAV|WAVS|MDE|SPLITS|GFXPACKAGE|GRAPHICS|MIXES|AUDIO|STEMS|5\.1)(?=(-|_)?)(?![0-9A-Z])"
DOC_PATTERN = r"((?<![0-9]|[A-Za-z])|(?<=(-|_)))(OUTGOING[-|_]?(QC|UHD)?)(?=(-|_)?)"

# Constants for abbreviations11
ABBREVIATIONS = {
    "PROMOSELECTS": "PSEL",
    "CLEANCOVERS": "CCOV",
    "CREDITPATCH": "CREDP",
    "DELETEDSCENES": "DSCN",
}


def csv_clean(date: str, parsed_csv: Optional[str] = None) -> Tuple[str, str]:
    config = cfg.get_config()
    root_path, db_path, csv_path = (
        config["paths"]["root_path"],
        config["paths"]["db_path"],
        config["paths"]["csv_path"],
    )

    os.chdir(csv_path)

    parsed_csv = parsed_csv or f"{date}_gor_diva_merged_parsed.csv"
    clean_csv = f"{date}_gor_diva_merged_cleaned.csv"
    logger.info("START GORILLA-DIVA DB CLEAN")

    try:
        df = pd.read_csv(parsed_csv, header=0)
        df.index.name = "ROWID"
        df = df.astype({"METAXML": str})  # Ensure METAXML is a string

        # Insert new columns with default values
        new_columns = {
            "TITLETYPE": "NULL",
            "FRAMERATE": "NULL",
            "CODEC": "NULL",
            "V_WIDTH": "NULL",
            "V_HEIGHT": "NULL",
            "TRAFFIC_CODE": "NULL",
            "DURATION_MS": "NULL",
            "XML_CREATED": 0,
            "PROXY_COPIED": 0,
            "CONTENT_TYPE": "NULL",
            "FILENAME": "NULL",
        }
        for col, val in new_columns.items():
            df.insert(len(df.columns), col, val)

        df.to_csv(clean_csv)

        for index, row in df.iterrows():
            cleaned_name = clean_name(str(row["NAME"]).upper())
            df.at[index, "NAME"] = cleaned_name
            df.at[index, "TRAFFIC_CODE"] = get_traffic_code(cleaned_name)

            if row["_merge"] != "both":
                df.drop(index=index, inplace=True)
                continue

            df.at[index, "METAXML"] = (
                clean_metaxml(row["METAXML"], cleaned_name)
                if not pd.isnull(row["METAXML"])
                else "NULL"
            )

            content_type_v = get_content_type_v(cleaned_name)
            content_type_a = get_content_type_a(cleaned_name)
            content_type_d = get_content_type_d(cleaned_name)
            content_type_misc = get_content_type_misc(cleaned_name)

            if content_type_d:
                set_document_info(df, index, cleaned_name)
            elif content_type_v and not content_type_a:
                set_video_info(df, index, row, cleaned_name, content_type_v)
            elif content_type_a and not content_type_v:
                set_archive_info(
                    df, index, cleaned_name, content_type_a, row["SOURCECREATEDT"]
                )
            elif content_type_v and content_type_a:
                set_mixed_info(
                    df,
                    index,
                    cleaned_name,
                    content_type_v,
                    content_type_a,
                    row["SOURCECREATEDT"],
                )
            elif (
                content_type_misc
                and not content_type_v
                and not content_type_a
                and not content_type_d
            ):
                logger.info(f"Miscellaneous content type: {content_type_misc}")
                set_video_info(df, index, row, cleaned_name, content_type_misc)
            else:
                logger.info(f"Cannot determine content type for: {cleaned_name}")
                logger.info(f"Defaulting to video content for: {cleaned_name}")
                logger.info(f"Checking for video metadata for: {cleaned_name}")
                set_video_info(df, index, row, cleaned_name, content_type_misc)

        df.drop("METAXML", axis=1, inplace=True)
        df.to_csv(clean_csv)

        return clean_csv

    except Exception as e:
        logger.exception(
            f"Exception raised on the Gor-Diva DB Clean. Error Message: {str(e)}"
        )
        raise


def get_content_type_v(cleaned_name: str) -> Optional[str]:
    patterns = [
        VIDEO_PATTERN_1,
        VIDEO_PATTERN_2,
        VIDEO_PATTERN_3,
        VIDEO_PATTERN_4,
        VIDEO_PATTERN_5,
    ]
    return ",".join(
        [
            re.search(p, cleaned_name, re.IGNORECASE).group(0)
            for p in patterns
            if re.search(p, cleaned_name, re.IGNORECASE)
        ]
    )


def get_content_type_a(cleaned_name: str) -> Optional[str]:
    match = re.search(ARCHIVE_PATTERN, cleaned_name, re.IGNORECASE)
    if match:
        if match.group(0) in [
            "SPLITS",
            "WAVS",
            "WAV",
            "MDE",
            "MIXES",
            "AUDIO",
            "STEMS",
            "5.1",
        ]:
            return "WAV"
        elif match.group(0) in ["GFX", "GFXPACKAGE", "GRAPHICS"]:
            return "GRFX"
        return match.group(0)
    return None


def get_content_type_d(cleaned_name: str) -> Optional[str]:
    return (
        re.search(DOC_PATTERN, cleaned_name, re.IGNORECASE).group(0)
        if re.search(DOC_PATTERN, cleaned_name, re.IGNORECASE)
        else None
    )


def get_content_type_misc(cleaned_name: str) -> Optional[str]:
    """
    Determine the miscellaneous content type from the cleaned name.
    """
    match_misc = re.finditer(MISC_VIDEO_PATTERN, cleaned_name, re.IGNORECASE)
    matches = [match.group(0) for match in match_misc if match is not None]
    logger.info(f"Matches for {cleaned_name}:  {matches}")

    return ",".join(matches) if matches else None


def set_document_info(df: pd.DataFrame, index: int, cleaned_name: str):
    df.at[index, "TITLETYPE"] = "document"
    df.at[index, "CONTENT_TYPE"] = "DOCX"
    df.at[index, "FILENAME"] = f"{cleaned_name}.docx"
    logger.info(f"{cleaned_name} TITLE TYPE: document, CONTENT TYPE: docx")


def set_video_info(
    df: pd.DataFrame, index: int, row: pd.Series, cleaned_name: str, content_type: str
):
    df.at[index, "TITLETYPE"] = "video"
    mediainfo = gmi.get_mediainfo(row, row["METAXML"])

    logger.info(f"MEDIA-INFO: {mediainfo}")

    if mediainfo:
        df.at[index, "CONTENT_TYPE"] = content_type
        df.at[index, "FRAMERATE"] = mediainfo["framerate"]
        df.at[index, "CODEC"] = mediainfo["codec"]
        df.at[index, "V_WIDTH"] = mediainfo["resolution"][0]
        df.at[index, "V_HEIGHT"] = mediainfo["resolution"][1]
        df.at[index, "DURATION_MS"] = mediainfo["duration_ms"]
        df.at[index, "FILENAME"] = mediainfo["filename"]
    else:
        df.at[index, "CONTENT_TYPE"] = content_type
        set_null_mediainfo(df, index)

    return


def set_archive_info(
    df: pd.DataFrame,
    index: int,
    cleaned_name: str,
    content_type_a: str,
    source_create_dt: str,
):
    df.at[index, "TITLETYPE"] = get_title_type(content_type_a)
    df.at[index, "CONTENT_TYPE"] = content_type_a
    df.at[index, "PROXY_COPIED"] = 3
    df.at[index, "FILENAME"] = (
        f"{cleaned_name}_{format_creation_date(source_create_dt)}.zip"
    )
    set_null_mediainfo(df, index)


def set_mixed_info(
    df: pd.DataFrame,
    index: int,
    cleaned_name: str,
    content_type_v: str,
    content_type_a: str,
    source_create_dt: str,
):
    df.at[index, "TITLETYPE"] = get_title_type(content_type_a)
    df.at[index, "CONTENT_TYPE"] = f"{content_type_a},{content_type_v}"
    df.at[index, "PROXY_COPIED"] = 3
    df.at[index, "FILENAME"] = (
        f"{cleaned_name}_{format_creation_date(source_create_dt)}.zip"
    )
    set_null_mediainfo(df, index)


def set_null_info(df: pd.DataFrame, index: int, cleaned_name: str):
    df.at[index, "TITLETYPE"] = "NULL"
    df.at[index, "CONTENT_TYPE"] = "NULL"
    df.at[index, "FILENAME"] = cleaned_name
    set_null_mediainfo(df, index)


def set_null_mediainfo(df: pd.DataFrame, index: int):
    df.at[index, "FRAMERATE"] = "NULL"
    df.at[index, "CODEC"] = "NULL"
    df.at[index, "V_WIDTH"] = "NULL"
    df.at[index, "V_HEIGHT"] = "NULL"
    df.at[index, "DURATION_MS"] = "NULL"


def clean_name(name: str) -> str:
    for key, value in ABBREVIATIONS.items():
        name = re.sub(rf"(?<![0-9A-Z])(?<=(-|_)){key}(?=(-|_|[1-5])?)", value, name)
    return name


def clean_metaxml(metaxml: str, name: str) -> str:
    """
    Clean the metaxml content by replacing problematic characters.

    Args:
        name (str): The name of the item being processed.
        xml_search (re.Match[str]): The regex match object containing problematic XML.
        r_metaxml (str): The replacement metaxml content.

    Returns:
        str: The cleaned metaxml content.
    """
    xml_pattern = re.compile(r"&[^;]+;")
    xml_search = xml_pattern.search(metaxml)

    if xml_search:
        # Replace characters in the matched XML string
        bad_xml = xml_search.group(0)
        metaxml_1 = bad_xml.replace("&", "and").replace("\\", "/")
        metaxml_2 = re.sub(r"&[^;]+;", "", metaxml_1)
        logger.info(f"metaxml for {name} was modified to remove illegal characters.")
    else:
        # Clean the replacement metaxml
        metaxml_2 = metaxml.replace("\\", "/")

    return metaxml_2


def get_traffic_code(cleaned_name: str) -> str:
    tcode = cleaned_name.split("_")
    if "_" in cleaned_name and len(tcode) >= 2:
        traffic_code = tcode[0]
    else:
        traffic_code = "UNKNOWN"
    return traffic_code


def get_title_type(content_type_a: str) -> str:
    return "graphic" if content_type_a == "GRFX" else "archive"


def format_creation_date(date_str: str) -> str:
    date_obj = pd.to_datetime(date_str)
    return date_obj.strftime("%Y%m%d")


if __name__ == "__main__":
    import sys

    date = sys.argv[1]
    parsed_csv = sys.argv[2] if len(sys.argv) > 2 else None
    csv_clean(date, parsed_csv)
