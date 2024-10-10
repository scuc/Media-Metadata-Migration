#! /usr/bin/env python3

import logging
import re
from xml.dom import minidom
import xml.etree.ElementTree as ET

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)


def get_mediainfo(df_row, metaxml):
    """
    Extract mediainfo from the metaxml field of the DB.
    """
    log_info(df_row, metaxml)

    if df_row["METAXML"] != "nan" and len(metaxml) != 0:
        try:
            metaxml = sanitize_xml_content(df_row["NAME"], metaxml)
            mediainfo = extract_from_metaxml(df_row, metaxml)
        except Exception as e:
            log_exception("Exception raised on 1st block of get_mediainfo.", df_row, e)
            mediainfo = get_estimated_mediainfo(df_row)
    else:
        try:
            mediainfo = get_estimated_mediainfo(df_row)
        except Exception as e:
            log_exception("Exception raised on 2nd block of get_mediainfo.", df_row, e)
            mediainfo = None

    return mediainfo


def log_info(df_row, metaxml):
    """
    Log information messages.
    """
    logger.info(f"Getting mediainfo for {df_row['NAME']}")
    logger.info(f"MetaXML has length = {len(metaxml)}")


def extract_from_metaxml(df_row, metaxml):
    """
    Extract mediainfo from the metaxml.
    """
    parser = ET.XMLParser(encoding="utf-8")
    tree = ET.ElementTree(ET.fromstring(metaxml, parser=parser))
    root = tree.getroot()

    # pretty_xml = prettify_xml(root)
    # logger.info(f"Pretty XML: {pretty_xml}")

    def safe_find_text(path):
        element = root.find(path)
        return element.text if element is not None else "NULL"

    mediainfo = {
        "framerate": safe_find_text("VideoTrack/Video/AverageFrameRate"),
        "codec": safe_find_text("VideoTrack/Video/Format"),
        "resolution": adjust_resolution(
            safe_find_text("VideoTrack/Video/Width"),
            safe_find_text("VideoTrack/Video/Height"),
        ),
        "duration_ms": safe_find_text("DurationInMs"),
        "filename": adjust_filename(safe_find_text("FileName")),
    }

    return mediainfo


def adjust_resolution(v_width, v_height):
    """
    Adjust resolution based on width and height.
    """
    if v_height == "1062" and v_width == "1888":
        v_width, v_height = "1920", "1080"
    elif v_height == "360" and v_width == "640":
        v_width, v_height = "1920", "1080"

    return v_width, v_height


def adjust_filename(filename):
    """
    Adjust filename based on the prefix.
    """
    if filename.startswith("NLE."):
        return filename[4:]
    return filename[7:]


def get_estimated_mediainfo(df_row):
    """
    Estimate mediainfo when metaxml is not available.
    """
    codec, codec_value = get_codec(df_row)
    v_width, v_height = est_resolution(df_row, codec_value)
    framerate = get_framerate(df_row)

    if v_height == "360" and v_width == "640" and codec == "ProRes":
        v_width, v_height = "1920", "1080"
        logger.info(
            f"{df_row['GUID']} - {df_row['NAME']} - filesize: {df_row['FILESIZE']} - Estimating file is HD:1920x1080."
        )

    duration = (
        int(df_row["CONTENTLENGTH"]) * 1000 if int(df_row["CONTENTLENGTH"]) != 0 else 0
    )
    filename = get_estimated_filename(df_row, codec)

    return {
        "framerate": framerate,
        "codec": codec,
        "resolution": (v_width, v_height),
        "duration_ms": duration,
        "filename": filename,
    }


def get_estimated_filename(df_row, codec):
    """
    Estimate filename based on codec and creation date.
    """
    date = df_row["SOURCECREATEDT"]
    creation_date = date.translate({ord(i): None for i in "- :"})
    if codec.upper() == "PRORES":
        return f"{df_row['NAME']}_{creation_date}.mov"
    if df_row["NAME"].upper().endswith("_MXF"):
        return f"{df_row['NAME']}_{creation_date}.mxf"
    return f"{df_row['NAME']}_{creation_date}.mov"


def get_codec(df_row):
    """
    Match the codec of a file using the info in the filename.
    """
    doc_pattern = r"(((?<![A-Z])|(?<=(-|_)))(UHD|XAVC|UHD|PRORES|XDCAM|DNX|IMX50|DV100)(?=(-|_|HQ|HD)?))"
    codec_match = re.search(doc_pattern, df_row["NAME"], re.IGNORECASE)

    codec_value = codec_match.group(0) if codec_match else "NULL"
    codec = (
        "VC-3"
        if codec_value == "DNXHD"
        else "XAVC" if codec_value == "UHD" else codec_value
    )
    logger.info(
        f'{df_row["GUID"]} - {df_row["NAME"]} - Estimating Codec based on filename'
    )

    return codec, codec_value


def get_framerate(df_row):
    """
    Match the framerate of a file using the info in the filename.
    """
    framerate_match = re.search(
        r"(?<![0-9A-Z])(?:23(?:\.?98|\.?976|\.?97|\.?94)?|25|29(?:\.?97)?|59(?:\.?94)?|NTSC|PAL|24P|720P)(?=[IPip]?)(?=(-|_)?)",
        df_row["NAME"][6:],
    )

    framerate_value = framerate_match.group(0) if framerate_match else "00"
    framerate_map = {
        "23": "23.98",
        "23976": "23.976",
        "2398": "23.98",
        "24P": "24",
        "25": "25",
        "29": "29.97",
        "2997": "29.97",
        "59": "59.94",
        "5994": "59.94",
        "720P": "59.94",
        "NTSC": "29.97",
        "PAL": "25",
    }
    framerate = framerate_map.get(framerate_value, framerate_value)
    logger.info(
        f"{df_row['GUID']} - {df_row['NAME']} - Framerate {framerate} value based on filename."
    )

    return framerate


def est_resolution(df_row, codec_value):
    """
    Estimate the resolution based on filesize and codec.
    """
    resolution_match = re.search(
        r"(?:525|625|720|1080|1080|2160)(?=[IPipKk]?)(?=(-|_)?)",
        df_row["NAME"],
    )

    resolution_map = {
        "525": ["720", "486"],
        "625": ["720", "576"],
        "720": ["1280", "720"],
        "1080": ["1920", "1080"],
        "2160": ["3840", "2160"],
        "IMX50": ["720", "486"],
    }

    filesize = int(df_row["FILESIZE"])
    if (
        18000000000 < filesize < 200000000000
        and codec_value not in ["XAVC", "UHD"]
        and df_row["CONTENTLENGTH"] != 0
    ):
        v_width, v_height = "1920", "1080"
        logger.info(
            f"{df_row['GUID']} - {df_row['NAME']} - filesize: {df_row['FILESIZE']} - Estimating file is HD: 1920x1080."
        )
    elif codec_value in ["XAVC", "UHD"]:
        v_width, v_height = "3840", "2160"
        logger.info(
            f"{df_row['GUID']} - {df_row['NAME']} - filesize: {df_row['FILESIZE']} - Estimating file is UHD:3840x2160."
        )
    elif (
        codec_value not in ["XAVC", "UHD"]
        and "_xdcam_" in df_row["NAME"].lower()
        or "_xdcamhd_" in df_row["NAME"].lower()
    ):
        v_width, v_height = "1920", "1080"

    elif resolution_match and codec_value not in ["XAVC", "UHD"]:
        resolution_value = resolution_match.group(0)
        v_width, v_height = resolution_map.get(resolution_value, ["00", "00"])

    else:
        v_width, v_height = "NULL", "NULL"
        logger.info(
            f"{df_row['GUID']} - {df_row['NAME']} - Cannot determine v_width or v_height, setting to Null."
        )

    print(v_height, v_width)
    return v_width, v_height


def log_exception(msg, df_row, e):
    """
    Log exception messages.
    """
    logger.exception(
        f"{msg}\nGUID: {df_row['GUID']}\nNAME: {df_row['NAME']}\nERROR: {str(e)}\n"
    )


def prettify_xml(elem):
    """Return a pretty-printed XML string for the Element."""

    rough_string = ET.tostring(elem, "utf-8")
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="  ")


def sanitize_xml_content(name: str, metaxml: str) -> str:
    """
    Replaces illegal characters in the provided XML content with valid XML entities and
    removes invalid control characters.
    """
    try:
        ET.fromstring(metaxml)  # Parses the XML string
    except ET.ParseError:
        logger.error(f"Invalid XML content found in: {name}")
        # Replace the five predefined XML entities
        metaxml = metaxml.replace("&", "And")  # Must be done first
        # metaxml = metaxml.replace("<", "&lt;")
        # metaxml = metaxml.replace(">", "&gt;")
        # metaxml = metaxml.replace('"', "&quot;")
        # metaxml = metaxml.replace("'", "&apos;")

        # Remove any control characters that are not allowed
        # Valid characters: U+0009 (tab), U+000A (line feed), U+000D (carriage return)
        # illegal_chars_pattern = r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]"
        # metaxml = re.sub(illegal_chars_pattern, "", metaxml)

        logger.info(f"XML content for {name} was sanitized.")
        logger.info(f"Sanitized XML: {metaxml}")

    return metaxml


if __name__ == "__main__":
    df_row_example = {
        "NAME": "051637_AIRPORTSECURITY2.5_EPISODE1080_2997I_PRORESHQ_TEXTLESS_CCD_091817",
        "METAXML": "<xml></xml>",
        "GUID": "example_guid",
        "FILESIZE": "20000000000",
        "CONTENTLENGTH": "1000",
        "SOURCECREATEDT": "2022-01-01 12:00:00",
    }
    get_mediainfo(df_row_example, df_row_example["METAXML"])
