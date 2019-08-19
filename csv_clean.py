#! /usr/bin/env python3

import logging
import os
import re

import config as cfg
import pandas as pd
import xml.etree.ElementTree as ET

import get_mediainfo as gmi

logger = logging.getLogger(__name__)


def csv_clean(date):
    """
    Parse out data from the db, insert into new fields,  and assign values for these fields in each row.
    """

    config = cfg.get_config()

    rootpath = config['paths']['rootpath']

    os.chdir(rootpath + "_CSV_Exports/")

    parsed_csv = date + "_" + "gor_diva_merged_parsed.csv"
    clean_csv = date + "_" + "gor_diva_merged_cleaned.csv"

    clean_1_msg = f"START GORILLA-DIVA DB CLEAN"
    logger.info(clean_1_msg)

    try:
        pd_reader = pd.read_csv(parsed_csv, header=0)
        df = pd.DataFrame(pd_reader)
        df = df.astype({"METAXML": str})

        df.insert(13, "TITLETYPE", 'NULL', allow_duplicates=True)
        df.insert(14, "FRAMERATE", 'NULL', allow_duplicates=True)
        df.insert(15, "CODEC", 'NULL', allow_duplicates=True)
        df.insert(16, "V_WIDTH", 'NULL', allow_duplicates=True)
        df.insert(17, "V_HEIGHT", 'NULL', allow_duplicates=True)
        df.insert(18, "TRAFFIC_CODE", 'NULL', allow_duplicates=True)
        df.insert(19, "DURATION_MS", 'NULL', allow_duplicates=True)
        df.insert(20, "XML_CREATED", 0, allow_duplicates=True)

        df.to_csv(clean_csv)

        for index, row in df.iterrows():

            name = str(row['NAME']).upper()
            name_clean = clean_name(name)
            df.at[index, 'NAME'] = name_clean
            print(str(index) + "    " + name_clean)

            traffic_code = get_traffic_code(name_clean)
            df.at[index, 'TRAFFIC_CODE'] = traffic_code

            df_row = df.loc[index]

            if pd.isnull(df_row['METAXML']) is not True:
                l_metaxml = df_row['METAXML']
                r_metaxml = r'{}'.format(l_metaxml)
                metaxml = clean_metaxml(r_metaxml, name_clean)
                df.at[index, 'METAXML'] = metaxml

            else:
                df.at[index, 'METAXML'] = 'NULL'
                metaxml = df.at[index, 'METAXML']

            if row['_merge'] is not "both":
                df.drop(df.index)

            video_check = re.search(r'([_]VM)|([_]EM)|([_]UHD)', name_clean)
            archive_check = re.search(
                r'([_]AVP)|([_]PPRO)|([_]FCP)|([_]PTS)|([_]AVP)|([_]GRFX)|([_]GFX)|([_]WAV)', name_clean)

            if (
                video_check is not None
                and archive_check is None
                # and metaxml is not 'NULL'
            ):
                df.at[index, 'TITLETYPE'] = 'video'
                content_type = re.sub('(_|-)', '', video_check.group(0))
                mediainfo = gmi.get_mediainfo(df_row, metaxml)

            # elif (video_check is not None
            #       and archive_check is None
            #       and metaxml is 'NULL'):
            #     df.at[index, 'TITLETYPE'] = 'video'
            #     content_type = re.sub('_', '', video_check.group(0))

            elif (video_check is None
                  and archive_check is not None):
                df.at[index, 'TITLETYPE'] = 'archive'
                content_type = re.sub('_', '', archive_check.group(0))
                mediainfo = ['NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', ]

            elif (video_check is not None
                  and archive_check is not None):
                df.at[index, 'TITLETYPE'] = 'archive'
                content_type = re.sub('_', '', archive_check.group(0))
                mediainfo = ['NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', ]

            else:
                df.at[index, 'TITLETYPE'] = 'NULL'
                clean_2_msg = f"TITLETYPE for {name} is NULL. "
                logger.info(clean_2_msg)
                content_type = 'NULL'
                mediainfo = ['NULL', 'NULL', 'NULL', 'NULL', 'NULL', 'NULL', ]

            df.at[index, 'CONTENT_TYPE'] = content_type
            df.at[index, 'FRAMERATE'] = mediainfo[0]
            df.at[index, 'CODEC'] = mediainfo[1]
            df.at[index, 'V_WIDTH'] = mediainfo[2]
            df.at[index, 'V_HEIGHT'] = mediainfo[3]
            df.at[index, 'DURATION_MS'] = mediainfo[4]

        df.to_csv(clean_csv)

        clean_3_msg = f"GORILLA-DIVA DB CLEAN COMPLETE"
        logger.info(clean_3_msg)

        os.chdir(rootpath)

        return clean_csv

    except Exception as e:
        db_clean_excp_msg = f"\n\
        Exception raised on the Gor-Diva DB Clean.\n\
        Error Message:  {str(e)} \n\
        Index Count: {index}\n\
        "
        logger.exception(db_clean_excp_msg)


def get_traffic_code(name_clean):
    """
    Validate the the traffic code begins with a 0, and contains the correct number of characters.
    """

    name = str(name_clean)

    if name.startswith('0') is not True:
        code_search = re.search(r"(_|-)?0[0-9]{5}(_|-)?", name)

        if code_search is not None:
            traffic_match = code_search.group(0)
            traffic_code_r = re.sub(r'(_|-)', '', traffic_match)
            traffic_code = "=\"" + traffic_code_r + "\""
        else:
            err_msg = f"Incompatible file ID - {str(name)}. traffic_code set to NULL"
            logger.error(err_msg)
            traffic_code = 'NULL'
    else:
        traffic_code = "=\"" + name[:6] + "\""

    return traffic_code


def clean_metaxml(r_metaxml, name):

    xml_search = re.search(r"[<FileName>].*&.*[</FileName>]", r_metaxml)

    if xml_search is not None:
        bad_xml = xml_search.group(0)
        good_xml = bad_xml.replace("&", "and")
        metaxml = good_xml.replace("\\", "/")
        clean_xml_msg = f"metaxml for {name} was modified to remove '&' characters."
        logger.info(clean_xml_msg)
    else:
        metaxml = r_metaxml.replace("\\", "/")

    return metaxml


def clean_name(name):

    name_search = re.search(r".*&.*", name)

    if name_search is not None:

        bad_name = name_search.group(0)
        good_name = bad_name.replace("&", "and")
        name_clean = good_name
        clean_name_msg = f"metaxml for {name} was modified to remove '&' characters."
        logger.info(clean_name_msg)
    else:
        name_clean = name

    return name_clean


if __name__ == '__main__':
    csv_clean('201908190949')
