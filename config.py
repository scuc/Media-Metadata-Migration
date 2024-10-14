#! /usr/bin/env python3

import os
from turtle import clear

import yaml


def get_config():
    """
    Setup configuration and credentials
    """
    path = "/Users/cucos001/GitHub/Media-Metadata-Migration/config.yaml"

    with open(path, "rt") as f:
        config = yaml.safe_load(f.read())

    return config


def ensure_dirs():

    dirs = ["_CSV_Exports", "_logs", "_xml"]

    for dir in dirs:
        if not os.path.exists(dir):
            os.makedirs(dir)

    return
