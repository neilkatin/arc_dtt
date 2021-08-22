#!/usr/bin/env python

import os
import re
import sys
import logging
import argparse
import datetime
import json
import io
import zipfile
import base64

import requests
import xlrd

import neil_tools
import message

import config as config_static

import arc_o365
#from O365_local.excel import WorkBook as o365_WorkBook
import O365
from O365.excel import WorkBook as o365_WorkBook



NOW = datetime.datetime.now().astimezone()
DATESTAMP = NOW.strftime("%Y-%m-%d")
TIMESTAMP = NOW.strftime("%Y-%m-%d %H:%M:%S %Z")
FILESTAMP = NOW.strftime("%Y-%m-%d %H-%M-%S %Z")
EMAILSTAMP = NOW.strftime("%Y-%m-%d %H-%M")


def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = neil_tools.init_config(config_static, ".env")

    for dr_id in args.dr_id:
        contents = message.fetch_dr_roster(config, dr_id)

        make_vcard_from_roster(contents)


def make_vcard_from_roster(contents):
    """ produce a vcard file from the roster """

    wb = xlrd.open_workbook(file_contents=contents)
    ws = wb.sheet_by_index(0)

    title = ws.cell_value(0, 0)
    dr = ws.cell_value(2, 0)

    log.debug(f"title '{ title }' dr '{ dr }'")




def parse_args():
    parser = argparse.ArgumentParser(
            description="generate an address book (in vcard format) from the current staff roster",
            allow_abbrev=False)
    parser.add_argument("--debug", help="turn on debugging output", action="store_true")
    parser.add_argument("--dr-id", help="the name of the DR (like 155-22)", required=True, action="append")

    args = parser.parse_args()

    return args


if __name__ == "__main__":
    neil_tools.init_logging(__name__)
    log = logging.getLogger(__name__)
    main()
else:
    log = logging.getLogger(__name__)

