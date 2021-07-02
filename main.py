#!/usr/bin/env python

import os
import re
import sys
import logging
import argparse
import datetime

#import openpyxl
#import openpyxl.utils
#import openpyxl.styles
#import openpyxl.styles.colors

import neil_tools

import config as config_static
import web_session

#import arc_o365.arc_o365



DATESTAMP = datetime.datetime.now().strftime("%Y-%m-%d")


def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = neil_tools.init_config(config_static, ".env")

    session = web_session.get_session(config)






def parse_args():
    parser = argparse.ArgumentParser(
            description="tools to support Disaster Transportation Tools reporting",
            allow_abbrev=False)
    parser.add_argument("--debug", help="turn on debugging output", action="store_true")

    #group = parser.add_mutually_exclusive_group(required=True)
    #group.add_argument("-p", "--prod", "--production", help="use production settings", action="store_true")
    #group.add_argument("-d", "--dev", "--development", help="use development settings", action="store_true")

    args = parser.parse_args()
    return args


if __name__ == "__main__":
    neil_tools.init_logging(__name__)
    log = logging.getLogger(__name__)
    main()
else:
    log = logging.getLogger(__name__)

