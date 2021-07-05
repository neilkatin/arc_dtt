#!/usr/bin/env python

import os
import re
import sys
import logging
import argparse
import datetime
import json

#import openpyxl
#import openpyxl.utils
#import openpyxl.styles
#import openpyxl.styles.colors

import neil_tools

import config as config_static
import web_session

import arc_o365



DATESTAMP = datetime.datetime.now().strftime("%Y-%m-%d")


def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = neil_tools.init_config(config_static, ".env")

    #session = web_session.get_session(config)

    # ZZZ: should validate session here, and re-login if it isn't working...
    #get_dr_list(config, session)

    #log.debug(f"DR_ID { config.DR_ID } DR_NAME { config.DR_NAME }")

    # get people and vehicles from the DTT
    #vehicles = get_vehicles(config, session)
    #people = get_people(config, session)

    # fetch the avis report
    account = init_o365(config)

    # fetch the avis spreadsheet
    fetch_avis(config, account)


def fetch_avis(config, account):
    """ fetch the latest avis vehicle report from sharepoint """

    #sharepoint = account.sharepoint()
    #site = sharepoint.get_site('americanredcross.sharepoint.com', '/sites/NHQDCSDLC')
    # hostname, site_collection_id, site_id
    #site = sharepoint.get_site(config.SITE_ID)

    #storage = account.storage()
    #item = storage.get_site_path('4e1787c4-bf1b-4828-876a-6d7b1613ddec', 'Shared Documents/Gray Sky/Avis Reports')

    #document_library = 


    storage = account.storage()
    drive = storage.get_drive(config.NHQDCSDLC_DRIVE_ID)
    fy21 = drive.get_item_by_path(config.FY21_ITEM_PATH)

    children = fy21.get_items()

    rental_re = re.compile('^ARC Open Rentals\s*-?\s*(\d{1,2})-(\d{1,2})-(\d{2,4})\.xlsx$')
    count = 0
    mismatch = 0
    newest_file_date = None
    newest_file = None
    for child in children:

        #log.debug(f"child { child.name } id { child.object_id }")

        # humans made the name, so there is good (but not perfect) adherence to a naming standard.
        # its something like "ARC Open Rentals - mm-dd-yyyy.xlsx", but there is variation...
        match = rental_re.match(child.name)
        count += 1
        if match is None:
            log.info(f"no pattern match for '{ child.name }'")
            mismatch += 1
        else:
            month = match.group(1).lstrip('0')
            day = match.group(2).lstrip('0')
            year = match.group(3).lstrip('0')

            log.debug(f"file match: { year }-{ month }-{ day }")

            file_date = datetime.date(int(year), int(month), int(day))
            if newest_file_date is None or newest_file_date < file_date:
                newest_file_date = file_date
                newest_file = child

    log.debug(f"total items: { count } mismatched { mismatch }")

    if newest_file is None:
        raise Exception("fetch_avis: no valid files found")
    log.debug(f"newest_file { newest_file.name }")

    # we now have the latest file.  Suck out all the data



def init_o365(config):
    """ do initial setup to get a handle on office 365 graph api """
    o365 = arc_o365.arc_o365.arc_o365(config)
    account = o365.get_account()
    if account is None:
        raise Exception("could not access office 365 graph api")

    return account

    

def get_people(config, session):
    """ Retrieve the people list from the DTT (as a json list) """

    data = get_json(config, session, 'People/Details')
    return data

def get_vehicles(config, session):
    """ Retrieve the vehicle list from the DTT (as a json list) """

    data = get_json(config, session, 'Vehicles')
    return data


def get_json(config, session, api_type):

    url = config.DTT_URL + f"api/Disaster/{ config.DR_ID }/" + api_type

    r = session.get(url)
    r.raise_for_status()

    data = r.json()

    log.debug(f"r.status { r.status_code } r.reason { r.reason } r.url { r.url } r.content_type { r.headers['content-type'] }")

    #log.debug(f"json { data }")
    log.debug(f"Returned data\n{ json.dumps(data, indent=2, sort_keys=True) }")

    return data

def get_dr_list(config, session):
    """ Get the list of DRs we have access to from the DTT """

    url = config.DTT_URL + "Vehicles"

    r = session.get(url)
    r.raise_for_status()

    codes = r.html.find('#DisasterCodes', first=True)

    options = codes.find('option')

    for option in options:
        value = option.attrs['value']
        text = option.text
        #log.debug(f"option value { value } text { text }")

    # for now, while we only have access to one DR: pick the last value
    config['DR_ID'] = value
    config['DR_NAME'] = text



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

