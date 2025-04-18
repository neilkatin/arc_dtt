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
import neil_tools.spreadsheet_tools as spreadsheet_tools

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

# flag field in vehicle structures
IN_AVIS = '__IN_AVIS__'

def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = neil_tools.init_config(config_static, ".env")

    account = init_o365(config, config.TOKEN_FILENAME_AVIS)

    for dr_id in args.dr_id:
        contents = fetch_dr_roster(config, dr_id)

        filename = f"staff-roster-{ dr_id }.xls"
        log.debug(f"saving roster to { filename }")
        with open(filename, "wb") as fd:
            fd.write(contents)



def fetch_dr_roster(config, dr_id, dr_config):
    """ get the most recent roster associated with the specified DR. """

    account = init_o365(config, config.TOKEN_FILENAME_AVIS)

    log.debug(f"dr_config { dr_config }  staffing_subject { dr_config.staffing_subject }")
    message_match_string = dr_config.staffing_subject
    if message_match_string == None:
        message_match_string = f"DR { dr_id } Automated Staffing Reports"
    attach_match_re = re.compile('^Staff Roster_.*')

    contents = search_mail(account, config.PROGRAM_EMAIL, message_match_string, attach_match_re)
    return contents

def convert_roster_to_objects(contents):

    # log.debug(f"about to open workbook.  contents { contents[0:8] }")

    wb = xlrd.open_workbook(file_contents=contents)
    ws = wb.sheet_by_index(0)

    nrows = ws.nrows
    ncols = ws.ncols

    title_row = 5

    assert nrows >= title_row + 2   # must be at least 2 rows: title row and at least one data entry

    # gather all the data
    matrix = []
    for index in range(title_row, nrows):
        matrix.append( ws.row_values(index) )

    objects = spreadsheet_tools.matrix_to_object_array(matrix)

    return objects


def search_mail(account, mailbox_email, message_match_string, attach_match_re):
    mailbox = account.mailbox(resource=mailbox_email)

    q = mailbox.new_query()
    q = q.order_by('sentDateTime', ascending=False)
    dt = datetime.datetime(1900, 1, 1)
    q = q.on_attribute('sentDateTime').greater_equal(dt)
    q = q.on_attribute('subject').contains(message_match_string)


    messages = mailbox.get_messages(query=q, limit=1, download_attachments=False)
    message = next(messages, None)

    if message is None:
        log.error(f"Failed to read any messages that match { q }")
        return None

    #log.debug(f"message { message } sent { message.sent }")

    message.attachments.download_attachments()
    attachments = message.attachments
    #log.debug(f"attachments len: { len(attachments) }")

    for attachment in attachments:
        #log.debug(f"attachment { attachment } size { attachment.size }")
        if attach_match_re.search(attachment.name) != None:
            content = base64.b64decode(attachment.content)
            log.debug(f"found a match: { attachment.name } size { attachment.size } len { len(content) }")
            return content

    return None



def init_o365(config, token_filename):
    """ do initial setup to get a handle on office 365 graph api """

    o365 = arc_o365.arc_o365.arc_o365(config, token_filename=token_filename, timezone="America/Los_Angeles")

    account = o365.get_account()
    if account is None:
        raise Exception("could not access office 365 graph api")

    return account



def parse_args():
    parser = argparse.ArgumentParser(
            description="tools to support Disaster Transportation Tools reporting",
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

