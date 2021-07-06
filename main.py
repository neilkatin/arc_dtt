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

import openpyxl
import openpyxl.utils
import openpyxl.styles
import openpyxl.styles.colors
import openpyxl.writer.excel

import neil_tools
from neil_tools import spreadsheet_tools

import config as config_static
import web_session

import arc_o365
from O365_local.excel import WorkBook as o365_WorkBook



NOW = datetime.datetime.now().astimezone()
DATESTAMP = NOW.strftime("%Y-%m-%d")
TIMESTAMP = NOW.strftime("%Y-%m-%d %H:%M:%S %Z")
FILESTAMP = NOW.strftime("%Y-%m-%d %H-%M-%S %Z")


def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = neil_tools.init_config(config_static, ".env")

    if True:
        session = web_session.get_session(config)

        # ZZZ: should validate session here, and re-login if it isn't working...
        get_dr_list(config, session)

        log.debug(f"DR_ID { config.DR_ID } DR_NAME { config.DR_NAME }")

        # get people and vehicles from the DTT
        vehicles = get_vehicles(config, session)
        people = get_people(config, session)

    # fetch the avis report
    account = init_o365(config)

    if True:
        # fetch the avis spreadsheet
        output_bytes = fetch_avis(config, account, vehicles)

    store_report(config, account, output_bytes)


def fetch_avis(config, account, vehicles):
    """ fetch the latest avis vehicle report from sharepoint """

    storage = account.storage()
    drive = storage.get_drive(config.NHQDCSDLC_DRIVEID)
    fy21 = drive.get_item_by_path(config.FY21_ITEM_PATH)

    children = fy21.get_items()

    rental_re = re.compile('^ARC Open Rentals\s*-?\s*(\d{1,2})-(\d{1,2})-(\d{2,4})\.xlsx$')
    count = 0
    mismatch = 0
    newest_file_date = None
    newest_ref = None
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

            #log.debug(f"file match: { year }-{ month }-{ day }")

            file_date = datetime.date(int(year), int(month), int(day))
            if newest_file_date is None or newest_file_date < file_date:
                newest_file_date = file_date
                newest_ref = child

    log.debug(f"total items: { count } mismatched { mismatch }")

    if newest_ref is None:
        raise Exception("fetch_avis: no valid files found")
    log.debug(f"newest_file { newest_ref.name }")
    config['AVIS_FILE'] = newest_ref.name

    output_wb = openpyxl.Workbook()
    insert_overview(output_wb, config)

    output_ws_open = output_wb.create_sheet("Open RA")

    # we now have the latest file.  Suck out all the data
    workbook = o365_WorkBook(newest_ref, persist=False)
    avis_title, avis_open_columns, avis_open = read_avis_sheet(config, workbook.get_worksheet('Open RA'))
    #avis_closed = read_avis_sheet(workbook.get_worksheet('Closed RA'))

    output_columns = copy_avis_sheet(output_ws_open, avis_title, avis_open_columns, avis_open)
    match_avis_sheet(output_ws_open, output_columns, avis_open, vehicles)



    # cleanup: delete the default sheet name in the output workbook
    default_sheet_name = 'Sheet'
    if default_sheet_name in output_wb:
        del output_wb[default_sheet_name]

    # now save the workbook to sharepoint...
    # ZZZ: how?
    iobuffer = io.BytesIO()
    zipbuffer = zipfile.ZipFile(iobuffer, mode='w')
    writer = openpyxl.writer.excel.ExcelWriter(output_wb, zipbuffer)
    writer.save()

    bufferview = iobuffer.getbuffer()
    log.debug(f"output spreadsheet length: { len(bufferview) }")

    #with open("temp.xlsx", "wb") as fb:
    #    fb.write(bufferview)

    return bufferview


def copy_avis_sheet(ws, title, columns, rows):
    """ copy avis data to the output ws """
    wrap_alignment = openpyxl.styles.Alignment(wrapText=True, horizontal='center')

    fixed_width_columns = {
            'Rental Region Desc': 20,
            'Rental Distict Desc': 20,
            'CO Date': 16,
            'Rental Loc Desc': 22,
            'Address Line 1': 30,
            'Address Line 3': 30,
            'Full Name': 30,
            'Exp CI Date': 16,
            'AWD Orgn Buildup Desc': 30,
            }

    # first write out the title
    row = 1
    col = 1
    output_columns = {}
    for index, key in enumerate(title):
        if key != '' and key != 'CO Time' and key != 'Exp CI Time':
            cell = ws.cell(row=row, column=col, value=key)
            cell.alignment = wrap_alignment
            col_letter = openpyxl.utils.get_column_letter(col)

            if key in fixed_width_columns:
                ws.column_dimensions[col_letter].width = fixed_width_columns[key]
            else:
                ws.column_dimensions[col_letter].auto_size = True

            output_columns[key] = col
            col += 1

    # now add the data
    time_regex = re.compile('(\d{2}):(\d{2}):(\d{2})')
    row = 1
    for row_data in rows:
        row += 1
        col = 1
        for key, value in row_data.items():

            # ignore columns without a name and meta-entries
            if key == '' or key.startswith('__'):
                continue

            # fold time columns into their date columns
            time_key = None
            if key == 'CO Date':
                time_key = 'CO Time'
            elif key == 'Exp CI Date':
                time_key = 'Exp CI Time'

            # process date/time column pairs
            if time_key:

                dt = spreadsheet_tools.excel_to_dt(value)

                time_string = row_data[time_key]
                time_match = time_regex.match(time_string)
                if time_match:
                    interval = datetime.timedelta(hours=int(time_match.group(1)), minutes=int(time_match.group(2)), seconds=int(time_match.group(3)))
                    dt += interval
                else:
                    log.debug(f"copy_avis_sheet: row { row } time { time_key } didn't parse: '{ time_string }'")



                #log.debug(f"adding row { row } column { col } title { key } value { dt } time_string '{ time_string }'")
                cell = ws.cell(row=row, column=col, value=dt)
                cell.number_format = 'yyyy-mm-dd hh:mm'

            # don't output time columns: they were already handled
            elif key == 'CO Time' or key == 'Exp CI Time':
                # don't increment column counter and ignore time columns
                continue

            # copy over all other columns
            else:

                #log.debug(f"adding row { row } column { col } title { key } value { value }")
                ws.cell(row=row, column=col, value=value)

            col += 1

    last_col_letter = openpyxl.utils.get_column_letter(col -1)
    table_ref = f"A1:{last_col_letter}{row}"
    log.debug(f"last col letter { last_col_letter }, table_ref { table_ref }")
    table = openpyxl.worksheet.table.Table(displayName='AvisOpen', ref=table_ref)
    ws.add_table(table)

    return output_columns



def make_index(vehicles, first_field, second_field=None, cleanup=False):

    #log.debug(f"make_index: vehicles len { len(vehicles) }, 1st_field { first_field } 2nd_field { second_field }")
    result = {}
    for row in vehicles:
        vehicle = row['Vehicle']
        if first_field not in vehicle:
            continue

        if second_field is not None and second_field not in vehicle:
            continue

        key = vehicle[first_field]
        if key is None:
            continue

        if cleanup:
            # the emailed reservation number looks like 12345678-US-6; put it in cannonical form
            key = key.replace('-', '').upper()

        if second_field is not None:
            second = vehicle[second_field]
            if second is None:
                continue
            key = key + " " + vehicle[second_field]

        #log.debug(f"make_index: key { key }")
        result[key] = row

    return result


def get_dtt_id(vehicle_dict, index):
    """ utility function to look up a field in the vehicle object from the DTT """

    if index not in vehicle_dict:
        return None

    #log.debug(f"index { index } v_dict { vehicle_dict[index] }")
    return vehicle_dict[index]['DisasterVehicleID']



def mark_cell(ws, fill, row_num, col_map, col_name):
    """ apply a fill to a particular cell """

    col_num = col_map[col_name]

    cell = ws.cell(row=row_num, column=col_num)
    cell.fill = fill


def match_avis_sheet(ws, columns, avis, vehicles):
    """ match entries from the DTT to entries in the Avis report. """
    fill_red = openpyxl.styles.PatternFill(fgColor="FFC0C0", fill_type = "solid")
    fill_green = openpyxl.styles.PatternFill(fgColor="C0FFC0", fill_type = "solid")
    fill_yellow = openpyxl.styles.PatternFill(fgColor="FFFFC0", fill_type = "solid")
    fill_blue = openpyxl.styles.PatternFill(fgColor="8080FF", fill_type = "solid")

    # generate different index for the vehicles
    v_ra = make_index(vehicles, 'RentalAgreementNumber')
    v_res= make_index(vehicles, 'RentalAgreementReservationNumber', cleanup=True)
    v_key = make_index(vehicles, 'KeyNumber')
    v_plate = make_index(vehicles, 'PlateState', 'Plate')

    #log.debug(f"plate keys: { v_plate.keys() }")

    spreadsheet_row = 1
    for row in avis:
        spreadsheet_row += 1

        ra = row['Rental Agreement No']
        res = row['Reservation No']
        key = row['MVA No']
        plate = row['License Plate State Code'] + ' ' + row['License Plate Number']

        # trim the leading zero from key; the Avis Report has 9 char key numbers
        if key.startswith('0') and len(key) > 1:
            key = key[1:]

        # use DisasterVehicleID as the DTT identity for a vehicle
        ra_id = get_dtt_id(v_ra, ra)
        res_id = get_dtt_id(v_res, res)
        key_id = get_dtt_id(v_key, key)
        plate_id = get_dtt_id(v_plate, plate)

        #log.debug(f"ra { ra_id } res { res_id } key { key_id } plate { plate_id }; raw { ra } { res } { key } { plate }")

        if ra_id == None and res_id == None and key_id == None and plate_id == None:
            # vehicle doesn't appear in the DTT at all; color it blue
            mark_cell(ws, fill_blue, spreadsheet_row, columns, 'Rental Agreement No')
            mark_cell(ws, fill_blue, spreadsheet_row, columns, 'Reservation No')
            mark_cell(ws, fill_blue, spreadsheet_row, columns, 'MVA No')
            mark_cell(ws, fill_blue, spreadsheet_row, columns, 'License Plate State Code')
            mark_cell(ws, fill_blue, spreadsheet_row, columns, 'License Plate Number')

        elif ra_id == res_id and ra_id == key_id and ra_id == plate_id:
            # all columns match: color it green
            mark_cell(ws, fill_green, spreadsheet_row, columns, 'Rental Agreement No')
            mark_cell(ws, fill_green, spreadsheet_row, columns, 'Reservation No')
            mark_cell(ws, fill_green, spreadsheet_row, columns, 'MVA No')
            mark_cell(ws, fill_green, spreadsheet_row, columns, 'License Plate State Code')
            mark_cell(ws, fill_green, spreadsheet_row, columns, 'License Plate Number')

        else:
            # else color yellow if value is found; red if value not found
            mark_cell(ws, fill_red if ra_id is None else fill_yellow, spreadsheet_row, columns, 'Rental Agreement No')
            mark_cell(ws, fill_red if res_id is None else fill_yellow, spreadsheet_row, columns, 'Reservation No')
            mark_cell(ws, fill_red if key_id is None else fill_yellow, spreadsheet_row, columns, 'MVA No')
            mark_cell(ws, fill_red if plate_id is None else fill_yellow, spreadsheet_row, columns, 'License Plate State Code')
            mark_cell(ws, fill_red if plate_id is None else fill_yellow, spreadsheet_row, columns, 'License Plate Number')



        





def insert_overview(wb, config):
    """ insert an overview (documentation) sheet in the workbook """
    ws = wb.create_sheet('Overview', 0)

    ws.column_dimensions['A'].width = 120       # hard coded width....
    ws.row_dimensions[1].height = 1000          # hard coded height....

    doc_string = f"""
This document has vehicles from the daily Avis report for this DR ({ config.DR_NUM.rjust(3, '0') }-{ config.DR_YEAR }).

On the Open RA (rental agreement) sheet there should be one row per vehicle marked as assigned to the DR
(from the 'Cost Control No' column).

Rows in the Open RA sheet have been checked against the DTT Vehicles data.  These columns are checked between
the two reports: License Plate State Code/License Plate Number, Reservation No, Rental Agreement No, Reservation No.

If all four fields match: the cells will be marked Green.

If a vehicle in the Avis report is not found in the DTT: the cells will be blue.

If all four fields don't match: any field that is not found in the DTT will be red.  Otherwise the fields
will be yellow.

This file based on the { config.AVIS_FILE } file
This file generated at { TIMESTAMP }

"""

    cell = ws.cell(row=1, column=1, value=doc_string)
    cell.alignment = openpyxl.styles.Alignment(wrapText=True, vertical='top')



def read_avis_sheet(config, sheet):

    log.debug(f"sheet name { sheet.name }")
    sheet_range = sheet.get_used_range()

    #log.debug(f"last_column { sheet_range.get_last_column() } last_row { sheet_range.get_last_row() }")

    values = sheet_range.values

    values.pop(0)               # sheet title text -- before the 'data table'

    #title_row = values.pop(0)   # column headers
    #log.debug(f"title_row { title_row }")

    title_row = values[0]
    avis_columns = spreadsheet_tools.title_to_dict(title_row)
    avis_all = spreadsheet_tools.matrix_to_object_array(values)

    # the DR number format (in the 'Cost Control No' column) isn't well 
    dr_column = 'Cost Control No'

    # trying to match patterns like:
    # 98, 098, DR098, DR098-21, DR098-2021, 098-21, etc...
    dr_regex = re.compile(f"(dr)?0*{ config.DR_NUM }(-(20)?{ config.DR_YEAR })?")

    #avis_dr = list(filter(lambda x: dr_regex.match(x), avis_all))
    f_result = filter(lambda row: dr_regex.match(row[dr_column]), avis_all)
    avis_dr = list(f_result)
    
    return title_row, avis_columns, avis_dr




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

    log.debug(f"r.status { r.status_code } r.reason { r.reason } r.url { r.url } r.content_type { r.headers['content-type'] } data rows { len(data) }")

    #log.debug(f"json { data }")
    #log.debug(f"Returned data\n{ json.dumps(data, indent=2, sort_keys=True) }")

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


def store_report(config, account, wb_bytes):
    """ store the report.  for now: stash it in the Transportation/Reports file store. """

    log.debug("called")
    storage = account.storage()
    drive = storage.get_drive(config.TRANS_REPORTS_DRIVEID)
    reports = drive.get_item_by_path(config.TRANS_REPORTS_FOLDER_PATH)

    stream = io.BytesIO(wb_bytes)
    stream.seek(0, io.SEEK_END)
    stream_size = stream.tell()
    stream.seek(0, io.SEEK_SET)

    log.debug(f"stream_size is { stream_size }")

    item_name = f"DR{ config.DR_NUM.rjust(3, '0') }-{ config.DR_YEAR } Avis Report { FILESTAMP }.xlsx"

    result = reports.upload_file(item_name, item_name=item_name, stream=stream, stream_size=stream_size, conflict_handling='replace')
    log.debug(f"after upload: result { result }")



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

