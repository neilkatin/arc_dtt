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

# flag field in vehicle structures
IN_AVIS = '__IN_AVIS__'

# flag fields in avis structures
MISSING_AVIS_OPEN = '__MISSING_AVIS_OPEN__'
MISSING_AVIS_ALL = '__MISSING_AVIS_ALL__'

FILL_RED = openpyxl.styles.PatternFill(fgColor="FFC0C0", fill_type = "solid")
FILL_GREEN = openpyxl.styles.PatternFill(fgColor="C0FFC0", fill_type = "solid")
FILL_YELLOW = openpyxl.styles.PatternFill(fgColor="FFFFC0", fill_type = "solid")
FILL_BLUE = openpyxl.styles.PatternFill(fgColor="5BB1CD", fill_type = "solid")
FILL_CYAN = openpyxl.styles.PatternFill(fgColor="A0FFFF", fill_type = "solid")

def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = neil_tools.init_config(config_static, ".env")

    # fetch from DTT
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

    # avis report
    if not args.ignore_avis:
        # fetch the avis spreadsheet
        output_bytes = make_avis(config, account, vehicles)

        if args.store:
            item_name = f"DR{ config.DR_NUM.rjust(3, '0') }-{ config.DR_YEAR } Avis Report { FILESTAMP }.xlsx"
            store_report(config, account, item_name, output_bytes)
        else:
            # save a local copy instead
            file_name = 'avis.xlsx'
            log.debug(f"storing avis report to { file_name }")
            with open(file_name, "wb") as fb:
                fb.write(output_bytes)

    # group vehicle report
    if not args.ignore_group:
        # generate the group report
        output_bytes = make_group_report(config, vehicles)

        if args.store:
            item_name = f"DR{ config.DR_NUM.rjust(3, '0') }-{ config.DR_YEAR } Vehicles by GAP { FILESTAMP }.xlsx"
            store_report(config, account, item_name, output_bytes)

        else:
            # save a local copy instead
            file_name = 'gap.xlsx'
            log.debug(f"storing gap report to { file_name }")
            with open(file_name, "wb") as fb:
                fb.write(output_bytes)

def fetch_avis(config, account):
    """ get the most recent avis workbook """

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
        raise Exception("make_avis: no valid files found")
    log.debug(f"newest_file { newest_ref.name }")
    config['AVIS_FILE'] = newest_ref.name

    workbook = o365_WorkBook(newest_ref, persist=False)

    return workbook

def make_avis(config, account, vehicles):
    """ fetch the latest avis vehicle report from sharepoint """

    workbook = fetch_avis(config, account)

    output_wb = openpyxl.Workbook()
    insert_avis_overview(output_wb, config)


    output_ws_open = output_wb.create_sheet("Open RA")

    # we now have the latest file.  Suck out all the data
    avis_title, avis_open_columns, avis_open, avis_all = read_avis_sheet(config, workbook.get_worksheet('Open RA'))
    add_missing_avis_vehicles(vehicles, avis_all, avis_open)

    # generate the 'Open RA' sheet
    output_columns = copy_avis_sheet(output_ws_open, avis_open_columns, avis_title, avis_open)
    match_avis_sheet(output_ws_open, output_columns, avis_open, vehicles)

    # now serialize the workbook
    bufferview = workbook_to_buffer(output_wb)

    return bufferview



def workbook_to_buffer(wb):
    """ serialize a workbook to a byte array so it can be saved (either to network or locally) """

    # cleanup: delete the default sheet name in the output workbook
    default_sheet_name = 'Sheet'
    if default_sheet_name in wb:
        del wb[default_sheet_name]


    iobuffer = io.BytesIO()
    zipbuffer = zipfile.ZipFile(iobuffer, mode='w')
    writer = openpyxl.writer.excel.ExcelWriter(wb, zipbuffer)
    writer.save()

    bufferview = iobuffer.getbuffer()
    log.debug(f"output spreadsheet length: { len(bufferview) }")

    return bufferview



def copy_avis_sheet(ws, columns, title, rows):
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
        if key != '' and key != 'CO Time' and key != 'Exp CI Time' and not key.startswith('__'):
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
        for key in output_columns:


            # ignore columns without a name and meta-entries
            if key == '' or key.startswith('__'):
                continue

            if key not in row_data:
                col += 1
                continue

            value = row_data[key]

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



def cleanup_v_field(vehicle, field_name):
    """ do field cleanup for vehicle entries """

    if field_name not in vehicle:
        return None

    value = vehicle[field_name]

    if value is None:
        return None

    if field_name == 'RentalAgreementNumber':
        # rental agreements always start with 'U'; add it if not present
        if value[0] != 'U':
            value = 'U' + value

    elif field_name == 'RentalAgreementReservationNumber':
        # the emailed reservation number looks like 12345678-US-6; put it in cannonical form
        value = value.replace('-', '').upper()

    elif field_name == 'KeyNumber':
        # make sure there are 9 digits in the number
        value = value.rjust(9, '0')


    return value


def make_vehicle_index(vehicles, first_field, second_field=None, reservation=False, keynumber=False):

    #log.debug(f"make_vehicle_index: vehicles len { len(vehicles) }, 1st_field { first_field } 2nd_field { second_field }")
    result = {}
    for row in vehicles:
        vehicle = row['Vehicle']
        if first_field not in vehicle:
            continue

        if second_field is not None and second_field not in vehicle:
            continue

        key = cleanup_v_field(vehicle, first_field)
        if key is None:
            continue

        if second_field is not None:
            second = vehicle[second_field]
            if second is None:
                continue
            key = key + " " + vehicle[second_field]

        #log.debug(f"make_vehicle_index: key { key }")
        result[key] = row

    return result

def add_missing_avis_vehicles(vehicles, avis_all, avis_open):
    """ find Avis vehicles that are not in the Avis report """

    # generate some indexes to look up values faster
    i_ra  =   spreadsheet_tools.make_index(avis_all, 'Rental Agreement No')
    i_res =   spreadsheet_tools.make_index(avis_all, 'Reservation No')
    i_key =   spreadsheet_tools.make_index(avis_all, 'MVA No')
    i_plate = spreadsheet_tools.make_index(avis_all, 'License Plate State Code', 'License Plate Number')

    # note: using avis_open for this index
    i_row =   spreadsheet_tools.make_index(avis_open, spreadsheet_tools.ROW_INDEX)

    # this is duplicating work in match_avis_sheet, but it doesn't seem worth it to save the info
    v_ra = make_vehicle_index(vehicles, 'RentalAgreementNumber')
    v_res= make_vehicle_index(vehicles, 'RentalAgreementReservationNumber')
    v_key = make_vehicle_index(vehicles, 'KeyNumber')
    v_plate = make_vehicle_index(vehicles, 'PlateState', 'Plate')

    # walk through the avis_open sheet and record all the matches
    for row in avis_open:

        ra = row['Rental Agreement No']
        res = row['Reservation No']
        key = row['MVA No']
        plate = row['License Plate State Code'] + ' ' + row['License Plate Number']

        # use DisasterVehicleID as the DTT identity for a vehicle
        ra_id = get_dtt_id(v_ra, ra)
        res_id = get_dtt_id(v_res, res)
        key_id = get_dtt_id(v_key, key)
        plate_id = get_dtt_id(v_plate, plate)


    # get_dtt_id() records which vehicles were looked up.  Find all the ones that weren't found so far, and see if there is a DTT entry for them.
    for record in vehicles:
        # vehicle is not active
        if record['Status'] != 'Active':
            continue

        # already in the Avis report
        if IN_AVIS in record:
            continue

        # make sure its an Avis rental
        if record['Vehicle']['Vendor'] != 'Avis':
            continue

        # see if this vehicle is in avis_all
        vehicle = record['Vehicle']

        if vehicle['KeyNumber'] is None:
            # don't bother if there is no key number: its just a reservation
            continue

        ra = cleanup_v_field(vehicle, 'RentalAgreementNumber')
        res = cleanup_v_field(vehicle, 'RentalAgreementReservationNumber')
        key = cleanup_v_field(vehicle, 'KeyNumber')
        plate = f"{ vehicle['PlateState'] } { vehicle['Plate'] }"

        if ra in i_ra or res in i_res or key in i_key or plate in i_plate:

            log.debug(f"Adding missing vehicle to OPEN { key }")

            # one or more fields from the DTT appeared in the avis_all list, but not the avis_open list.
            # add the row to the avis_open  list
            new_avis_records = []
            if ra in i_ra:
                new_avis_records.append(i_ra[ra])
            if res in i_res:
                new_avis_records.append(i_res[res])
            if key in i_key:
                new_avis_records.append(i_key[key])
            if plate in i_plate:
                new_avis_records.append(i_plate[plate])

            # add any records to avis_open that we don't already have
            for row in new_avis_records:
                row_index = row[spreadsheet_tools.ROW_INDEX]
                if row_index not in i_row:
                    # this row isn't in avis_open yet...
                    row[MISSING_AVIS_OPEN] = True
                    avis_open.append(row)
                    i_row[row_index] = row

        else:
            # this is an entirely new vehicle that doesn't match anything in avis_all
            log.debug(f"Adding missing vehicle to ALL { key }")

            row = make_avis_from_vehicle(record)
            row[MISSING_AVIS_ALL] = True
            row[spreadsheet_tools.ROW_INDEX] = len(avis_all) + 1
            avis_open.append(row)
            avis_all.append(row)



def make_avis_from_vehicle(record):
    """ synthesize an entirely new avis record from a vehicle record """

    vehicle = record['Vehicle']

    def get_field(name):
        """ utility function to safely get field values """
        value = cleanup_v_field(vehicle, name)
        if value is None:
            return ""
        return value

    avis = {}

    # map avis fields to vehicle fields
    fields = {
            'MVA No':                   'KeyNumber',
            'License Plate State Code': 'PlateState',
            'License Plate Number':     'Plate',
            'Make':                     'Make',
            'Model':                    'Model',
            'Ext Color Code':           'Color',
            'Reservation No':           'RentalAgreementReservationNumber',
            'Rental Agreement No':      'RentalAgreementNumber',
            'Full Name':                'RentalAgreementPerson',
            }

    avis['Cost Control No'] = 'MISSING'
    for f_avis, f_vehicle in fields.items():
        avis[f_avis] = get_field(f_vehicle)

    log.debug(f"new avis record: { avis }")
    return avis



def get_dtt_id(vehicle_dict, index):
    """ utility function to look up a field in the vehicle object from the DTT """

    if index not in vehicle_dict:
        return None

    # mark the vehicle as found-in-avis-report
    vehicle_dict[index][IN_AVIS] = True

    #log.debug(f"index { index } v_dict { vehicle_dict[index] }")
    return vehicle_dict[index]['DisasterVehicleID']



def mark_cell(ws, fill, row_num, col_map, col_name):
    """ apply a fill to a particular cell """

    col_num = col_map[col_name]

    cell = ws.cell(row=row_num, column=col_num)
    cell.fill = fill


def match_avis_sheet(ws, columns, avis, vehicles):
    """ match entries from the DTT to entries in the Avis report. """

    # generate different index for the vehicles
    v_ra = make_vehicle_index(vehicles, 'RentalAgreementNumber')
    v_res= make_vehicle_index(vehicles, 'RentalAgreementReservationNumber')
    v_key = make_vehicle_index(vehicles, 'KeyNumber')
    v_plate = make_vehicle_index(vehicles, 'PlateState', 'Plate')

    #log.debug(f"plate keys: { v_plate.keys() }")

    spreadsheet_row = 1
    for row in avis:
        spreadsheet_row += 1

        ra = row['Rental Agreement No']
        res = row['Reservation No']
        key = row['MVA No']
        plate = row['License Plate State Code'] + ' ' + row['License Plate Number']

        # use DisasterVehicleID as the DTT identity for a vehicle
        ra_id = get_dtt_id(v_ra, ra)
        res_id = get_dtt_id(v_res, res)
        key_id = get_dtt_id(v_key, key)
        plate_id = get_dtt_id(v_plate, plate)

        #log.debug(f"ra { ra_id } res { res_id } key { key_id } plate { plate_id }; raw { ra } { res } { key } { plate }")

        if ra_id == None and res_id == None and key_id == None and plate_id == None:
            # vehicle doesn't appear in the DTT at all; color it blue
            fill = FILL_BLUE

            if MISSING_AVIS_ALL in row:
                fill = FILL_CYAN

            mark_cell(ws, fill, spreadsheet_row, columns, 'Rental Agreement No')
            mark_cell(ws, fill, spreadsheet_row, columns, 'Reservation No')
            mark_cell(ws, fill, spreadsheet_row, columns, 'MVA No')
            mark_cell(ws, fill, spreadsheet_row, columns, 'License Plate State Code')
            mark_cell(ws, fill, spreadsheet_row, columns, 'License Plate Number')

        elif ra_id == res_id and ra_id == key_id and ra_id == plate_id:
            # all columns match: color it green

            fill = FILL_GREEN

            if MISSING_AVIS_ALL in row:
                fill = FILL_CYAN
            mark_cell(ws, fill, spreadsheet_row, columns, 'Rental Agreement No')
            mark_cell(ws, fill, spreadsheet_row, columns, 'Reservation No')
            mark_cell(ws, fill, spreadsheet_row, columns, 'MVA No')
            mark_cell(ws, fill, spreadsheet_row, columns, 'License Plate State Code')
            mark_cell(ws, fill, spreadsheet_row, columns, 'License Plate Number')

        else:
            # else color yellow if value is found; red if value not found
            mark_cell(ws, FILL_RED if ra_id is None else FILL_YELLOW, spreadsheet_row, columns, 'Rental Agreement No')
            mark_cell(ws, FILL_RED if res_id is None else FILL_YELLOW, spreadsheet_row, columns, 'Reservation No')
            mark_cell(ws, FILL_RED if key_id is None else FILL_YELLOW, spreadsheet_row, columns, 'MVA No')
            mark_cell(ws, FILL_RED if plate_id is None else FILL_YELLOW, spreadsheet_row, columns, 'License Plate State Code')
            mark_cell(ws, FILL_RED if plate_id is None else FILL_YELLOW, spreadsheet_row, columns, 'License Plate Number')

        if MISSING_AVIS_OPEN in row:
            mark_cell(ws, FILL_YELLOW, spreadsheet_row, columns, 'Cost Control No')
        elif MISSING_AVIS_ALL in row:
            mark_cell(ws, FILL_CYAN, spreadsheet_row, columns, 'Cost Control No')




        





def insert_avis_overview(wb, config):
    """ insert an overview (documentation) sheet in the workbook """

    doc_string = f"""
This document has vehicles from the daily Avis report for this DR ({ config.DR_NUM.rjust(3, '0') }-{ config.DR_YEAR }).

On the Open RA (rental agreement) sheet there should be one row per vehicle marked as assigned to the DR
(from the 'Cost Control No' column).

Rows in the Open RA sheet have been checked against the DTT Vehicles data.  These columns are checked between
the two reports: License Plate State Code/License Plate Number, Reservation No, Rental Agreement No, Reservation No.

If all four fields match: the cells will be marked Green.

If a vehicle in the Avis report is not found in the DTT: the cells will be blue.

If all four fields don't match: any field that is not found in the DTT will be red.
Otherwise the fields will be yellow.  Red fields are usually typos in the DTT data entry.
A row of all yellow often means a data entry error where fields from different vehicles were
entered on the same DTT entry.

Cyan fields are used when a vehicle is in the DTT but not in the Avis report.

The 'Cost Control No' column is a bit different.  This column is used to encode the DR number.
The program tries to figure out if which DR matches, but it is not perfect.

If there is a vehicle in the DTT for this DR that has a different entry in the Cost Control No
field: the entry will be yellow.

This file based on the { config.AVIS_FILE } file
This file generated at { TIMESTAMP }

"""

    ws = insert_overview(wb, doc_string)

    cell = ws.cell(row=2, column=1, value="Example Green cell")
    cell.fill = FILL_GREEN
    cell = ws.cell(row=3, column=1, value="Example Yellow cell")
    cell.fill = FILL_YELLOW
    cell = ws.cell(row=4, column=1, value="Example Red cell")
    cell.fill = FILL_RED
    cell = ws.cell(row=5, column=1, value="Example Blue cell")
    cell.fill = FILL_BLUE
    cell = ws.cell(row=6, column=1, value="Example Cyan cell")
    cell.fill = FILL_CYAN

    return ws

def insert_group_overview(wb, config):

    doc_string = f"""
This document has vehicles for DR{ config.DR_NUM.rjust(3, '0') }-{ config.DR_YEAR }.

Directions for completing the Group Vehice Report

Transportation: Send report to Job Director, Section Leads/Assistant
Directors and Group Leads. Use Contact Roster in IAP for current list.

Section Leads/Assistant Directors/ Group Leads - Please have Vehicle Report reviewed for accuracy.                                     

Asset Management is the responsibility of the Group Leads.

1) Groups are separated by TABS. Please review the appropriate TAB and complete Coumns P,Q R and S.

2) If the answer is NO in Columns P,Q and R, please mark appropriate line with an X. If the work
location is incorrect,   please note current location of Driver/vehicle.

3) If the answer to Column S is YES and you can relinquish a vehicle, please mark appropriate line with an X.                       

4) Return Worksheet within 2 days of receipt to DRXXX-XXLOG-TRA1@redcross.org. If there were no changes
to your Group's vehicle list type in the body of the email - Group name (IDC, IP,..) and "No Changes".                                                         
5) Copy Section Lead/Assistant Director on email.

This file generated at { TIMESTAMP }

"""

    ws = insert_overview(wb, doc_string)
    return ws




def insert_overview(wb, doc_string):

    ws = wb.create_sheet('Overview', 0)

    ws.column_dimensions['A'].width = 120       # hard coded width....
    ws.row_dimensions[1].height = 500           # hard coded height....

    cell = ws.cell(row=1, column=1, value=doc_string)
    cell.alignment = openpyxl.styles.Alignment(wrapText=True, vertical='top')

    return ws

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
    dr_regex = re.compile(f"(dr)?\s*0*{ config.DR_NUM }(-(20)?{ config.DR_YEAR })?", flags=re.IGNORECASE)

    #avis_dr = list(filter(lambda x: dr_regex.match(x), avis_all))
    f_result = filter(lambda row: dr_regex.match(row[dr_column]), avis_all)
    avis_dr = list(f_result)
    
    return title_row, avis_columns, avis_dr, avis_all




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


def store_report(config, account, item_name, wb_bytes):
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

    result = reports.upload_file(item_name, item_name=item_name, stream=stream, stream_size=stream_size, conflict_handling='replace')
    log.debug(f"after upload: result { result }")


def make_group_report(config, vehicles):
    """ make a workbook of all vehicles, arranged by GAPs """

    wb = openpyxl.Workbook()
    insert_group_overview(wb, config)

    # sort the vehicles by GAP
    vehicles = sorted(vehicles, key=lambda x: x['Vehicle']['GAP'])

    # add the master sheet
    add_gap_sheet(wb, 'Master', vehicles)

    gap_list = get_vehicle_gap_groups(vehicles)

    for group in gap_list:
        group_vehicles = filter_by_gap_group(group, vehicles)
        add_gap_sheet(wb, group, group_vehicles)

    # serialize the workbook
    bufferview = workbook_to_buffer(wb)

    return bufferview


gap_to_group_re = re.compile('^([A-Z]+)')
def gap_to_group(gap):
    """ turn a Gap (Group/Activity/Position) name into just the Group portion """

    match = gap_to_group_re.match(gap)

    if match is None:
        return None

    return match.group(1)



def get_vehicle_gap_groups(vehicles):
    """ return a sorted list of groups from the gaps in the vehicles """

    group_re = re.compile('^([A-Z]+)')

    groups = {}

    for vehicle in vehicles:
        gap = vehicle['Vehicle']['GAP']
        group = gap_to_group(gap)

        if group is None:
            log.info("Could not find the Group in gap '{ gap }' vehicle { vehicle }")
        else:
            # remember that we saw this group; don't worry about duplicates
            groups[group] = True

    return sorted(groups.keys())



def filter_by_gap_group(group, vehicles):
    """ only return the vehicles whos GAP starts with the group """

    # we cheat a bit because we 'know' no GAP Groups are a prefix of another
    return filter(lambda x: x['Vehicle']['GAP'].startswith(group), vehicles)


def add_gap_sheet(wb, sheet_name, vehicles):
    """ make a new sheet with the specified vehicles """

    ws = wb.create_sheet(sheet_name)

    column_defs = {
            'GAP':          { 'width': 15, 'key': lambda x: x['GAP'], },
            'Driver':       { 'width': 30, 'key': lambda x: x['CurrentDriverName'], },
            'Key Number':   { 'width': 12, 'key': lambda x: x['KeyNumber'], },
            'Vendor':       { 'width': 20, 'key': lambda x: x['Vendor'], },
            'Car Info':     { 'width': 25, 'key': lambda x: f"{ x['Make'] } { x['Model'] } { x['Color'] }", },
            'Plate':        { 'width': 15, 'key': lambda x: f"{ x['PlateState'] } { x['Plate'] }", },
            'Type':         { 'width': 10, 'key': lambda x: x['VehicleType'], },
            'District':     { 'width': 10, 'key': lambda x: x['District'], },
            'Location':     { 'width': 30, 'key': lambda x: x['CurrentDriverWorkLocationName'], },
            'Lodging':      { 'width': 30, 'key': lambda x: x['CurrentDriverLodging'], },
            'Reservation':  { 'width': 15, 'key': lambda x: x['RentalAgreementReservationNumber'], },
            'Outprocessed': { 'width': 15, 'key': lambda x: "" if not x['OutProcessed'] else "OUTPROCESSED", },
        }

    row = 1
    col = 0
    for key, defs in column_defs.items():
        col = col + 1
        title = key
        cell = ws.cell(row=row, column=col, value=title)

        col_letter = openpyxl.utils.get_column_letter(col)
        if 'width' in defs:
            #log.debug(f"setting column { col_letter } to width { defs['width'] }")
            ws.column_dimensions[col_letter].width = defs['width']
        else:
            ws.column_dimensions[col_letter].auto_size = True

    row = 1
    previous_group = None
    for vehicle in vehicles:

        # only process active vehicles
        if vehicle['Status'] != 'Active':
            #log.debug(f"ignorning inactive vehicle { vehicle }")
            continue

        row += 1
        col = 0

        row_vehicle = vehicle['Vehicle']
        gap = row_vehicle['GAP']
        group = gap_to_group(gap)

        # leave a blank row when the group changes
        if group != previous_group:
            if previous_group is not None:
                row += 1
            previous_group = group

        for key, defs in column_defs.items():
            col = col + 1

            try:
                lookup_func = defs['key']
                value = lookup_func(row_vehicle)

                if value == 'None' or value == 'None None' or value == 'None None None':
                    # no valid data; make it a blank
                    value = ''
                cell = ws.cell(row=row, column=col, value=value)
            except:
                log.info(f"Could not expand column { key } row_vehicle { row_vehicle }: { sys.exc_info()[0] }")

    last_col_letter = openpyxl.utils.get_column_letter(col)
    table_ref = f"A1:{ last_col_letter }{ row }"
    table = openpyxl.worksheet.table.Table(displayName=f"{ sheet_name }Table", ref=table_ref)
    ws.add_table(table)



def parse_args():
    parser = argparse.ArgumentParser(
            description="tools to support Disaster Transportation Tools reporting",
            allow_abbrev=False)
    parser.add_argument("--debug", help="turn on debugging output", action="store_true")
    parser.add_argument("-s", "--store", help="Store file on server", action="store_true")
    parser.add_argument("--ignore-avis", help="Don't generate an Avis match report", action="store_true")
    parser.add_argument("--ignore-group", help="Don't generate a Group Vehicle report", action="store_true")

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

