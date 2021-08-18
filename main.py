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
from openpyxl.comments import Comment

import requests

import neil_tools
from neil_tools import spreadsheet_tools
from neil_tools import gen_templates

import config as config_static
import web_session

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

# flag field in avis object array
AVIS_SOURCE = '__AVIS_SOURCE__'
AVIS_SOURCE_OPEN       = 'OPEN'
AVIS_SOURCE_OPEN_ALL   = 'OPEN_ALL'
AVIS_SOURCE_CLOSED     = 'CLOSED'
AVIS_SOURCE_CLOSED_ALL = 'CLOSED_ALL'
AVIS_SOURCE_MISSING    = 'MISSING'

FILL_RED = openpyxl.styles.PatternFill(fgColor="FFC0C0", fill_type = "solid")
FILL_GREEN = openpyxl.styles.PatternFill(fgColor="C0FFC0", fill_type = "solid")
FILL_YELLOW = openpyxl.styles.PatternFill(fgColor="FFFFC0", fill_type = "solid")
FILL_BLUE = openpyxl.styles.PatternFill(fgColor="5BB1CD", fill_type = "solid")
FILL_CYAN = openpyxl.styles.PatternFill(fgColor="A0FFFF", fill_type = "solid")

# flag in person object that they have a vehicle
PERSON_HAS_VEHICLE = '__HAS_VEHICLE__'

COMMENT_AUTHOR = "Avis Report Reconciler Program"

NO_GAP_GROUP = 'ZZZ-No-GAP'

def main():
    args = parse_args()
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
    log.debug("running...")

    config = neil_tools.init_config(config_static, ".env")

    errors = False
    for dr in args.dr_id:
        if dr not in config.DR_CONFIGURATIONS:
            log.error(f"specified DR ID '{ dr }' not found in configured DRs.  Valid DRs are { list(config.DR_CONFIGURATIONS.keys()) }")
            errors = True

    if errors:
        return

    account_avis = None
    for dr in args.dr_id:
        dr_config = config.DR_CONFIGURATIONS[dr]

        # fetch from DTT
        session = web_session.get_session(config, dr_config)

        # ZZZ: should validate session here, and re-login if it isn't working...
        if 'DR_ID' not in config:
            get_dr_list(config, session)

        log.debug(f"DR_ID { config.DR_ID } DR_NAME { config.DR_NAME }")

        # get people and vehicles from the DTT
        vehicles = get_vehicles(config, args, session)
        people = get_people(config, args, session)
        agencies = get_agencies(config, args, session)

        # debuging only
        for row in vehicles:
            v = row['Vehicle']
            if v['Expiry'] != False:
                #log.debug(f"found expiry: { v }\n")
                pass


        account_mail = None

        if args.send or args.test_send:
            log.debug(f"initializing mail account: { dr_config.token_filename }")
            account_mail = init_o365(config, dr_config.token_filename)


        if args.do_car or args.do_no_car:
            do_status_messages(config, args, account_mail, vehicles, people)


        # avis report
        if args.do_avis:
            if account_avis == None:
                account_avis = init_o365(config, config.TOKEN_FILENAME_AVIS)

            # fetch the avis spreadsheet
            output_bytes = make_avis(config, dr_config, account_avis, vehicles, agencies)

            file_name = f"DR{ dr_config.dr_id } { FILESTAMP } Avis Report.xlsx"
            if args.store:
                store_report(config, account_avis, file_name, output_bytes)

            # save a local copy
            log.debug(f"storing avis report to { file_name }")
            with open(file_name, "wb") as fb:
                fb.write(output_bytes)

            if args.send or args.test_send:
                send_avis_report(dr_config, args, account_mail, file_name)


        # group vehicle report
        if args.do_group:
            # generate the group report
            output_bytes = make_group_report(config, dr_config, vehicles)
            file_name = f"DR{ dr_config.dr_id } { FILESTAMP } Group Vehicle Report.xlsx"

            if args.store:
                store_report(config, account_mail, file_name, output_bytes)

            # save a local copy for attachment
            log.debug(f"storing gap report to { file_name }")
            with open(file_name, "wb") as fb:
                fb.write(output_bytes)

            if args.send or args.test_send:
                send_group_report(dr_config, args, account_mail, file_name)

        if args.do_vehicles:
            output_bytes = make_vehicle_backup(config, dr_config, vehicles)

            if args.store:
                item_name = f"DR{ dr_config.dr_id } { FILESTAMP } Vehicle Backup.xlsx"
                store_report(config, account_mail, item_name, output_bytes)




def make_vehicle_backup(config, dr_config, vehicles):
    """ generate a spreadsheet of all the vehicles as a backup """

    wb = openpyxl.Workbook()
    ws = wb.create_sheet("Backup")

    class ColumnDef:

        def __init__(name, dtype="string", width=10):
            self._name = name
            self._dtype = dtype
            self._width = width

        @property 
        def name(self):
            return self._name

        @property 
        def dtype(self):
            return self._dtype

        @property 
        def width(self):
            return self._width

    columns = [
            ColumnDef("Expiring"),
            ColumnDef("Ctg"),
            ]

    for row in vehicles:
        for v in row['Vehicle']:
            pass


    wb_buffer = workbook_to_buffer(wb)
    return wb_buffer


def send_avis_report(dr_config, args, account, file_name):
    """ send an email with the avis report (contained in file_name).

        Respect the args.send and args.test_send flags (at least one of which must be set
    """

    message_body = \
f"""
<p>
Hello everyone.  This is an automated report matching vehicles in the DTT against
a list of vehicles provided by Avis.  The Avis list tends to be delayed by 24-48 hours.
</p>

<p>
If you have suggestions about the contents of the report or bugs in the program itself,
or have other tasks you think should be automated on a DR: email
<a href='mailto:{ dr_config.program_email }'>{ dr_config.program_email }</a>.
</p>
"""

    send_report_common(dr_config, args, account, file_name, "Group Vehicle Report", message_body, dr_config.reply_email)



def send_group_report(dr_config, args, account, file_name):
    """ send an email with the group vehicle report (contained in file_name).

        Respect the args.send and args.test_send flags (at least one of which must be set
    """

    message_body = \
f"""
<p>
Hello everyone.  This is an automated report showing vehicles on the DR organized by GAP 'Group'
</p>

<p>
If you have any updates or corrections: please send them to
<a href='mailto:{ dr_config.reply_email }'>{ dr_config.reply_email }</a>.
</p>

<p>
If you have suggestions about the contents of the report or bugs in the program itself,
or have other tasks you think should be automated on a DR: email
<a href='mailto:{ dr_config.program_email }'>{ dr_config.program_email }</a>.
</p>
"""

    send_report_common(dr_config, args, account, file_name, "Group Vehicle Report", message_body, dr_config.target_list)



def send_report_common(dr_config, args, account, file_name, report_type, message_body, dest_email):

    mailbox = account.mailbox()
    message = mailbox.new_message()

    if args.test_send:
        message.bcc.add(dr_config.email_bcc)

    if args.send:
        message.bcc.add(dest_email)
        log.debug(f"sending { file_name } to { dest_email }")
        posting = f"<p>This message was sent to { dest_email }.  Please do *not* reply to the whole list</p>\n"
    else:
        posting = \
f"""
<p>
DEBUG Version: not sent to the list
</p>
"""

    message.body = \
f"""
<!DOCTYPE html>
<html>
<meta http-equiv="Content-type" content="text/html" charset="UTF8" />
<title>DR{ dr_config.dr_id } { report_type }</title>
</head>
<body>

<h1>DR{ dr_config.dr_id } { report_type }</h1>
{ posting }

{ message_body }

</body>
</html>
"""

    message.subject = file_name
    message.attachments.add( file_name )

    try:
        message.send(save_to_sent_folder=True)
    except requests.RequestException as e:
        log.error(f"got an error: { e }, response json { e.response.json }")
        raise e



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

def make_avis(config, dr_config, account, vehicles, agencies):
    """ fetch the latest avis vehicle report from sharepoint """

    workbook = fetch_avis(config, account)

    output_wb = openpyxl.Workbook()
    insert_avis_overview(output_wb, config, dr_config)


    output_ws_open = output_wb.create_sheet("Open RA")

    # we now have the latest file.  Suck out all the data
    avis_open_title, avis_open_columns, avis_open, avis_open_all = read_avis_sheet(dr_config, workbook.get_worksheet('Open RA'))
    add_missing_avis_vehicles(vehicles, avis_open_all, avis_open, closed=False)

    #avis_closed_title, avis_closed_columns, avis_closed, avis_closed_all = read_avis_sheet(dr_config, workbook.get_worksheet('Closed RA'))
    #add_missing_avis_vehicles(vehicles, avis_closed_all, avis_closed, closed=True)

    # generate the 'Open RA' sheet
    output_columns = copy_avis_sheet(output_ws_open, avis_open_columns, avis_open_title, avis_open)
    match_avis_sheet(output_ws_open, output_columns, avis_open, vehicles, agencies)

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

                if isinstance(value, datetime.datetime):
                    dt = value
                else:
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

        #log.debug(f"make_vehicle_index: key { key.upper() }")
        result[key.upper()] = row

    return result

def add_missing_avis_vehicles(vehicles, avis_all, avis_open, closed):
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

    missing = []


    # walk through the avis_open sheet and record all the matches
    for row in avis_open:

        ra = row['Rental Agreement No']
        res = row['Reservation No']
        key = row['MVA No']
        plate = row['License Plate State Code'] + ' ' + row['License Plate Number']
        addr_line_3 = row['Address Line 3']


        # use DisasterVehicleID as the DTT identity for a vehicle
        ra_id = get_dtt_id(v_ra, ra)
        res_id = get_dtt_id(v_res, res)
        key_id = get_dtt_id(v_key, key)
        plate_id = get_dtt_id(v_plate, plate)
        
        row[AVIS_SOURCE] = AVIS_SOURCE_OPEN


    # get_dtt_id() records which vehicles were looked up.  Add vehicles in DTT that are not in AVIS
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
                    row[AVIS_SOURCE] = AVIS_SOURCE_OPEN_ALL
                    avis_open.append(row)
                    i_row[row_index] = row

        else:
            # ZZZ: need to check closed roster before adding missing vehicles...
            if True:
                # this is an entirely new vehicle that doesn't match anything in avis_all
                log.debug(f"Adding missing vehicle to ALL { key }")

                row = make_avis_from_vehicle(record)
                row[AVIS_SOURCE] = AVIS_SOURCE_MISSING
                row[spreadsheet_tools.ROW_INDEX] = len(avis_all) + 1
                missing.append(row)


    # add the missing rows to the end
    avis_open.extend(missing)
    avis_all.extend(missing)


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
            'Rental Loc Desc':          'PickupAgencyName',
            }

    avis['Cost Control No'] = 'MISSING'
    for f_avis, f_vehicle in fields.items():
        avis[f_avis] = get_field(f_vehicle)

    # strip off the timezone; openpyxl can't write it
    pickupDate = re.sub(r'-\d\d:\d\d$', '', get_field('RentalAgreementPickupDate'))
    avis['CO Date'] = datetime.datetime.fromisoformat(pickupDate)

    #log.debug(f"new avis record: { avis }")
    return avis



def get_dtt_id(vehicle_dict, index):
    """ utility function to look up a field in the vehicle object from the DTT """

    index = index.upper()
    if index not in vehicle_dict:
        return None

    # mark the vehicle as found-in-avis-report
    vehicle_dict[index][IN_AVIS] = True

    #log.debug(f"index { index } v_dict { vehicle_dict[index] }")
    return vehicle_dict[index]['DisasterVehicleID']



def mark_cell(ws, fill, row_num, col_map, col_name, comment=None):
    """ apply a fill to a particular cell """

    col_num = col_map[col_name]

    cell = ws.cell(row=row_num, column=col_num)

    if fill is not None:
        cell.fill = fill

    if comment is not  None:
        cell.comment = comment



def make_agency_key(agency):
    """ generate a (hopefully) unique key for the DTT agencies.  Combine Address, City, State, Zip.

        Name seems the obvious choice, but DTT names don't match AVIS names at all
    """

    address = agency['Address']
    city = agency['City']
    state = agency['State']
    zipcode = agency['Zip']

    key = f"{ address }/{ city } { state }{ zipcode }".upper()
    #log.debug(f"key { key }")
    return key

dtt_to_avis_make_dict = {
        'Honda': 'HOND',
        'Toyota': 'TOYO',
        'Kia': 'KIA ',
        'Hyundai': 'HYUN',
        'Jeep': 'JEEP',
        'Ford': 'FORD',
        'Mitsubishi': 'MITS',
        'Nissan': 'NISS',
        'Mazda': 'MAZD',
        'Subaru': 'SUBA',
        'Volkswagon': 'VOLK',
        }
dtt_to_avis_model_dict = {
        'Corolla': 'CRLA',
        'Civic': 'CIVI',
        'Frontier': 'FRO4',
        'Outlander': 'OUTL',
        'Forte': 'FORT',
        'Sorento': 'SO7F',
        'Compass': 'CMPS',
        'Altima': 'ALTI',
        'Escort': 'ECOA',
        'Prius': 'PRIH',
        'Elantra': 'ELAN',
        'CX-5': 'CX5A',
        'Camry': 'CAMR',
        'Explorer': 'EXL4',
        'Escape': 'ESCA',
        'HR-V': 'HRVA',
        'RAV 4': 'RAV4',
        'Sportage': 'SPO2',
        'Outback': 'OUTB',
        'Sentra': 'SENT',
        'Ecosport': 'ECOA',
        'Highlander': 'HIGH',
        'Golf': 'GOLF',
        'Legacy': 'LEGA',
        'Sonata': 'SONA',
        'Optima': 'OPTI',
        }
dtt_to_avis_color_dict = {
        'Silver': 'SIL',
        'White': 'WHI',
        'Blue': 'BLU',
        'Black': 'BLK',
        'Gray': 'GRY',
        'Red': 'RED',
        }

def dtt_to_avis_make(make_tuple):
    """ convert DTT make/model/color to avis """

    make = None
    model  = None
    color = None

    if make_tuple[0] in dtt_to_avis_make_dict:
        make = dtt_to_avis_make_dict[make_tuple[0]]
        
    if make_tuple[1] in dtt_to_avis_model_dict:
        model = dtt_to_avis_model_dict[make_tuple[1]]
    else:
        log.debug(f"could not find model '{ make_tuple[1] }'")

    if make_tuple[2] in dtt_to_avis_color_dict:
        color = dtt_to_avis_color_dict[make_tuple[2]]


    return (make, model, color)

def match_avis_sheet(ws, columns, avis, vehicles, agencies):
    """ match entries from the DTT to entries in the Avis report. """

    # generate different index for the vehicles
    v_ra = make_vehicle_index(vehicles, 'RentalAgreementNumber')
    v_res= make_vehicle_index(vehicles, 'RentalAgreementReservationNumber')
    v_key = make_vehicle_index(vehicles, 'KeyNumber')
    v_plate = make_vehicle_index(vehicles, 'PlateState', 'Plate')

    vid_dict = dict( (v['DisasterVehicleID'], v) for v in vehicles)

    #log.debug(f"plate keys: { v_plate.keys() }")

    multispace_re = re.compile('\s+')

    #log.debug(f"make translation dict: { dtt_to_avis_model_dict }")

    spreadsheet_row = 1
    for row in avis:
        spreadsheet_row += 1

        ra = row['Rental Agreement No']
        res = row['Reservation No']
        key = row['MVA No']
        plate = row['License Plate State Code'] + ' ' + row['License Plate Number']
        cost_control = row['Cost Control No']

        if cost_control == 'MISSING':
            # this is a synthetic row created from the DTT, not AVIS; don't bother matching
            continue

        addr_line = ''
        if 'Address Line 1' in row and 'Address Line 3' in row:
            addr_line = f"{ row['Address Line 1'] }/{ row['Address Line 3'] }"
            addr_line = multispace_re.sub(' ', addr_line)

        # use DisasterVehicleID as the DTT identity for a vehicle
        ra_id = get_dtt_id(v_ra, ra)
        res_id = get_dtt_id(v_res, res)
        key_id = get_dtt_id(v_key, key)
        plate_id = get_dtt_id(v_plate, plate)

        #log.debug(f"ra { ra_id } res { res_id } key { key_id } plate { plate_id }; raw { ra } { res } { key } { plate }")

        if AVIS_SOURCE in row:
            avis_source = row[AVIS_SOURCE]
        else:
            avis_source = None

        if ra_id == None and res_id == None and key_id == None and plate_id == None:
            # vehicle doesn't appear in the DTT at all; color it blue
            fill = FILL_BLUE

            if avis_source == AVIS_SOURCE_OPEN_ALL:
                fill = FILL_CYAN

            mark_cell(ws, fill, spreadsheet_row, columns, 'Rental Agreement No')
            mark_cell(ws, fill, spreadsheet_row, columns, 'Reservation No')
            mark_cell(ws, fill, spreadsheet_row, columns, 'MVA No')
            mark_cell(ws, fill, spreadsheet_row, columns, 'License Plate State Code')
            mark_cell(ws, fill, spreadsheet_row, columns, 'License Plate Number')

        elif ra_id == res_id and ra_id == key_id and ra_id == plate_id:
            # all columns match: color it green

            fill = FILL_GREEN

            #if avis_source == AVIS_SOURCE_OPEN_ALL:
            #    fill = FILL_CYAN

            mark_cell(ws, fill, spreadsheet_row, columns, 'Rental Agreement No')
            mark_cell(ws, fill, spreadsheet_row, columns, 'Reservation No')
            mark_cell(ws, fill, spreadsheet_row, columns, 'MVA No')
            mark_cell(ws, fill, spreadsheet_row, columns, 'License Plate State Code')
            mark_cell(ws, fill, spreadsheet_row, columns, 'License Plate Number')

            # since all four 'unique' fields match: check additional fields
            vrow = vid_dict[ra_id]
            vehicle = vrow['Vehicle']

            # check 'agency' (aka pickup location)
            agency_key = vehicle['PickupAgencyId']
            if agency_key not in agencies:
                log.debug(f"could not find agency key '{ agency_key }' in v_agencies ")
            else:
                agency = agencies[agency_key]
                agency_string = agency['AvisAgencyString']
                comment = None

                if agency_string == addr_line:
                    fill = FILL_GREEN
                else:
                    fill = FILL_YELLOW
                    log.debug(f"no agency match '{ addr_line }' / '{ agency_string }'")

                    comment = Comment(f"DTT location is { agency['Name'] } / { agency_string }", COMMENT_AUTHOR)

                mark_cell(ws, fill, spreadsheet_row, columns, 'Rental Loc Desc', comment=comment)
                mark_cell(ws, fill, spreadsheet_row, columns, 'Address Line 1')
                mark_cell(ws, fill, spreadsheet_row, columns, 'Address Line 3')


            # check pickup and expected dropoff date
            for (dtt_col, avis_col) in (('RentalAgreementPickupDate', 'CO Date'), ('DueDate', 'Exp CI Date')):
                dtt_date = datetime.datetime.fromisoformat(vehicle[dtt_col]).date()
                col_num = columns[avis_col]
                cell = ws.cell(row=spreadsheet_row, column=col_num)

                avis_pickup_dt = cell.value

                if avis_pickup_dt != None:

                    avis_pickup_date = avis_pickup_dt.date()

                    if dtt_date == avis_pickup_date:
                        fill = FILL_GREEN
                    else:
                        #log.debug(f"pickup date: key { key } dtt { dtt_date }/{vehicle[dtt_col]} avis { avis_pickup_date }/{ avis_pickup_dt }")
                        fill = FILL_YELLOW
                        cell.comment = Comment(f"DTT date is { dtt_date }", COMMENT_AUTHOR)

                    cell.fill = fill

            # check make/model/color
            avis_make_cols = ['Make', 'Model', 'Ext Color Code']
            avis_veh_make = list(row[col_name] for col_name in avis_make_cols)
            #avis_veh_make = (row['Make'], row['Model'], row['Ext Color Code'])
            dtt_veh_make_orig = (vehicle['Make'].strip(), vehicle['Model'].strip(), vehicle['Color'].strip())
            dtt_veh_make = dtt_to_avis_make(dtt_veh_make_orig)

            #log.debug(f"avis make {  avis_veh_make } dtt { dtt_veh_make } orig { dtt_veh_make_orig }")

            for (i, col_name) in enumerate(avis_make_cols):
                v = dtt_veh_make[i]

                fill = None
                comment = None

                if v is None:
                    # DTT string not found in our mapping table; ignore
                    comment = Comment(f"No mapping for DTT value { dtt_veh_make_orig[i] }", COMMENT_AUTHOR)

                else:

                    if v == avis_veh_make[i]:
                        fill = FILL_GREEN
                    else:
                        fill = FILL_YELLOW
                        comment = Comment(f"DTT value is { dtt_veh_make_orig[i] } -> { v }", COMMENT_AUTHOR)

                mark_cell(ws, fill, spreadsheet_row, columns, col_name, comment=comment)
            

        else:
            # else color yellow if value is found; red if value not found
            mark_cell(ws, FILL_RED if ra_id is None else FILL_YELLOW, spreadsheet_row, columns, 'Rental Agreement No')
            mark_cell(ws, FILL_RED if res_id is None else FILL_YELLOW, spreadsheet_row, columns, 'Reservation No')
            mark_cell(ws, FILL_RED if key_id is None else FILL_YELLOW, spreadsheet_row, columns, 'MVA No')
            mark_cell(ws, FILL_RED if plate_id is None else FILL_YELLOW, spreadsheet_row, columns, 'License Plate State Code')
            mark_cell(ws, FILL_RED if plate_id is None else FILL_YELLOW, spreadsheet_row, columns, 'License Plate Number')

        if avis_source == AVIS_SOURCE_OPEN_ALL:
            mark_cell(ws, FILL_YELLOW, spreadsheet_row, columns, 'Cost Control No')
        elif avis_source == AVIS_SOURCE_MISSING or avis_source is None:
            mark_cell(ws, FILL_CYAN, spreadsheet_row, columns, 'Cost Control No')






        





def insert_avis_overview(wb, config, dr_config):
    """ insert an overview (documentation) sheet in the workbook """

    doc_string = f"""
This document has vehicles from the daily Avis report for this DR ({ dr_config.dr_id }).

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

def insert_group_overview(wb, dr_config):

    doc_string = f"""
This document has vehicles for DR{ dr_config.dr_id }.

Asset Management is the responsibility of the Group Leads.  This report shows
who has a vehicle that Transportation has in its database.

Please help us keep the database up to date.  Send all updates to { dr_config.reply_email }.

Notes:

* Groups are separated into separate tabs.

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

def read_avis_sheet(dr_config, sheet):

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
    dr_num = dr_config.dr_num
    dr_num = dr_num.lstrip('0')
    dr_regex = re.compile(f"(dr)?\s*0*{ dr_num }(-(20)?{ dr_config.dr_year })?", flags=re.IGNORECASE)

    #avis_dr = list(filter(lambda x: dr_regex.match(x), avis_all))
    f_result = filter(lambda row: dr_regex.match(row[dr_column]), avis_all)
    avis_dr = list(f_result)

    log.debug(f"found { len(avis_dr) } vehicles associated with the DR")
    
    return title_row, avis_columns, avis_dr, avis_all




def init_o365(config, token_filename=None):
    """ do initial setup to get a handle on office 365 graph api """

    if token_filename != None:
        o365 = arc_o365.arc_o365.arc_o365(config, token_filename=token_filename)
    else:
        o365 = arc_o365.arc_o365.arc_o365(config)

    account = o365.get_account()
    if account is None:
        raise Exception("could not access office 365 graph api")

    return account

    

def get_people(config, args, session):
    """ Retrieve the people list from the DTT (as a json list) """

    data = get_json(config, args, session, 'People/Details')
    return data


def get_vehicles(config, args, session):
    """ Retrieve the vehicle list from the DTT (as a json list) """

    data = get_json(config, args, session, 'Vehicles')
    return data

def get_agencies(config, args, session):
    """ retrieve the current rental agencies for this DR

        return a dict keyed by the AgencyID
    """

    data = get_json(config, args, session, 'Agencies', prefix='api/Disasters/')

    # construct a dict keyed by AgencyID from the data array
    d = dict( (h['AgencyID'], h) for h in data )

    for h in data:
        h['AvisAgencyString'] = make_agency_key(h)



    #log.debug(f"got agencies: { d }")
    return d



def get_json(config, args, session, api_type, prefix='api/Disaster/'):

    file_name = f"cached_{ re.sub(r'/.*', '', api_type) }.json"

    data = None
    if args.cached_input:
        log.debug(f"reading cached input from { file_name }")
        with open(file_name, "rb") as f:
            buffer = f.read()

        data = json.loads(buffer)

    else:

        url = config.DTT_URL + f"{ prefix }{ config.DR_ID }/" + api_type

        r = session.get(url)
        r.raise_for_status()

        #log.debug(f"response headers { r.headers }")
        #log.debug(f"response { r.content }")

        if args.save_input:
            log.debug(f"saving to { file_name }")
            with open(file_name, "wb") as f:
                f.write(r.content)

        data = r.json()
        #log.debug(f"r.status { r.status_code } r.reason { r.reason } r.url { r.url } r.content_type { r.headers['content-type'] } data rows { len(data) }")

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


def make_group_report(config, dr_config, vehicles):
    """ make a workbook of all vehicles, arranged by GAPs """

    wb = openpyxl.Workbook()
    insert_group_overview(wb, dr_config)

    # sort the vehicles by GAP
    vehicles = sorted(vehicles, key=lambda x: vehicle_to_gap(x['Vehicle']))

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
motor_pool_re = re.compile('^Motor Pool \(([^\)]*)\)$')
def vehicle_to_group(vehicle):
    """ turn a Gap (Group/Activity/Position) name into just the Group portion """

    gap = vehicle['GAP']
    if gap is None:
        gap = ''

    driver = vehicle['CurrentDriverName']

    pool = vehicle_in_pool(vehicle)
    if pool is not None:
        return pool

    group = gap_to_group_re.match(gap)
    if group is not None:
        group_name = group.group(1)
        #log.debug(f"Found a group '{ group_name }'")
        return group_name

    return NO_GAP_GROUP


def vehicle_to_gap(vehicle):
    """ returns the motor pool name if its a pool, otherwise None """

    pool = vehicle_in_pool(vehicle)
    if pool is not None:
        return pool

    gap = vehicle['GAP']
    if gap is None:
        gap = ''

    return gap



def vehicle_in_pool(vehicle):
    """ return pool name if vehicle is in a pool, otherwise None """
    driver = vehicle['CurrentDriverName']

    if driver is None:
        return None

    pool = motor_pool_re.match(driver)

    if pool is not None:
        pool_name = pool.group(1)
        #log.debug(f"Found a pool '{ pool_name }'")
        return pool_name

    return None



def get_vehicle_gap_groups(vehicles):
    """ return a sorted list of groups from the gaps in the vehicles """

    group_re = re.compile('^([A-Z]+)')

    groups = {}

    for vehicle in vehicles:

        if vehicle['Status'] != 'Active':
            continue

        gap = vehicle['Vehicle']['GAP']
        group = vehicle_to_group(vehicle['Vehicle'])

        if group is None:
            if gap != '':
                log.info(f"Count not find Group in gap '{ gap }' vehicle { vehicle }")
            else:
                groups[NO_GAP_GROUP] = True
                pass
            pass
        else:
            # remember that we saw this group; don't worry about duplicates
            #log.debug(f"Adding group { group } from gap { gap }")
            groups[group] = True

    group_list = sorted(groups.keys())
    #log.debug(f"group_list { group_list }")
    return group_list



def filter_by_gap_group(group, vehicles):
    """ only return the vehicles whos GAP starts with the group """

    return filter(lambda x: vehicle_to_group(x['Vehicle']) == group, vehicles)


def add_gap_sheet(wb, sheet_name, vehicles):
    """ make a new sheet with the specified vehicles """

    ws = wb.create_sheet(sheet_name)

    column_defs = {
            'GAP':          { 'width': 15, 'key': lambda x: vehicle_to_gap(x), },
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
        group = vehicle_to_group(row_vehicle)

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

                if value is None or value == 'None' or value == 'None None' or value == 'None None None':
                    # no valid data; make it a blank
                    value = ''
                #log.debug(f"adding cell { row },{ col }, value { value }")
                cell = ws.cell(row=row, column=col, value=value)
            except:
                log.info(f"Could not expand column { key } row_vehicle { row_vehicle }: { sys.exc_info()[0] }")

    # ZZZ: horrible hack; for some reason the no-gap table makes excel unhappy; I haven't been
    # able to figure out why, so I'm just not adding a table for that sheet
    if sheet_name != 'ZZZ-No-GAP':

        last_col_letter = openpyxl.utils.get_column_letter(col)
        table_ref = f"A1:{ last_col_letter }{ row }"
        table = openpyxl.worksheet.table.Table(displayName=f"{ sheet_name }Table", ref=table_ref)

        ws.add_table(table)
    else:
        log.debug(f"Ignoring table for ZZZ-No-Gap sheet")

def do_status_messages(config, args, account, vehicles, people):


    #vehicle_to_driver = make_vehicle_index(vehicles, 'CurrentDriverPersonId')

    # make an index of person ids
    id_to_person = dict( (p['PersonID'], p) for p in people )

    templates = gen_templates.init()
    date = datetime.datetime.now().strftime("%Y-%m-%d %H%M")


    # group vehicles by who owns them
    person_to_vehicle = {}

    for row in vehicles:

        # ignore non-active vehicles
        if row['Status'] != 'Active':
            continue

        vehicle = row['Vehicle']

        driver_id = vehicle['CurrentDriverPersonId']
        driver = vehicle['CurrentDriverName']

        if driver_id is None:
            # no driver assigned yet
            continue

        # ZZZ: just do avis for now, because others are less well checked
        if vehicle['Vendor'] != 'Avis':
            continue


        pool = motor_pool_re.match(driver)
        if pool is not None:
            # pool vehicle; don't bother messaging
            continue


        if driver_id not in id_to_person:
            # assumption is we should always be able to find a driver
            log.error(f"Could not find driver_id { driver_id } in person list ({ vehicle['CurrentDriverNameAndMemberNo'] })")
            continue

        person = id_to_person[driver_id]

        # mark this person as having a vehicle
        person[PERSON_HAS_VEHICLE] = True

        # mark that person as having this vehicle
        if driver_id not in person_to_vehicle:
            person_to_vehicle[driver_id] = []
        person_to_vehicle[driver_id].append(row)

    templates = gen_templates.init()
    t_vehicle = templates.get_template("mail_vehicle.html")

    # now generate the emails
    if args.do_car:
        count = 0
        for person_id, l in person_to_vehicle.items():
            person = id_to_person[person_id]

            first_name = person['FirstName']
            last_name = person['LastName']
            email = person['Email']

            log.debug(f"person { first_name } { last_name } has { len(l) } vehicles")

            context = {
                    'first_name': first_name,
                    'last_name': last_name,
                    'email': email,
                    'vehicles': l,
                    'reply_email': config.REPLY_EMAIL,
                    'date': date,
                    }

            body = t_vehicle.render(context)

            log.debug(f"body { body }")

            m = account.new_message()
            if args.test_send:
                m.bcc.add(config.EMAIL_BCC)
            if args.send:
                m.to.add(email)

            m.subject = f"Vehicle Status - { date } - { first_name } { last_name }"
            m.body = body

            if args.test_send or args.send:
                m.send()

            
            # debug only
            count += 1
            if args.mail_limit and count >= 5:
                break




    # now do people without vehicles

    if args.do_no_car:
        for i, person in enumerate(people):
            
            # ignore people that have vehicles
            if PERSON_HAS_VEHICLE in person:
                continue

            status = person['Status']

            # don't hassle folks who are off the job
            if status == 'Out Process':
                continue

            first_name = person['FirstName']
            last_name = person['LastName']
            email = person['Email']

            log.debug(f"({ i }) person { first_name } { last_name } { status } has no vehicles")

            #if i > 10:
            #    break





def parse_args():
    parser = argparse.ArgumentParser(
            description="tools to support Disaster Transportation Tools reporting",
            allow_abbrev=False)
    parser.add_argument("--debug", help="turn on debugging output", action="store_true")
    parser.add_argument("-s", "--store", help="Store file on server", action="store_true")
    parser.add_argument("--do-avis", help="Generate an Avis match report", action="store_true")
    parser.add_argument("--do-group", help="Generate a Group Vehicle report", action="store_true")
    parser.add_argument("--do-vehicles", help="Generate a Vehicle backup", action="store_true")
    parser.add_argument("--do-car", help="Generate vehicle status messages", action="store_true")
    parser.add_argument("--do-no-car", help="Generate vehicle status messages", action="store_true")
    parser.add_argument("--send", help="Send messages to the actual intended recipients", action="store_true")
    parser.add_argument("--test-send", help="Add the test email account to message recipients", action="store_true")
    parser.add_argument("--mail-limit", help="debug flag to limit # of emails sent", action="store_true")
    parser.add_argument("--dr-id", help="the name of the DR (like 155-22)", required=True, action="append")

    group = parser.add_mutually_exclusive_group(required=False)
    group.add_argument("--save-input", help="Save a copy of server inputs", action="store_true")
    group.add_argument("--cached-input", help="Use cached server input", action="store_true")

    args = parser.parse_args()

    if not args.do_avis and not args.do_group and not args.do_car and not args.do_no_car:
        log.error("At least one of do-avis, do-group, do-car, or do-no-car must be specified")
        sys.exit(1)

    return args


if __name__ == "__main__":
    neil_tools.init_logging(__name__)
    log = logging.getLogger(__name__)
    main()
else:
    log = logging.getLogger(__name__)

