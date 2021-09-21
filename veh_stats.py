
# compute vehicle ratio statistics.

import logging
import re

import openpyxl.utils
import openpyxl.styles

import main

log = logging.getLogger(__name__)


def compute_stats(vehicles, roster, gap_prefix):
    """ compute vehicle statistics by group """

    if gap_prefix != '' and gap_prefix != 'ALL':
        gap_regex = re.compile(f"^{ gap_prefix }/")
    else:
        gap_regex = re.compile(f"^")

    #log.debug(f"vehicle len { len(list(vehicles)) }, roster len { len(roster) } gap_prefix '{ gap_prefix }'")

    # compute total matching MDA people on the roster
    roster_count = 0
    for r in roster.values():
        gap = r['GAP(s)']
        tnm = r['T&M']

        if tnm != "MDA":
            continue

        if gap_regex.match(gap):
            roster_count += 1
            #log.debug(f"roster: adding { r['Name'] } gap { gap } prefix '{ gap_prefix }'")

    vehicle_count = 0
    for r in vehicles:

        if r['Status'] != 'Active':
            continue

        veh = r['Vehicle']

        if veh['VehicleCategoryCode'] != 'R':
            continue

        gap = main.vehicle_to_gap(r['Vehicle'])

        if gap_regex.match(gap):
            vehicle_count += 1

    log.debug(f"gap_prefix { gap_prefix } roster_count { roster_count } vehicle_count { vehicle_count }")
    return roster_count, vehicle_count


