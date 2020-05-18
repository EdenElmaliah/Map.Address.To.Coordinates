"""
This script receives an excel file containing addresses and other fields of interest, and returns a new table
containing these fields + the longtitude & langtitude values of each address
NOTE: You need a valid KEY from google maps API in order for this code to work
"""
import pandas as pd
import googlemaps
from xlwt import Workbook
import os
import sys

KEY = 'AIzaSyAPNZ3OfLPiOXpKKbcHd28vLvgSOea6Gm8'  # for example, this key is no longer valid
COLUMNS = ['Address', 'Type']  # change here according to your columns in the original file
SHEET_NAME = "Sheet 1"
OUTPUT_NAME = "/Users/edene_236/Desktop/GoogleMaps Locations.xls"


if __name__ == "__main__":
    script_name = os.path.relpath(__file__)
    if len(sys.argv) != 2:
        print("Usage: {} <excel file>".format(script_name))
        exit()

    filename = sys.argv[1]
    file = pd.read_excel(filename, header=None)
    file.columns = COLUMNS

    additional_fields = COLUMNS[1:]
    gmaps = googlemaps.Client(key=KEY)
    addresses = dict()
    for i in range(len(file)):
        address = file.at[i, 'Address']
        addresses[address] = dict()
        loc = gmaps.geocode(address)
        addresses[address]['x'] = loc[0]['geometry']['location']['lng']
        addresses[address]['y'] = loc[0]['geometry']['location']['lat']
        for field in additional_fields:
            addresses[address][field] = file.at[i, field]

    wb = Workbook(encoding='utf-8')
    sheet1 = wb.add_sheet(SHEET_NAME, cell_overwrite_ok=True)

    new_cols = COLUMNS.copy()
    new_cols.insert(1, 'y')
    new_cols.insert(1, 'x')

    for i, header in enumerate(new_cols):
        sheet1.write(0, i, header)

    for i, add in enumerate(addresses):
        sheet1.write(i + 1, 0, add)
        sheet1.write(i + 1, 1, addresses[add]['x'])
        sheet1.write(i + 1, 2, addresses[add]['y'])
        for j, field in enumerate(additional_fields):
            sheet1.write(i + 1, j + 3, addresses[add][field])

    wb.save(OUTPUT_NAME)

