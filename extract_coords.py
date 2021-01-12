"""
This script receives an excel file containing addresses and other fields of interest, and returns a new table
containing these fields + the longtitude & langtitude values of each address
NOTE: You need a valid KEY from google maps API in order for this code to work
@author: Eden Elmaliah
"""
import pandas as pd
import googlemaps
from pandas import ExcelWriter
import os
import sys
from tqdm import tqdm

KEY = 'AIzaSyAPNZ3OfLPiOXpKKbcHd28vLvgSOea6Gm8'  # Enter here a valid google API key
COLUMNS = ['Tag']  # change here according to your columns in the original file (additional columns)
SHEET_NAME = "Sheet 1"
OUTPUT_NAME = "output_example.xlsx"


if __name__ == "__main__":
    script_name = os.path.relpath(__file__)
    if len(sys.argv) != 2:
        print("Usage: {} <excel file>".format(script_name))
        exit()

    filename = sys.argv[1]
    file = pd.read_excel(filename)
    streets = file['Street']
    numbers = file['House']

    gmaps = googlemaps.Client(key=KEY)
    columns = ['Address', 'x', 'y'] + COLUMNS
    df = dict()
    for col in columns:
        df[col] = []

    for i in tqdm(range(len(file))):
        street = streets[i]
        number = numbers[i]
        address = f"{street} {number}, ירושלים"
        loc = gmaps.geocode(address)
        for l, name in zip(['lng', 'lat'], ['x', 'y']):
            add = loc[0]['geometry']['location'][l]
            df[name].append(add)
        df['Address'].append(address)
        for col in COLUMNS:
            df[col].append(file[col][i])

    writer = ExcelWriter(OUTPUT_NAME)
    df = pd.DataFrame(df)
    df.to_excel(writer, SHEET_NAME)
    writer.save()