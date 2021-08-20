# Your code goes here.
# You can delete these comments, but do not change the name of this file
# Write your code to expect a terminal of 80 characters wide and 24 rows high

# Copied and modified from Code Institute's "Love Sandwiches - Essentials Project"
# August 18th, 2021

import gspread
from google.oauth2.service_account import Credentials



SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
    ]

CREDS = Credentials.from_service_account_file('creds.json')
SCOPED_CREDS = CREDS.with_scopes(SCOPE)
GSPREAD_CLIENT = gspread.authorize(SCOPED_CREDS)
SHEET = GSPREAD_CLIENT.open('distortion')

sales = SHEET.worksheet('Sheet1')

data = sales.get_all_values()

#print(data)


import pandas as pd
from gspread_pandas import Spread, Client


df = pd.read_csv(sales)
