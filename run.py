# Your code goes here.
# You can delete these comments, but do not change the name of this file
# Write your code to expect a terminal of 80 characters wide and 24 rows high

# Copied and modified from Code Institute's
# "Love Sandwiches - Essentials Project" on
# August 18th, 2021.

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

def load_sheets():
    """
    This function checks that all required sheets are present in Google Drive
    and read the worksheets if they are.
    """
    print("Loading worksheets...")
    try:
        SHEET_PARAMETERS = GSPREAD_CLIENT.open('PARAMETERS')
        TOLERANCES = SHEET_PARAMETERS.worksheet('Tolerances')
        SHEET_DAILY_REPORT = GSPREAD_CLIENT.open('daily_report')
        production_statistics = SHEET_DAILY_REPORT.worksheet('Daily_Report')
        SHEET_DISTORTION = GSPREAD_CLIENT.open('distortion')
        production_statistics = SHEET_DISTORTION.worksheet('Distortion')
        SHEET_AV_FORCE = GSPREAD_CLIENT.open('average_force')
        production_statistics = SHEET_AV_FORCE.worksheet('Average_Force')
        SHEET_POSTIONING = GSPREAD_CLIENT.open('positioning')
        production_statistics = SHEET_POSTIONING.worksheet('Positioning')
        SHEET_QCSDA = GSPREAD_CLIENT.open('QCSDA')
        production_statistics = SHEET_QCSDA.worksheet('Statistics')
    except ValueError as e:
        print(f"Check that all required files are saved.\n")
        return False
    return True

#data = sales.get_all_values()
#print(data)


# Main part of program, calling all functions
def main():
    """
    Run all program funcions
    """
    load_sheets()
    #validate_data
    #

print ("Welcome to your Daily Quality Control of Seismic Dcquisition Data")
main()