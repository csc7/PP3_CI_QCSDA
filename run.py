# Your code goes here.
# You can delete these comments, but do not change the name of this file
# Write your code to expect a terminal of 80 characters wide and 24 rows high

import pandas as pd

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
    print("\nLoading worksheets...\n")
    try:
        SHEET_PARAMETERS = GSPREAD_CLIENT.open('PARAMETERS')
        TOLERANCES = SHEET_PARAMETERS.worksheet('Tolerances')

        SHEET_DAILY_REPORT = GSPREAD_CLIENT.open('daily_report')
        production_statistics = SHEET_DAILY_REPORT.worksheet('Daily_Report')

        SHEET_DISTORTION = GSPREAD_CLIENT.open('distortion')
        production_statistics = SHEET_DISTORTION.worksheet('Distortion')

        SHEET_AV_FORCE = GSPREAD_CLIENT.open('average_force')
        production_statistics = SHEET_AV_FORCE.worksheet('Average_Force')

        SHEET_POSITIONING = GSPREAD_CLIENT.open('positioning')
        positioning_data_w = SHEET_POSITIONING.worksheet('Positioning')
        #positioning_data_w.row_values(1)
        #positioning_data = pd.DataFrame(positioning_data_w.get_all_records())

        SHEET_QCSDA = GSPREAD_CLIENT.open('QCSDA')
        production_statistics = SHEET_QCSDA.worksheet('Statistics')

        #print(production_statistics)
        #print(positioning_data_w.row_values(20)[0])
        #print(type(positioning_data_w.row_values(1)[0]))
        #print(type(positioning_data_w.row_values(20)))

    except gspread.exceptions.SpreadsheetNotFound:
        print("Some files are missing or have a different name.")
        print("Please check all files are in place with the correct names.")
        return False
    return True

def validate_data():
    """
    This function checks the data in the sheets has the proper format
    """
    print("\nValidating data in the sheets...\n")



# Main part of program, calling all functions
def main(run_program):
    """
    Run all program funcions
    """
    while(run_program == "y" or run_program == "Y"):
        function_return = load_sheets()
        if (function_return == False):
            break
        validate_data()




        run_again = input('Press "y" to run the program again o any other key to close the program\n')
        if (run_again == "y" or run_again == "Y"):
            run_program = run_again
        else:
            print("Program closed.")
            break



print("-----------------------------------------------------------------")
print("Welcome to your Daily Quality Control of Seismic Dcquisition Data")
print("-----------------------------------------------------------------")
print("")
print("Make sure you have the following sheets in the working folder:")
print("")
print("PARAMETERS (you can manually enter the parameters if you prefer")
print("    or if the file is not available")
print("daily_report")
print("distortion")
print("average_force")
print("positioning")
print("QCSDA")
print("")
run_program = input('Press "y" to continue or other key to close the program\n')
if (run_program == "y" or run_program == "Y"):
    main(run_program)
else:
    print("Program closed.")

