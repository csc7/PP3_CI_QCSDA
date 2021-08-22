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



def load_sheets_from_Google_Drive():
    """
    This function checks that all required sheets are present in Google Drive
    and read the worksheets if they are.
    """
    print("\nLoading worksheets...\n")
    try:
        SHEET_PARAMETERS = GSPREAD_CLIENT.open('PARAMETERS')
        tolerances_data = SHEET_PARAMETERS.worksheet('Tolerances')

        SHEET_DAILY_REPORT = GSPREAD_CLIENT.open('daily_report')
        daily_report_data = SHEET_DAILY_REPORT.worksheet('Daily_Report')

        SHEET_DISTORTION = GSPREAD_CLIENT.open('distortion')
        distortion_data = SHEET_DISTORTION.worksheet('Distortion')

        SHEET_AV_FORCE = GSPREAD_CLIENT.open('average_force')
        average_force_data = SHEET_AV_FORCE.worksheet('Average_Force')

        SHEET_POSITIONING = GSPREAD_CLIENT.open('positioning')
        positioning_data = SHEET_POSITIONING.worksheet('Positioning')
        #positioning_data_w.row_values(1)
        #positioning_data = pd.DataFrame(positioning_data_w.get_all_records())

        # This sheet contains five worksheets!
        SHEET_QCSDA = GSPREAD_CLIENT.open('QCSDA')
        #production_statistics = SHEET_QCSDA.worksheet('Statistics')

        #print(production_statistics)
        #print(positioning_data_w.row_values(20)[0])
        #print(type(positioning_data_w.row_values(1)[0]))
        #print(type(positioning_data_w.row_values(20)))

    except gspread.exceptions.SpreadsheetNotFound:
        print("Some files are missing or have a different name.")
        print("Please check all files are in place with the correct names.")
        return False

    return [tolerances_data, distortion_data, distortion_data, average_force_data,
            positioning_data, SHEET_QCSDA]


def validate_data(data_to_validate):
    """
    This function checks the data in the sheets has the proper format
    and load the values that will be used quality control in a
    dictionary
    """
    print("\nValidating data in the sheets...\n")

    print(type(data_to_validate))
    print(data_to_validate[0].cell(3, 3).value)

    try:

        # Define here number of header lines for each file
        header_lines_in_distorion_file = 12
        header_lines_in_av_force_file = 12
        header_lines_in_positioning_file = 12


        distortion = []
        #distortion = data_to_validate[2].get_all_values()[header_lines_in_distorion_file:],

        for i in range (13, 480):
            for item in data_to_validate[2].row_values(i):
                float (item)
   
        #distortion_df = pd.DataFrame(list(distortion))
        #distortion_df = distortion_df.transpose()
        #distortion_df2 = pd.DataFrame(distortion_df)





        #for item in data_to_validate[2].get_all_values()[header_lines_in_distorion_file + 1:]:
        #    distortion.append(float(item))

        qc_dictionary = {
            "tolerances": {
                "fleets": float(data_to_validate[0].cell(3, 3).value),
                "vibs_per_fleet": float(data_to_validate[0].cell(4, 3).value),
                "max_cog_dist": float(data_to_validate[0].cell(9, 3).value),
                "max_distortion": float(data_to_validate[0].cell(10, 3).value),
                "min_av_force": float(data_to_validate[0].cell(11, 4).value),
                "max_av_force": float(data_to_validate[0].cell(11, 5).value),
            },
            "daily_report": {
                "date": data_to_validate[1].cell(9, 2).value,
                "daily_prod": float(data_to_validate[1].cell(24, 3).value),
                "daily_layout": float(data_to_validate[1].cell(25, 3).value),
                "daily_pickup": float(data_to_validate[1].cell(26, 3).value),
            },
            #"distortion": distortion,
            "distortion": distortion_df,
            "average_force": data_to_validate[3].get_all_values()[header_lines_in_av_force_file + 1:],
            "positioning": data_to_validate[4].get_all_values()[header_lines_in_positioning_file + 1:],
        }
        for item in qc_dictionary["tolerances"].values():
            if item == 70:
                print(distortion)
                print(item)

    except TypeError as e:
        print(f"Data could not be validted: {e}. Please check format is correct for each file.\n")
        return False

    return qc_dictionary




# Main part of program, calling all functions
def main(run_program):
    """
    Run all program funcions
    """
    while(run_program == "G" or run_program == "g" or
          run_program == "L" or run_program == "l"):
        # 1st Function: Loading
        data_loaded = load_sheets_from_Google_Drive()
        if (data_loaded == False):
            break
        print (data_loaded)
        # 2nd Function: Validation
        qc_data = validate_data(data_loaded[slice(0, 5)])
        #print(qc_data)




        print('Press "G" + "enter" to read data from Google Drive')
        print('Press "L" + "enter" to read data locally')
        print('Press any other key + "enter" to close the program')
        run_again = input('Select option: ')
        if (run_program == "G" or run_program == "g" or
            run_program == "L" or run_program == "l"):
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
print('Press "G" + "enter" to read data from Google Drive')
print('Press "L" + "enter" to read data locally')
print('Press any other key + "enter" to close the program')

run_program = input('Select option: ')

if (run_program == "G" or run_program == "g" or
    run_program == "L" or run_program == "l"):
    main(run_program)
else:
    print("Program closed.")

