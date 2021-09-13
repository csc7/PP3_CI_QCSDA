# Import libraries
import pandas as pd
import openpyxl
import numpy as np


# gspread and google.oauth2
# copied and modified from Code Institute's
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

# Define here number of header lines for each file
header_lines_in_distorion_file = 12
header_lines_in_av_force_file = 12
header_lines_in_positioning_file = 12


# Load Google Sheets
def load_sheets_from_Google_Drive():
    """
    This function checks that all required sheets are present in Google Drive
    and read the worksheets if they are.
    """
    print("\nLoading spreadsheet and worksheets...\n")
    try:
        # Read daily report data
        SHEET_DAILY_REPORT = GSPREAD_CLIENT.open('daily_report')
        daily_report_data = SHEET_DAILY_REPORT.worksheet('Daily_Report')
        daily_report_data = pd.DataFrame(daily_report_data.get_all_values())

        # Read distortion data
        SHEET_DISTORTION = GSPREAD_CLIENT.open('distortion')
        distortion_data = SHEET_DISTORTION.worksheet('Distortion')
        distortion_data = pd.DataFrame(distortion_data.get_all_values())

        # Read average force data
        SHEET_AV_FORCE = GSPREAD_CLIENT.open('average_force')
        average_force_data = SHEET_AV_FORCE.worksheet('Average_Force')
        average_force_data = pd.DataFrame(average_force_data.get_all_values())

        # Read positioning data
        SHEET_POSITIONING = GSPREAD_CLIENT.open('positioning')
        positioning_data = SHEET_POSITIONING.worksheet('Positioning')
        positioning_data = pd.DataFrame(positioning_data.get_all_values())

        # If all right, print message to indicate this.
        print("\nSpreadsheet and worksheets loaded.\n")

    # If file/s and/or data are missing give message and close program
    except gspread.exceptions.SpreadsheetNotFound:
        print("Some QC files are missing or have a different name.")
        print("Please check all files are in place with the correct names.")
        print("\nProgram closed.\n")
        return False

    # Check if parameters are available
    try:
        SHEET_PARAMETERS = GSPREAD_CLIENT.open('PARAMETERS')
        tolerances_data = SHEET_PARAMETERS.worksheet('Tolerances')
        tolerances_data = pd.DataFrame(tolerances_data.get_all_values())

    # If not available give the user the chance to input them
    except gspread.exceptions.SpreadsheetNotFound:
        print("No parameters file.")
        param = input('Press "P" to enter them manually or other key to ' +
                      'close the program.\n')
        # Initialize and assing zero to tolerances, so data structure is
        # defined when returning the function value
        if (param == "P" or param == "p"):
            tolerances_data = {
                "tolerances": {
                    "fleets": float(0),
                    "vibs_per_fleet": float(0),
                    "max_cog_dist": float(0),
                    "max_distortion": float(0),
                    "min_av_force": float(0),
                    "max_av_force": float(0),
                }
            }
            # Return data
            return [tolerances_data, daily_report_data, distortion_data,
                    average_force_data, positioning_data]
        else:
            return False

    # Return data
    return [tolerances_data, daily_report_data, distortion_data,
            average_force_data, positioning_data]


# Load data locally, from Microsoft Excel files
def load_sheets_locally():
    """
    This function checks that all required sheets are present in the local
    drive and read the worksheets if they are.
    """
    print("\nLoading spreadsheet and worksheets...\n")
    # Read all data except tolerances
    try:
        daily_report_data = pd.read_excel('qcdata/daily_report.xlsx',
                                          engine='openpyxl')
        distortion_data = pd.read_excel('qcdata/distortion.xlsx',
                                        engine='openpyxl')
        average_force_data = pd.read_excel('qcdata/average_force.xlsx',
                                           engine='openpyxl')
        positioning_data = pd.read_excel('qcdata/positioning.xlsx',
                                         engine='openpyxl')
        SHEET_QCSDA = pd.ExcelFile('qcdata/QCSDA.xlsx',
                                   engine='openpyxl')

        print("\nSpreadsheet and worksheets loaded.\n")

    # If file/s and/or data are missing give message and close program
    except FileNotFoundError as e:
        print(f"Something went wrong, file not found: {e}. " +
              "Please try again.\n")
        print("Program closed.")
        return False

    except UnboundLocalError as e:
        print(f"Something went wrong: {e}. Please try again.\n")
        print("Program closed.")
        return False

    # Check if parameters are available
    try:
        tolerances_data = pd.read_excel('qcdata/PARAMETERS.xlsx',
                                        engine='openpyxl')

    except FileNotFoundError as e:
        print("No parameters file.")
        param = input('Press "P" to enter them manually or other key to ' +
                      'close the program.\n')
        # Initialize and assing zero to tolerances, so data structure is
        # defined when returning the function value
        if (param == "P" or param == "p"):
            tolerances_data = {
                "tolerances": {
                    "fleets": float(0),
                    "vibs_per_fleet": float(0),
                    "max_cog_dist": float(0),
                    "max_distortion": float(0),
                    "min_av_force": float(0),
                    "max_av_force": float(0),
                }
            }
            # Return data
            return [tolerances_data, daily_report_data, distortion_data,
                    average_force_data, positioning_data, SHEET_QCSDA]
        else:
            return False

    # If not available give the user the chance to input them
    except UnboundLocalError as e:
        print(f"Something went wrong, {e}. Please try again.\n")
        return False

    # Return data
    return [tolerances_data, daily_report_data, distortion_data,
            average_force_data, positioning_data, SHEET_QCSDA]


# Load parameters, checking first if they are available. If they are not
# ask the user to input them
def load_parameters(qc_dict):
    """
    This functions checks if a parameters is available (by trying to read them
    from the correspoinding value in the dictionary that was previously
    loaded). If a value does not exist, the user is required to input it.
    """
    qc_dictionary = qc_dict
    # Check for number of fleets.
    # Keep asking for it if it does not exist.
    while True:
        try:
            qc_dictionary['tolerances']['fleets'] = \
                float(input("\nEnter number of fleets: \n"))
            break
        except ValueError:
            print("Please enter a number.")
    # Check for number of vibrators per fleet.
    # Keep asking for it if it does not exist.
    while True:
        try:
            qc_dictionary['tolerances']['vibs_per_fleet'] = \
                float(input("Enter number of vibrators per fleets: \n"))
            break
        except ValueError:
            print("Please enter a number.")
    # Check for maximum center of gravity (COG) distance.
    # Keep asking for it if it does not exist.
    while True:
        try:
            qc_dictionary['tolerances']['max_cog_dist'] = \
                float(input("Enter maximum distance to COG: \n"))
            break
        except ValueError:
            print("Please enter a number.")
    # Check for maximum distortion.
    # Keep asking for it if it does not exist.
    while True:
        try:
            qc_dictionary['tolerances']['max_distortion'] = \
                float(input("Enter maximum distortion: \n"))
            break
        except ValueError:
            print("Please enter a number.")
    # Check for minimum average force.
    # Keep asking for it if it does not exist.
    while True:
        try:
            qc_dictionary['tolerances']['min_av_force'] = \
                float(input("Enter minimum average force: \n"))
            break
        except ValueError:
            print("Please enter a number.")
    # Check for maximum average force.
    # Keep asking for it if it does not exist.
    while True:
        try:
            qc_dictionary['tolerances']['max_av_force'] = \
                float(input("Enter maximum average force: \n"))
            break
        except ValueError:
            print("Please enter a number.")
    # Return data
    return qc_dictionary


# Validate data in Google Sheets
def validate_data_from_Google(data_to_validate):
    """
    This function checks the data in the Google Sheets has the proper format
    and load the values that will be used for quality control in a
    dictionary
    """
    print("\nValidating data in the sheets...\n")

    # Check if first element of dictionary is another dictionary.
    # The input of this funciont is the output of
    # load_sheets_from_Google_Drive() or load_sheets_locally(),
    # which returns an initialized dictionary if acquisition parameters need
    # to be asked for, or returns a different data structure (gspread table
    # for Google Sheets or Pandas dataframe for local Microsoft Excel).
    # Therefore, if the first element is a dictionary, it means the
    # acquisition parameters (tolerances) were not entered and it is needed to
    # ask for them by calling load_parameters(). If not a dictionary, just
    # read the data in the "else" statement.
    if (isinstance(data_to_validate[0], dict)):
        qc_dictionary = load_parameters(data_to_validate[0])
        # Check if other files and data (other than acquisition parameters)
        # are available and print an error message if they are not
        try:
            qc_dictionary.update({
                "daily_report": {
                    "date": data_to_validate[1].iloc[6, 1],
                    "daily_prod": float(data_to_validate[1].iloc[21, 2]),
                    "daily_layout": float(data_to_validate[1].iloc[22, 2]),
                    "daily_pick_up": float(data_to_validate[1].iloc[23, 2]),
                },
                "distortion": data_to_validate[2].iloc[
                    (header_lines_in_distorion_file-2):],
                "average_force": data_to_validate[3].iloc[
                    (header_lines_in_av_force_file-2):],
                "positioning": data_to_validate[4].iloc[
                    (header_lines_in_positioning_file-2):, 0:8],
            })
            print("\nCurrent Acquisition Parameters:")
            # Call function to print current acquisition parameters
            print_acq_param(qc_dictionary)
        except TypeError as e:
            print(f"Data could not be validted: {e}. Please check format is " +
                  "correct for each file.\n")
            return False
    else:
        try:
            qc_dictionary = {
                "tolerances": {
                    "fleets": float(data_to_validate[0].iloc[0, 2]),
                    "vibs_per_fleet": float(data_to_validate[0].iloc[1, 2]),
                    "max_cog_dist": float(data_to_validate[0].iloc[6, 2]),
                    "max_distortion": float(data_to_validate[0].iloc[7, 2]),
                    "min_av_force": float(data_to_validate[0].iloc[8, 3]),
                    "max_av_force": float(data_to_validate[0].iloc[8, 4]),
                },
                "daily_report": {
                    "date": data_to_validate[1].iloc[6, 1],
                    "daily_prod": float(data_to_validate[1].iloc[21, 2]),
                    "daily_layout": float(data_to_validate[1].iloc[22, 2]),
                    "daily_pick_up": float(data_to_validate[1].iloc[23, 2]),
                },
                "distortion": data_to_validate[2].iloc[
                    (header_lines_in_distorion_file-2):],
                "average_force": data_to_validate[3].iloc[
                    (header_lines_in_av_force_file-2):],
                "positioning": data_to_validate[4].iloc[
                    (header_lines_in_positioning_file-2):, 0:8],
            }
            qc_dictionary = ask_to_overwrite_parameters(qc_dictionary)
        except ValueError as e:
            # If an acquisition parameter value is missing, give the user
            # the option to enter all parameters
            print("At least one parameters is missing.")
            param = input('Press "P" to enter them manually or other key ' +
                          'to close the program.\n')
            # Initialize and assing zero to tolerances, so data structure is
            # defined when calling load_parameters() function
            if (param == "P" or param == "p"):
                tolerances_data = {
                    "tolerances": {
                        "fleets": float(0),
                        "vibs_per_fleet": float(0),
                        "max_cog_dist": float(0),
                        "max_distortion": float(0),
                        "min_av_force": float(0),
                        "max_av_force": float(0),
                    }
                }
                qc_dictionary = load_parameters(tolerances_data)
                qc_dictionary.update({
                    "daily_report": {
                        "date": data_to_validate[1].iloc[6, 1],
                        "daily_prod": float(data_to_validate[1].iloc[21, 2]),
                        "daily_layout": float(data_to_validate[1].iloc[22, 2]),
                        "daily_pick_up":
                            float(data_to_validate[1].iloc[23, 2]),
                    },
                    "distortion": data_to_validate[2].iloc[
                        (header_lines_in_distorion_file-2):],
                    "average_force": data_to_validate[3].iloc[
                        (header_lines_in_av_force_file-2):],
                    "positioning": data_to_validate[4].iloc[
                        (header_lines_in_positioning_file-2):, 0:8],
                })
            else:
                # Return false to close the program
                return False

        # Round data to two digits avoid differences when reading from
        # Google Drive or locally, as the sheets from different sources seem
        # to have different float values.
        for item in qc_dictionary["distortion"]:
            item = round(float(item), 2)
        for item in qc_dictionary["average_force"]:
            item = round(float(item), 2)
        for item in qc_dictionary["positioning"]:
            item = round(float(item), 2)

        # Read coordinates of real and planned points to later compute
        # distance to COG
        x_p = qc_dictionary['positioning'].iloc[:, 1]
        x_planned = pd.to_numeric(x_p)
        y_p = qc_dictionary['positioning'].iloc[:, 2]
        y_planned = pd.to_numeric(y_p)
        z_p = qc_dictionary['positioning'].iloc[:, 3]
        z_planned = pd.to_numeric(z_p)
        x_c = qc_dictionary['positioning'].iloc[:, 4]
        x_cog = pd.to_numeric(x_c)
        y_c = qc_dictionary['positioning'].iloc[:, 5]
        y_cog = pd.to_numeric(y_c)
        z_c = qc_dictionary['positioning'].iloc[:, 6]
        z_cog = pd.to_numeric(z_c)

        # Compute distance from real position of point to the
        # center of gravity (COG)
        distance = ((x_planned - x_cog)*(x_planned - x_cog) +
                    (y_planned - y_cog)*(y_planned - y_cog) +
                    (z_planned - z_cog)*(z_planned - z_cog))**0.5

        # Create new column with distance to COG
        qc_dictionary['positioning'].iloc[:, 7] = distance

    # Return updated dictionary
    return (qc_dictionary)


def validate_data_locally(data_to_validate):
    """
    This function checks the data in the local drive sheets has the proper
    format and load the values that will be used quality control in a
    dictionary
    """
    print("\nValidating data in the sheets...\n")

    # As before, check if first element of dictionary is another dictionary.
    # The input of this funciont is the output of
    # load_sheets_from_Google_Drive() or load_sheets_locally(),
    # which returns an initialized dictionary if acquisition parameters need
    # to be asked for, or returns a different data structure (gspread table
    # for Google Sheets or Pandas dataframe for local Microsoft Excel).
    # Therefore, if the first element is a dictionary, it means the
    # acquisition parameters (tolerances) were not entered and it is needed to
    # ask for them by calling load_parameters(). If not a dictionary, just
    # read the data in the "else" statement.
    if (isinstance(data_to_validate[0], dict)):
        qc_dictionary = load_parameters(data_to_validate[0])
        # Check if other files and data (other than acquisition parameters)
        # are available and print an error message if they are not
        try:
            qc_dictionary.update({
                "daily_report": {
                    "date": data_to_validate[1].iloc[7, 1],
                    "daily_prod": float(data_to_validate[1].iloc[22, 2]),
                    "daily_layout": float(data_to_validate[1].iloc[23, 2]),
                    "daily_pick_up": float(data_to_validate[1].iloc[24, 2]),
                },
                "distortion": round(data_to_validate[2].iloc[
                    (header_lines_in_distorion_file-1):], 2),
                "average_force": round(data_to_validate[3].iloc[
                    (header_lines_in_av_force_file-1):], 2),
                "positioning": round(data_to_validate[4].iloc[
                    (header_lines_in_positioning_file-1):, 0:8], 2),
            })
            print("\nCurrent Acquisition Parameters:")
            # Call function to print current acquisition parameters
            print_acq_param(qc_dictionary)

        except TypeError as e:
            print(f"Data could not be validted: {e}. " +
                  "Please check format is correct for each file.\n")
            return False

    else:
        try:
            qc_dictionary = {
                "tolerances": {
                    "fleets": float(data_to_validate[0].iloc[1, 2]),
                    "vibs_per_fleet": float(data_to_validate[0].iloc[2, 2]),
                    "max_cog_dist": float(data_to_validate[0].iloc[7, 2]),
                    "max_distortion": float(data_to_validate[0].iloc[8, 2]),
                    "min_av_force": float(data_to_validate[0].iloc[9, 3]),
                    "max_av_force": float(data_to_validate[0].iloc[9, 4]),
                },
                "daily_report": {
                    "date": data_to_validate[1].iloc[7, 1],
                    "daily_prod": float(data_to_validate[1].iloc[22, 2]),
                    "daily_layout": float(data_to_validate[1].iloc[23, 2]),
                    "daily_pick_up": float(data_to_validate[1].iloc[24, 2]),
                },
                "distortion": round(data_to_validate[2].iloc[
                    (header_lines_in_distorion_file-1):], 2),
                "average_force": round(data_to_validate[3].iloc[
                    (header_lines_in_av_force_file-1):], 2),
                "positioning": round(data_to_validate[4].iloc[
                    (header_lines_in_positioning_file-1):, 0:8], 2),
            }

            # Check if a parameter is missing in the Microsoft Excel file
            # and give the user the option to enter acquisition parameters
            if (np.isnan(qc_dictionary['tolerances']['fleets']) or
                    np.isnan(qc_dictionary['tolerances']['vibs_per_fleet']) or
                    np.isnan(qc_dictionary['tolerances']['max_cog_dist']) or
                    np.isnan(qc_dictionary['tolerances']['max_distortion']) or
                    np.isnan(qc_dictionary['tolerances']['min_av_force']) or
                    np.isnan(qc_dictionary['tolerances']['max_av_force'])):
                print("At least one parameters is missing.")
                param = input('Press "P" to enter them manually or other ' +
                              'key to close the program.\n')
                # Initialize and assing zero to tolerances, so data structure
                # is defined when calling load_parameters() function
                if (param == "P" or param == "p"):
                    tolerances_data = {
                        "tolerances": {
                            "fleets": float(0),
                            "vibs_per_fleet": float(0),
                            "max_cog_dist": float(0),
                            "max_distortion": float(0),
                            "min_av_force": float(0),
                            "max_av_force": float(0),
                        }
                    }
                    qc_dictionary = load_parameters(tolerances_data)
                    qc_dictionary.update({
                        "daily_report": {
                            "date": data_to_validate[1].iloc[7, 1],
                            "daily_prod":
                                float(data_to_validate[1].iloc[22, 2]),
                            "daily_layout":
                                float(data_to_validate[1].iloc[23, 2]),
                            "daily_pick_up":
                                float(data_to_validate[1].iloc[24, 2]),
                        },
                        "distortion": round(data_to_validate[2].iloc[
                            (header_lines_in_distorion_file-1):], 2),
                        "average_force": round(data_to_validate[3].iloc[
                            (header_lines_in_av_force_file-1):], 2),
                        "positioning": round(data_to_validate[4].iloc[
                            (header_lines_in_positioning_file-1):, 0:8], 2),
                    })
                else:
                    # Return false to close the program
                    return False
            else:
                qc_dictionary = ask_to_overwrite_parameters(qc_dictionary)
        except TypeError as e:
            # Inform the user if there is other error
            print(f"Data could not be validted: {e}. Please check format is " +
                  "correct for each file.\n")
            return False

    # Read coordinates of real and planned points to later compute
    # distance to COG
    x_planned = qc_dictionary['positioning'].iloc[:, 1]
    y_planned = qc_dictionary['positioning'].iloc[:, 2]
    z_planned = qc_dictionary['positioning'].iloc[:, 3]
    x_cog = qc_dictionary['positioning'].iloc[:, 4]
    y_cog = qc_dictionary['positioning'].iloc[:, 5]
    z_cog = qc_dictionary['positioning'].iloc[:, 6]

    # Create new column with distance to COG
    distance = ((x_planned - x_cog)*(x_planned - x_cog) +
                (y_planned - y_cog)*(y_planned - y_cog) +
                (z_planned - z_cog)*(z_planned - z_cog))**0.5

    # Create new column with distance to COG
    qc_dictionary['positioning'].iloc[:, 7] = distance

    # Return updated dictionary
    return (qc_dictionary)


# Ask to overwrite acquisition parameters (tolerances)
def ask_to_overwrite_parameters(qc_dictionary):
    """
    This functions prints the current acquisition parameters (tolerances) and
    ask the user to overwrite them by calling load_parameters() functions
    if that is the case
    """
    print("\nCurrent Acquisition Parameters:")
    print_acq_param(qc_dictionary)
    print("Would you like to assing new parameters?\n")
    overwrite = input('Press "Y" to overwrite or other key to continue:\n')
    if (overwrite == 'Y' or overwrite == 'y'):
        load_parameters(qc_dictionary)
        print("\nNew Acquisition Parameters:")
        print_acq_param(qc_dictionary)
    return qc_dictionary


# Print acquisition parameters
def print_acq_param(qc_dictionary):
    """
    This functions prints the current acquisition parameters (tolerances)
    """
    print(f"{int(qc_dictionary['tolerances']['fleets'])} fleets and " +
          f"{int(qc_dictionary['tolerances']['vibs_per_fleet'])} vibrators " +
          f"per fleet")
    print(f"Maximum distance from COG: " +
          f"{qc_dictionary['tolerances']['max_cog_dist']}")
    print(f"Maximum distortion: " +
          f"{qc_dictionary['tolerances']['max_distortion']}")
    print(f"{qc_dictionary['tolerances']['min_av_force']}% as minimum and " +
          f"{qc_dictionary['tolerances']['max_av_force']}% as maximum " +
          "average force\n")
    return


# Get daily production/amounts and check if they match the quantities in the
# QC files
def get_daily_amounts(qc_dictionary):
    """
    This function compares the data available in the distortion, average force,
    and positioning files and compare them with the daily acquisition of
    VPs, generating a warning message if they do not match.
    """
    index = qc_dictionary['distortion'].index
    distortion_measurements = len(index)

    index = qc_dictionary['average_force'].index
    av_force_measurements = len(index)

    index = qc_dictionary['positioning'].index
    positioning_measurements = len(index)

    qc_dictionary.update({
        "Distortion_Measurements": distortion_measurements,
        "Av_Force_Measurements": av_force_measurements,
        "Positioning_Measurements": positioning_measurements,
    })

    # Divide the amount of data by the number of vibrators per fleet, since
    # there is on entry per vibrator, not per point
    chk1 = qc_dictionary["Distortion_Measurements"] / \
        qc_dictionary["tolerances"]["vibs_per_fleet"]
    chk2 = qc_dictionary["Av_Force_Measurements"] / \
        qc_dictionary["tolerances"]["vibs_per_fleet"]
    chk3 = positioning_measurements

    # Return a warning message, if any.
    warning_message = []

    # Compare amount of points/data with daily report and create warning
    # message if there is any mismatch
    if chk1 != qc_dictionary["daily_report"]['daily_prod']:
        warning_message.append(f"- distortion file ({chk1} VPs)")
    if chk2 != qc_dictionary["daily_report"]['daily_prod']:
        warning_message.append(f"- average force file ({chk2} VPs)")
    if chk3 != qc_dictionary["daily_report"]['daily_prod']:
        warning_message.append(f"- positioning file ({chk3} VPs)")

    # Return data and warning messages (if any)
    return (qc_dictionary, warning_message)


# Identify points out of specifications
def get_points_to_reaquire(qc_dictionary):
    """
    This function computes the points that are out of specifications and
    need to be reaquired.
    """
    # Read acquisition parameters/tolerances
    max_distortion = float(qc_dictionary['tolerances']['max_distortion'])
    min_av_force = float(qc_dictionary['tolerances']['min_av_force'])
    max_av_force = float(qc_dictionary['tolerances']['max_av_force'])
    max_cog_dist = float(qc_dictionary['tolerances']['max_cog_dist'])

    # Convert acquisition parameters/tolerances to float type unsing NumPy
    qc_dictionary['distortion'] = qc_dictionary['distortion']\
        .astype(np.float32)
    qc_dictionary['average_force'] = qc_dictionary['average_force']\
        .astype(np.float32)
    qc_dictionary['positioning'] = qc_dictionary['positioning']\
        .astype(np.float32)

    # Select points out of specifications by distortion issues
    out_of_spec_distortion = qc_dictionary['distortion']
    [qc_dictionary['distortion'].iloc[:, 4] > max_distortion]

    # Select points out of specifications by average force issues. Since the
    # tolerance has maximun and minimum values, compute them separately and
    # concatenate all of them
    out_of_spec_force_max = qc_dictionary['average_force']
    [qc_dictionary['average_force'].iloc[:, 4] > max_av_force]
    out_of_spec_force_min = qc_dictionary['average_force']
    [qc_dictionary['average_force'].iloc[:, 4] < min_av_force]
    temp_out_force = [out_of_spec_force_max, out_of_spec_force_min]
    out_of_spec_force = pd.concat(temp_out_force)
    out_of_spec_force = out_of_spec_force[out_of_spec_force.iloc[:, 4] <
                                          min_av_force]

    # Select points out of specifications by positioning issues
    out_of_spec_cog = qc_dictionary['positioning']
    [qc_dictionary['positioning'].iloc[:, 7] > 1]

    # Total VPs out of specifications
    index = out_of_spec_distortion.index
    out_distor_total = len(index)
    index = out_of_spec_force.index
    out_force_total = len(index)
    index = out_of_spec_cog.index
    out_cog_total = len(index)

    # Create dictionary with points out of specifications
    out_of_spec_dictionary = {
        "Out_of_Spec_Distortion": out_of_spec_distortion,
        "Out_of_Spec_Force": out_of_spec_force,
        "Out_of_Spec_COG": out_of_spec_cog,
        "Total_Out_Distortion": out_distor_total,
        "Total_Out_Force": out_force_total,
        "Total_Out_COG": out_cog_total,
    }

    # Print amount of points out of specifications by category
    print("\nTotal points to reaquire by distortion issues: " +
          f"{out_distor_total}")
    print("Total points to reaquire by average force issues: " +
          f"{out_force_total}")
    print("Total points to reaquire by positioning issues: " +
          f"{out_cog_total}\n")

    return out_of_spec_dictionary


# Visualization options for the data
def visualize_data(*data_to_visualize):
    """
    This function provides many options to visualize the data.
    """
    while(True):
        # Menu
        print('\nSelect one option to show below and press the number + ' +
              '"enter":')
        print("1 - Daily Statistics")
        print("2 - Acquisition Parameters")
        print("3 - Amount of points to be reacquired")
        print("4 - Points to be reacquired")
        print("Other key - Return to main menu")
        answer = input("\nSelect option: \n")

        if (answer == '1'):
            # Daily Statistics
            print("-----------------------")
            print("Daily production: " +
                  f"{data_to_visualize[0]['daily_report']['daily_prod']}")
            print("Daily layout: " +
                  f"{data_to_visualize[0]['daily_report']['daily_layout']}")
            print(f"Daily pick-up: " +
                  f"{data_to_visualize[0]['daily_report']['daily_pick_up']}")
            print("-----------------------")
        elif (answer == '2'):
            # Acquisition Parameters
            print("-------------------------------")
            print("Vibrator fleets: " +
                  f"{data_to_visualize[0]['tolerances']['fleets']}")
            print("Vibrators per fleet: " +
                  f"{data_to_visualize[0]['tolerances']['vibs_per_fleet']}")
            print("Maximum COG-planned distance: " +
                  f"{data_to_visualize[0]['tolerances']['max_cog_dist']}")
            print("Maximum distortion: " +
                  f"{data_to_visualize[0]['tolerances']['max_distortion']}")
            print("Minimum average force: " +
                  f"{data_to_visualize[0]['tolerances']['min_av_force']}")
            print("Maximum average force: " +
                  f"{data_to_visualize[0]['tolerances']['max_av_force']}")
            print("-------------------------------")
        elif (answer == '3'):
            # Amount of points to be reacquired
            # Check first if they were computed. If they were not, ask the
            # user to compute them.
            try:
                disto = data_to_visualize[1]['Total_Out_Distortion']
                av_for = data_to_visualize[1]['Total_Out_Force']
                pos = data_to_visualize[1]['Total_Out_COG']
                print("-----------------------------" +
                      "-----------------------------")
                print(f"Total points to reacquire:")
                print(f"By distorition issues: {disto}")
                print(f"By distorition average force issues: {av_for}")
                print(f"By positioning issues: {pos}")
                print("-----------------------------" +
                      "-----------------------------\n")
            except IndexError:
                print("\nPlease compute points to reacquire first.\n")
        elif (answer == '4'):
            # Points to be reacquired
            # Check first if they were computed. If they were not, ask the
            # user to compute them.
            # Show useful and summarised information: line, station, and
            # X- and Y-coordinates.
            try:
                disto_p = data_to_visualize[1]['Out_of_Spec_Distortion']\
                    .iloc[:, [0, 1, 5, 6]]
                av_for_p = data_to_visualize[1]['Out_of_Spec_Force']\
                    .iloc[:, [0, 1, 5, 6]]
                # Do not show header and select columns
                pos_p = data_to_visualize[1]['Out_of_Spec_COG']\
                    .iloc[:, 0:4]
                print("-----------------------------" +
                      "-----------------------------")
                # Points to reacquire by distortion issues
                print("Points to reacquire by distortion issues:\n")
                print("LINE - STATION - X - Y")
                print(disto_p)
                print("-----------------------------" +
                      "-----------------------------\n\n")
                print("-----------------------------" +
                      "-----------------------------")
                # Points to reacquire by average force issues
                print("Points to reacquire by average force issues:\n")
                print("LINE - STATION - X - Y")
                print(av_for_p)
                print("-----------------------------" +
                      "-----------------------------\n\n")
                print("-----------------------------" +
                      "-----------------------------")
                # Points to reacquire by positioning issues
                print("Points to reacquire by positioning issues:\n")
                print("LINE - STATION - X - Y")
                print(pos_p)
                print("-----------------------------" +
                      "-----------------------------\n\n")
            except IndexError:
                # Print message to ask the user to compute points to reacquire
                # first (if they were not computed)
                print("\nPlease compute points to reacquire first.\n")
        else:
            break


# Update QCSDA Google Sheet or Microsoft Excel file
def update_qcsda(qc_dictionary, date, source):
    """
    This function update the QCSDA Google Sheet/Microsoft Excel file
    by adding extra worksheets/sheets with the points that need to be
    reacquired the next day.
    """
    # Ask the user to select where to write, remembering where the data were
    # read from last
    print("\nSelect where to save the list of points to be reacquired:")
    print('"G" + "enter": Google Drive')
    print('"L" + "enter": Local Drive\n')
    print(f"Last data were collected from {source}\n")
    answer = input("Select option: \n")

    # To write Google Sheets in Google Drive
    if (answer == "G" or answer == "g"):
        print("\nUpdating files in Google Drive...\n")

        # Create specific name for sheet, with date of acquisition of points.
        # Check if the sheet was not written before,
        # ask to create if it was not.
        # Points out of specifications by distortion issues
        sheet_name_temp = 'Redo_distortion_' + str(date)[0:10]
        SHEET_QCSDA = GSPREAD_CLIENT.open('QCSDA')
        try:
            SHEET_QCSDA.worksheet(sheet_name_temp)\
                .update(qc_dictionary['Out_of_Spec_Distortion']
                        .values.tolist())
        except gspread.exceptions.WorksheetNotFound:
            print(f"\nSheet {sheet_name_temp} does not exist; " +
                  "it will be created.\n")
            SHEET_QCSDA.add_worksheet(title=sheet_name_temp,
                                      rows="0", cols="0")
            SHEET_QCSDA.worksheet(sheet_name_temp)\
                .update(qc_dictionary['Out_of_Spec_Distortion']
                        .values.tolist())

        # Create specific name for sheet, with date of acquisition of points.
        # Check if the sheet was not written before,
        # ask to create if it was not.
        # Points out of specifications by average force issues
        sheet_name_temp = 'Redo_force_' + str(date)[0:10]
        SHEET_QCSDA = GSPREAD_CLIENT.open('QCSDA')
        try:
            SHEET_QCSDA.worksheet(sheet_name_temp)\
                .update(qc_dictionary['Out_of_Spec_Force'].values.tolist())
        except gspread.exceptions.WorksheetNotFound:
            print(f"\nSheet {sheet_name_temp} does not exist; " +
                  "it will be created.\n")
            SHEET_QCSDA.add_worksheet(title=sheet_name_temp,
                                      rows="0", cols="0")
            SHEET_QCSDA.worksheet(sheet_name_temp)\
                .update(qc_dictionary['Out_of_Spec_Force'].values.tolist())

        # Create specific name for sheet, with date of acquisition of points.
        # Check if the sheet was not written before,
        # ask to create if it was not.
        # Points out of specifications by positioning issues
        sheet_name_temp = 'Redo_COG_' + str(date)[0:10]
        SHEET_QCSDA = GSPREAD_CLIENT.open('QCSDA')
        try:
            SHEET_QCSDA.worksheet(sheet_name_temp)\
                .update(qc_dictionary['Out_of_Spec_COG'].values.tolist())
        except gspread.exceptions.WorksheetNotFound:
            print(f"\nSheet {sheet_name_temp} does not exist; " +
                  "it will be created.\n")
            SHEET_QCSDA.add_worksheet(title=sheet_name_temp,
                                      rows="0", cols="0")
            SHEET_QCSDA.worksheet(sheet_name_temp)\
                .update(qc_dictionary['Out_of_Spec_COG'].values.tolist())

        # Inform the user that sheets were created and file updated
        print("\nFile Updated.\n")

    if (answer == "L" or answer == "l"):

        print("\nUpdating files in local drive...\n")

        # Create specific name for sheet, with date of acquisition of points.
        # Check if the sheet was not written before,
        # ask to create if it was not.
        # Points out of specifications by distortion issues
        sheet_name_temp = 'Redo_distortion_' + str(date)[0:10]
        with pd.ExcelWriter('qcdata/QCSDA.xlsx', mode='a', engine='openpyxl',
                            if_sheet_exists='replace') as writer:
            qc_dictionary['Out_of_Spec_Distortion']\
                .to_excel(writer, sheet_name=sheet_name_temp, startrow=0,
                          startcol=0, header=False, index=False)

        # Create specific name for sheet, with date of acquisition of points.
        # Check if the sheet was not written before,
        # ask to create if it was not.
        # Points out of specifications by average force issues
        sheet_name_temp = 'Redo_force_' + str(date)[0:10]
        with pd.ExcelWriter('qcdata/QCSDA.xlsx', mode='a', engine='openpyxl',
                            if_sheet_exists='replace') as writer:
            qc_dictionary['Out_of_Spec_Force']\
                .to_excel(writer, sheet_name=sheet_name_temp, startrow=0,
                          startcol=0, header=False, index=False)

        # Create specific name for sheet, with date of acquisition of points.
        # Check if the sheet was not written before,
        # ask to create if it was not.
        # Points out of specifications by positioning issues
        sheet_name_temp = 'Redo_COG_' + str(date)[0:10]
        with pd.ExcelWriter('qcdata/QCSDA.xlsx', mode='a', engine='openpyxl',
                            if_sheet_exists='replace') as writer:
            qc_dictionary['Out_of_Spec_COG']\
                .to_excel(writer, sheet_name=sheet_name_temp, startrow=0,
                          startcol=0, header=False, index=False)

        # Inform the user that sheets were created and file updated
        print("\nFiles Updated.\n")
