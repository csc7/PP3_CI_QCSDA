# Your code goes here.
# You can delete these comments, but do not change the name of this file
# Write your code to expect a terminal of 80 characters wide and 24 rows high

import pandas as pd
import openpyxl


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

# Define here number of header lines for each file
header_lines_in_distorion_file = 12
header_lines_in_av_force_file = 12
header_lines_in_positioning_file = 12

def load_sheets_from_Google_Drive():
    """
    This function checks that all required sheets are present in Google Drive
    and read the worksheets if they are.
    """
    print("\nLoading spreadsheet and worksheets...\n")
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
        print("\nSpreadsheet and worksheets loaded.\n")

    except gspread.exceptions.SpreadsheetNotFound:
        print("Some files are missing or have a different name.")
        print("Please check all files are in place with the correct names.")
        return False

    return [tolerances_data, daily_report_data, distortion_data, average_force_data,
            positioning_data, SHEET_QCSDA]



def load_sheets_locally():
    
    print("\nLoading spreadsheet and worksheets...\n")
    try:
        tolerances_data = pd.read_excel('qcdata/PARAMETERS.xlsx', engine='openpyxl')
        daily_report_data = pd.read_excel('qcdata/daily_report.xlsx', engine='openpyxl')
        distortion_data = pd.read_excel('qcdata/distortion.xlsx', engine='openpyxl')
        average_force_data = pd.read_excel('qcdata/average_force.xlsx', engine='openpyxl')
        positioning_data = pd.read_excel('qcdata/positioning.xlsx', engine='openpyxl')
        SHEET_QCSDA = pd.ExcelFile('qcdata/QCSDA.xlsx', engine='openpyxl')

        print("\nSpreadsheet and worksheets loaded.\n")
    
    except FileNotFoundError as e:
        print(f"Something went wrong, {e}. Please try again.\n")
        return False
    
    except UnboundLocalError as e:
        print(f"Something went wrong, {e}. Please try again.\n")
        return False
    
    return [tolerances_data, daily_report_data, distortion_data, average_force_data,
            positioning_data, SHEET_QCSDA]


def validate_data(data_to_validate):
    """
    This function checks the data in the sheets has the proper format
    and load the values that will be used quality control in a
    dictionary
    """
    print("\nValidating data in the sheets...\n")

    print(type(data_to_validate))
    
    try:

        #distortion = []
        #distortion = data_to_validate[2].get_all_values()[header_lines_in_distorion_file:],

        #for i in range (13, 480):
        #    for item in data_to_validate[2].row_values(i):
        #        float (item)
   
        #distortion_df = pd.DataFrame(list(distortion))
        #distortion_df = distortion_df.transpose()
        #distortion_df2 = pd.DataFrame(distortion_df)

        #for item in data_to_validate[2].get_all_values()[header_lines_in_distorion_file + 1:]:
        #    distortion.append(float(item))

        qc_dictionary = {
            "tolerances": {
                "fleets": data_to_validate[0].iloc[1, 2],
                "vibs_per_fleet": data_to_validate[0].iloc[2, 2],
                "max_cog_dist": data_to_validate[0].iloc[7, 2],
                "max_distortion": data_to_validate[0].iloc[8, 2],
                "min_av_force": data_to_validate[0].iloc[9, 3],
                "max_av_force": data_to_validate[0].iloc[9, 4],
            },
            "daily_report": {
                "date": data_to_validate[1].iloc[7, 1],
                "daily_prod": data_to_validate[1].iloc[22, 2],
                "daily_layout": data_to_validate[1].iloc[23, 2],
                "daily_pick_up": data_to_validate[1].iloc[24, 2],
            },
            #"distortion": distortion,
            "distortion": data_to_validate[2].iloc[(header_lines_in_distorion_file-1):],
            "average_force": data_to_validate[3].iloc[(header_lines_in_av_force_file-1):],
            "positioning": data_to_validate[4].iloc[(header_lines_in_positioning_file-1):, 1:9],
        }

        # Calculate distance from COG to planned coordinate before retuning the dictionary
        x_planned = qc_dictionary['positioning'].iloc[:, 1]
        y_planned = qc_dictionary['positioning'].iloc[:, 2]
        z_planned = qc_dictionary['positioning'].iloc[:, 3]
        x_cog = qc_dictionary['positioning'].iloc[:, 4]
        y_cog = qc_dictionary['positioning'].iloc[:, 5]
        z_cog = qc_dictionary['positioning'].iloc[:, 6]

        distance = ((x_planned - x_cog)*(x_planned - x_cog) + (y_planned - y_cog)*(y_planned - y_cog) + (z_planned - z_cog)*(z_planned - z_cog))**0.5

        qc_dictionary['positioning'].iloc[:, 7] = distance

        #temp_pos = [qc_dictionary['positioning'], distance]
        #posi = pd.concat(temp_pos)

        #index = x_planned.index
        #ind = len(index)
        #for i in range (0, ind+1):
        #    distance[i] = math.sqrt((x_planned[i] - x_cog[i])*(x_planned[i] - x_cog[i]) + (y_planned[i] - y_cog[i])*(y_planned[i] - y_cog[i]) + (z_planned[i] - z_cog[i])*(z_planned[i] - z_cog[i]))

        QCSDA_SPREADSHEET = data_to_validate[5]

    except TypeError as e:
        print(f"Data could not be validted: {e}. Please check format is correct for each file.\n")
        return False

    return (qc_dictionary, QCSDA_SPREADSHEET)

def get_daily_amounts(qc_dictionary):

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

    chk1 = qc_dictionary["Distortion_Measurements"] / qc_dictionary["tolerances"]["vibs_per_fleet"]
    chk2 = qc_dictionary["Av_Force_Measurements"] / qc_dictionary["tolerances"]["vibs_per_fleet"]
    chk3 = positioning_measurements

    warning_message = []

    if chk1 != qc_dictionary["daily_report"]['daily_prod']:
        warning_message.append(f"- distortion file ({chk1} VPs)")
    if chk2 != qc_dictionary["daily_report"]['daily_prod']:
        warning_message.append(f"- average force file ({chk2} VPs)")
    if chk3 != qc_dictionary["daily_report"]['daily_prod']:
        warning_message.append(f"- positioning file ({chk3} VPs)")

    return (qc_dictionary, warning_message)



def get_points_to_reaquire(qc_dictionary):

    max_distortion = qc_dictionary['tolerances']['max_distortion']
    min_av_force = qc_dictionary['tolerances']['min_av_force']
    max_av_force = qc_dictionary['tolerances']['max_av_force']
    max_cog_dist = qc_dictionary['tolerances']['max_cog_dist']

    out_of_spec_distortion = qc_dictionary['distortion'][ qc_dictionary['distortion'].iloc[:, 4] > max_distortion] 
    
    out_of_spec_force_max = qc_dictionary['average_force'][ qc_dictionary['average_force'].iloc[:, 4] > max_av_force]
    out_of_spec_force_min = qc_dictionary['average_force'][ qc_dictionary['average_force'].iloc[:, 4] < min_av_force]
    temp_out_force = [out_of_spec_force_max, out_of_spec_force_min]
    out_of_spec_force = pd.concat(temp_out_force)    
    out_of_spec_force = out_of_spec_force[ out_of_spec_force.iloc[:, 4] < min_av_force]

    out_of_spec_cog = qc_dictionary['positioning'][ qc_dictionary['positioning'].iloc[:, 7] > 1 ]
    
    # Total VPs out of specifications
    index = out_of_spec_distortion.index
    out_distor_total = len(index)
    index = out_of_spec_force.index
    out_force_total = len(index)
    index = out_of_spec_cog.index
    out_cog_total = len(index)   

    out_of_spec_dictionary = {
        "Out_of_Spec_Distortion" : out_of_spec_distortion,
        "Out_of_Spec_Force" : out_of_spec_force,
        "Out_of_Spec_COG" : out_of_spec_cog,
        "Total_Out_Distortion": out_distor_total,
        "Total_Out_Force": out_force_total,
        "Total_Out_COG": out_cog_total,
    }

    print(out_distor_total) 
    print(out_force_total) 
    print(out_of_spec_cog) 

    return out_of_spec_dictionary



def visualize_data(*data_to_visualize):
    #print(qc_dictionary['positioning'])

    while(True):
        print('\nSelect one option to show below and press the number + "enter":')
        print("1 - Daily Production")
        print("2 - Acquisition Parameters")
        print("3 - Amount of points to be reacquired")
        print("4 - Points to be reacquired")
        print("Other key - Return to main menu")
        answer = input("\nSelect option: \n")
    
        if (answer == '1'):
            print(f"Daily production: {data_to_visualize[0]['daily_report']['daily_prod']}")
        elif (answer =='2'):
            print(answer)
        elif (answer =='3'):
            try:
                print(f"Total points to reacquire by distorition issues: {data_to_visualize[1]['Total_Out_Distortion']}")

            except IndexError:
                print("\nPlease get points to reacquire first.\n")
        elif (answer =='4'):
            print(answer)
        else:
            break
    #print((data_to_visualize[0]))
    

    #xls = pd.ExcelFile('qcdata/QCSDA.xlsx')
    #df2 = pd.read_excel(xls, 'Redo_distortion')
    #qc_dictionary['positioning'].to_excel("qcdata/QCSDA.xlsx", sheet_name='Redo_distortion') 
    #print(df2)


def update_qcsda(qc_dictionary, date):

    print("\nUpdating files...\n")

    sheet_name = 'Redo_distortion_' + str(date)[0:10]
    with pd.ExcelWriter('qcdata/QCSDA.xlsx', mode ='a', engine='openpyxl', if_sheet_exists = 'replace') as writer:
        qc_dictionary['Out_of_Spec_Distortion'].to_excel(writer, sheet_name=sheet_name, startrow = 0, startcol = 0, header = False, index = False)
    
    sheet_name = 'Redo_force_' + str(date)[0:10]
    with pd.ExcelWriter('qcdata/QCSDA.xlsx', mode ='a', engine='openpyxl', if_sheet_exists = 'replace') as writer:
        qc_dictionary['Out_of_Spec_Force'].to_excel(writer, sheet_name=sheet_name, startrow = 0, startcol = 0, header = False, index = False)
    
    sheet_name = 'Redo_COG_' + str(date)[0:10]
    with pd.ExcelWriter('qcdata/QCSDA.xlsx', mode ='a', engine='openpyxl', if_sheet_exists = 'replace') as writer:
        qc_dictionary['Out_of_Spec_COG'].to_excel(writer, sheet_name=sheet_name, startrow = 0, startcol = 0, header = False, index = False)
    
    print("\nFiles Update...\n")
    #writer.save()
    #writer.close()



# Main part of program, calling all functions
def main(run_program):
    """
    Run all program funcions
    """
    while(run_program == "G" or run_program == "g" or
          run_program == "L" or run_program == "l"):

        # 1st Function: Loading
        if (run_program == "G" or run_program == "g"):
            data_loaded = load_sheets_from_Google_Drive()
            if (data_loaded == False):
                break
        if (run_program == "L" or run_program == "l"):
            data_loaded = load_sheets_locally()
            if (data_loaded == False):
                break
        
        # 2nd Function: Validation
        #qc_data = validate_data(data_loaded[slice(0, 5)])
        qc_data, QCSDA_EXCEL_FILE = validate_data(data_loaded)
        print(type(qc_data))


        # 3rd Function: Get daily acquisition numbers/totals
        daily_amounts, warning_message = get_daily_amounts (qc_data)
        if (warning_message != []):
            print("WARNING:")
            print(f"Acording to daily report, daily production is {daily_amounts['daily_report']['daily_prod']} VPs.")
            print(f"The amount of data in the following file/s does/do not match the daily production:\n{warning_message}\n")
            print("If you continue you will get statistics for an incomplete data set.")
            print('Press "Y" + "enter" to continue or other key + "enter" to close the program.\n')
            cont_option = input('Select option: \n')
            if (cont_option == "Y" or cont_option == "y"):
                continue
            else:
                print("Program closed.")
                break               
        print(f"\n{daily_amounts['daily_report']['date']} daily production: {daily_amounts['daily_report']['daily_prod']},")
        print(f"with {daily_amounts['daily_report']['daily_layout']} planted geophones and {daily_amounts['daily_report']['daily_pick_up']} picked up.\n")
        
        # Data OK, present menu options to user
        print('Select one option below and press the number + "enter":')
        print("1 - Get Points to reacquire")
        print("2 - Visualize data")
        print("3 - Update QCSDA Spreadsheet with points to reacquire")
        print("4 - Restart")
        print("Other key - Close\n")
        answer = input("Select option: \n")
        print(daily_amounts)
        if (answer == '1'):
            points_to_reaquire = get_points_to_reaquire(daily_amounts)
        elif (answer == '2'):
            try:
                visualize_data(daily_amounts, points_to_reaquire)
            except UnboundLocalError:
                print("\nWARNING!\nPoints to be reaquired not computed, passing existing and daily data only!\n")
                visualize_data(qc_data)
        elif (answer == '3'):
            try:
                update_qcsda(points_to_reaquire, daily_amounts['daily_report']['date'])
            except UnboundLocalError:
                print("\nPlease get points to reacquire first (run option 1) before updating\n")
        elif (answer == '4'):
            pass
        else:
            break

        






        # End of program, ask to restart or close

        print('Press "G" + "enter" to read data from Google Drive')
        print('Press "L" + "enter" to read data locally')
        print('Press any other key + "enter" to close the program.\n')
        run_again = input('Select option: \n')
        if (run_program == "G" or run_program == "g" or
            run_program == "L" or run_program == "l"):
            run_program = run_again
        else:
            print("Program closed.")
            break


print("")
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

run_program = input('Select option: \n')

if (run_program == "G" or run_program == "g" or
    run_program == "L" or run_program == "l"):
    main(run_program)
else:
    print("Program closed.")

