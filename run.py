# Import all required functions from functions.py file
from functions import *


# Main part of program, calling all functions
def main(run_program):
    """
    Main function with main menu, which calls all other funcions.
    The program guides the user to read, compute and write
    the QC data, alerting about missing or non-matching data
    """

    # Run/Start the program while "G", "g", "L" or "l" character is selected
    # Select other character to close the program
    # Select "G" or "g" to read from Google Drive,
    # or select "L" or "l" to read from local disk/computer
    while(run_program == "G" or run_program == "g" or
          run_program == "L" or run_program == "l"):

        # Loading and validating data from Google Drive
        if (run_program == "G" or run_program == "g"):
            data_source = 'Google Drive'
            data_loaded = load_sheets_from_Google_Drive()
            if (not data_loaded):
                break
            qc_data = validate_data_from_Google(data_loaded)
            if (not qc_data):
                print("\nProgram closed.\n")
                break

        # Loading and validating data locally
        if (run_program == "L" or run_program == "l"):
            data_source = 'Local Drive'
            data_loaded = load_sheets_locally()
            if (not data_loaded):
                break
            qc_data = validate_data_locally(data_loaded)
            if (not qc_data):
                print("\nProgram closed.\n")
                break

        # Get daily acquisition numbers/totals
        daily_amounts, warning_message = get_daily_amounts(qc_data)
        # Alert if daily production does not match the amount of points
        # in the QC files
        if (warning_message != []):
            print("WARNING:")
            print("Acording to daily report, " +
                  f"{str(daily_amounts['daily_report']['date'])[0:10]} "
                  "daily production is " +
                  f"{daily_amounts['daily_report']['daily_prod']} VPs.")
            print("The amount of data in the following file/s does/do not " +
                  f"match the daily production:\n{warning_message}\n")
            print("If you continue you will get statistics " +
                  "for an incomplete data set.")
            print('Press "Y" + "enter" to continue or other key + "enter" ' +
                  'to close the program.\n')
            cont_option = input('Select option: \n')
            if (cont_option == "Y" or cont_option == "y"):
                pass
            else:
                print("Program closed.")
                break
        # If quantities match, print daily data
        print(f"\n{str(daily_amounts['daily_report']['date'])[0:10]} daily " +
              f"production: {daily_amounts['daily_report']['daily_prod']}, " +
              "with ")
        print(f"{daily_amounts['daily_report']['daily_layout']} " +
              "planted geophones and " +
              f"{daily_amounts['daily_report']['daily_pick_up']} picked up.\n")

        # Data are OK, present menu options to user
        while(True):
            # Submenu
            print('Select one option below and press the number + "enter":')
            print("1 - Compute Points to reacquire")
            print("2 - Visualize data")
            print("3 - Update QCSDA Spreadsheet with points to reacquire")
            print("Other key - Back to main menu\n")
            answer = input("Select option: \n")

            if (answer == '1'):
                points_to_reaquire = get_points_to_reaquire(daily_amounts)
            elif (answer == '2'):
                try:
                    visualize_data(daily_amounts, points_to_reaquire)
                except UnboundLocalError:
                    # Alert the user if data are missing
                    print("\nWARNING!\nPoints to be reaquired not computed,\
                        passing existing and daily data only!\n")
                    visualize_data(qc_data)
            elif (answer == '3'):
                try:
                    update_qcsda(points_to_reaquire,
                                 daily_amounts['daily_report']['date'],
                                 data_source)
                except UnboundLocalError:
                    # Instruct the user to compute points to reacquire
                    # before trying to write them in the QCSDA file
                    print("\nPlease get points to reacquire first. ")
                    print ("Run option 1 before updating.\n")
                    continue
            else:
                break

        # End of program, ask to restart or close
        print('Press "G" + "enter" to read data from Google Drive')
        print('Press "L" + "enter" to read data locally')
        print('Press any other key + "enter" to close the program.\n')
        run_again = input('Select option: \n')
        if (run_again == "G" or run_again == "g" or
                run_again == "L" or run_again == "l"):
            run_program = run_again
        else:
            print("Program closed.")
            break


# Welcome screen
print("")
print("-----------------------------------------------------------------")
print("Welcome to your Daily Quality Control of Seismic Data Acquisition")
print("-----------------------------------------------------------------")
print("")
print("Make sure you have the following sheets in the working folder:")
print("")
print(" * PARAMETERS (you can manually enter the parameters if you prefer")
print("    or if the file is not available)")
print(" * daily_report")
print(" * distortion")
print(" * average_force")
print(" * positioning")
print(" * QCSDA")
print("")
print('Press "G" + "enter" to read data from Google Drive')
print('Press "L" + "enter" to read data locally')
print('Press any other key + "enter" to close the program.\n')

# Ask the user to select an option
run_program = input('Select option: \n')

# Call main program function for the first time
# Select "G" or "g" to read from Google Drive,
# or select "L" or "l" to read from local disk/computer
if (run_program == "G" or run_program == "g" or
        run_program == "L" or run_program == "l"):
    main(run_program)
else:
    print("Program closed.")
