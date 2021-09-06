from functions import *

# Main part of program, calling all functions
def main(run_program):
    """
    Main function with main menu, which calls all other funcions.
    """
    while(run_program == "G" or run_program == "g" or
          run_program == "L" or run_program == "l"):

        # 1st Function: Loading
        if (run_program == "G" or run_program == "g"):
            data_source = 'Google Drive'
            data_loaded = load_sheets_from_Google_Drive()
            if (data_loaded == False):
                break
            qc_data = validate_data_from_Google(data_loaded)
            

        if (run_program == "L" or run_program == "l"):
            data_source = 'Local Drive'
            data_loaded = load_sheets_locally()
            if (data_loaded == False):
                break
            qc_data = validate_data_locally(data_loaded)
            

        # 2nd Function: Validation
        #qc_data = validate_data(data_loaded[slice(0, 5)])

        #qc_data, QCSDA_EXCEL_FILE = validate_data(data_loaded)

        # 3rd Function: Get daily acquisition numbers/totals
        daily_amounts, warning_message = get_daily_amounts (qc_data)
        if (warning_message != []):
            print("WARNING:")
            print(f"Acording to daily report, {str(daily_amounts['daily_report']['date'])[0:10]} daily production is {daily_amounts['daily_report']['daily_prod']} VPs.")
            print(f"The amount of data in the following file/s does/do not match the daily production:\n{warning_message}\n")
            print("If you continue you will get statistics for an incomplete data set.")
            print('Press "Y" + "enter" to continue or other key + "enter" to close the program.\n')
            cont_option = input('Select option: \n')
            if (cont_option == "Y" or cont_option == "y"):
                pass
            else:
                print("Program closed.")
                break               
        print(f"\n{str(daily_amounts['daily_report']['date'])[0:10]} daily production: {daily_amounts['daily_report']['daily_prod']},")
        print(f"with {daily_amounts['daily_report']['daily_layout']} planted geophones and {daily_amounts['daily_report']['daily_pick_up']} picked up.\n")
        
        # Data OK, present menu options to user
        while(True):
        
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
                    print("\nWARNING!\nPoints to be reaquired not computed, passing existing and daily data only!\n")
                    visualize_data(qc_data)
            elif (answer == '3'):
                try:
                    update_qcsda(points_to_reaquire, daily_amounts['daily_report']['date'], data_source)
                except UnboundLocalError:
                    print("\nPlease get points to reacquire first (run option 1) before updating\n")
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


print("")
print("-----------------------------------------------------------------")
print("Welcome to your Daily Quality Control of Seismic Dcquisition Data")
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

run_program = input('Select option: \n')

if (run_program == "G" or run_program == "g" or
    run_program == "L" or run_program == "l"):
    main(run_program)
else:
    print("Program closed.")