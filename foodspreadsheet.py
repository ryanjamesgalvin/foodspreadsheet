import PySimpleGUI as sg
import gspread
import gspread_formatting as gsf
import os
#connect to the service account which is used to edit the spreadsheet
os.chdir("C:\\Users\\balvin\\Desktop\\code stuff\\food spreadsheet")
gc = gspread.service_account(filename="spreadsheet_credentials.json")

#connect to the sheet 
sh = gc.open("Kristen & Ryan's Restaurant Experiences").sheet1

# grab all data from current spreadsheet in a list of dictionaries
data = sh.get_all_records(head=3,expected_headers=sh.row_values(3))

#to do - run through this data and sort all items for same restaurant

# put all unique existing Restaurant names into a list
restaurantNames = []
for item in data:
    if item.get("Restaurant") not in restaurantNames and item.get("Restaurant") != "":
        restaurantNames.append(item.get("Restaurant"))

# create the initial layout for the GUI window
layout = [[sg.Text("This is a simple tool to easily add new entries to the \"Ryan & Kristen's restaurant experiences\" spreadsheet.", font=("Calibri", 24, "bold"))],
           [sg.Text("Select an existing restaurant to add new menu items to below:", font=("Calibri", 24))],
           [sg.Combo(restaurantNames, default_value=restaurantNames[0], enable_events=True, key='existing_restaurant', font=(12))],
           [sg.Text('Enter the menu item name')],
           [sg.Input(default_text="", key='menu_item')],
           [sg.Text("Enter Ryan's Rating (0-10)")],
           [sg.Input(default_text="", key='ryans_rating')],
           [sg.Text("Enter Kristen's Rating (0-10)")],
           [sg.Input(default_text="", key='kristens_rating')],
           [sg.Text("Enter Ryan's Thoughts (Optional)")],
           [sg.Input(default_text="", key='ryans_thoughts')],
           [sg.Text("Enter Kristen's Thoughts (Optional)")],
           [sg.Input(default_text="", key='kristens_thoughts')],
           [sg.Button('Add to selected existing restaurant')],
           [sg.Text("Or enter a new restaurant below:", font=("Calibri", 24))],
           [sg.Text("Enter a new restaurant to be added (using the above data)")], 
           [sg.Text("New Restaurant Name")],
           [sg.Input(default_text="", key='new_restaurant')],
           [sg.Text("New Restaurant Type")],
           [sg.Input(default_text="", key='new_restaurant_type')],
           [sg.Text("New Restaurant URL")],
           [sg.Input(default_text="", key='new_restaurant_URL')],
           [sg.Button('Add new restaurant'), sg.Button('Cancel')]]

# initialize window 
window = sg.Window("Ryan & Kristen's Food Spreadsheet Addition Tool", layout)

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
    elif event == 'Add to selected existing restaurant':
        # check to make sure all required values are entered
        if values['menu_item'] == "":
                sg.popup("Please make sure all values that are not marked optional are filled out before adding a new row.", title="Unfilled Values Error")
                continue
    
        #find the cell of the existing restaurant the user selected
        restaurant_cell = sh.find(values['existing_restaurant'])
        
        #iterate down the sheet to find where to add the menu item 
        for rows in sh.get_all_values():
            if rows[0] == values['existing_restaurant']: # found the row where the existing restaurant starts

                # insert the new row after the restaurant's row, inherting the formatting
                try:
                    sh.insert_row(["","",values['menu_item'],int(values['ryans_rating']),
                                int(values['kristens_rating']),values['ryans_thoughts'],
                                values['kristens_thoughts']], index=(restaurant_cell.row+1),
                                inherit_from_before=True)
                except ValueError:
                    sg.popup("Please enter a number for Ryan & Kristen's rating, from 0-10")
                    continue
                sg.popup("Row added successfully!")
                break # done adding the row, so no reason to continue
            else:
                continue

    elif event == 'Add new restaurant':
        try:
            # check to make sure all required values are entered           
            if values['new_restaurant_URL'] == "" or values['new_restaurant'] == "" or values['new_restaurant_type'] == "" or values['menu_item'] == "":
                sg.popup("Please make sure all values that are not marked optional are filled out before adding a new row.", title="Unfilled Values Error")
                continue
            #add the black row that separates new restaurants
            separatorRow = sh.append_row(values=[]) # return dictionary with information about added row

            #format the black row
            blackFormat = gsf.cellFormat(backgroundColor=gsf.color(0, 0, 0))
            gsf.format_cell_range(sh, separatorRow['updates']['updatedRange'][8:], cell_format=blackFormat)
            #get the location of the added row, so that the below "append_row" call adds to the row after the separator
            newRowLocation = int(separatorRow['updates']['updatedRange'][8:]) + 1
            # add new row to the end of the sheet
            newRestaurantRow = sh.append_row(["=HYPERLINK("+'"' + values['new_restaurant_URL'] + '","' + values['new_restaurant']+ '")', values['new_restaurant_type'] ,values['menu_item'],int(values['ryans_rating']),
                            int(values['kristens_rating']),values['ryans_thoughts'],
                            values['kristens_thoughts']], value_input_option="USER_ENTERED", table_range="A" + str(newRowLocation))
            #format the new restaurant row with size 14 font and italic
            italicFormat = gsf.cellFormat(textFormat=gsf.textFormat(italic=True, fontSize=14))
            gsf.format_cell_range(sh, "A"+str(newRowLocation), cell_format=italicFormat)
            # add the new restaurant to the list of existing restaurants
            restaurantNames.append(values['new_restaurant'])
            
        except ValueError:
            sg.popup("Please enter a number for Ryan & Kristen's rating, from 0-10")
            continue
        sg.popup("Row added successfully!")
        
        # to do - figure out way to refresh window with new restaurant


window.close()
