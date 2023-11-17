import PySimpleGUI as sg
import datetime
import pymongo
import pandas as pd

sg.theme("DarkPurple1")

# Connect to your MongoDB database
client = pymongo.MongoClient('mongodb+srv://Masterskimaru:Masterkimaru@equipment.rld5bgd.mongodb.net/')
db = client['EquipmentDataEntry']
collection = db['Equipment']

# Connect to an Excel file
EXCEL_FILE = 'EditableForm.xlsx'
df = pd.read_excel(EXCEL_FILE)

disposal_options = ['Auction', 'Re-sell in parts', 'Take for Recycling', 'Repair']  # Available disposal options

# Initialize sub-fields as hidden
sub_fields_visible = False

layout = [
    [sg.Text('Fill out the following fields of the editing form')],
    [sg.Text('Serial Number', size=(15, 1)), sg.InputText(key='Serial Number'), sg.Button('Retrieve Details')],
    [sg.Text('Equipment', size=(15, 1)), sg.InputText(key='Equipment', visible=sub_fields_visible)],
    [sg.Text('Asset Tag', size=(15, 1)), sg.InputText(key='Asset Tag', visible=sub_fields_visible)],
    [sg.Text('Type of Model', size=(15, 1)), sg.InputText(key='Type of Model', visible=sub_fields_visible)],
    [sg.Text('Username', size=(15, 1)), sg.InputText(key='Username')],
    [sg.Text('Date of Issue 1', size=(15, 1)), sg.InputText(key='DateOfIssue1', default_text=datetime.date.today().strftime('%Y-%b-%d')),
     sg.CalendarButton('Calendar', target='DateOfIssue1', format='%Y-%b-%d')],
    [sg.Text('Remarks', size=(15, 1)), sg.Multiline(key='Remarks', size=(30, 5))],

    [sg.Text('CONFIGURATIONS', size=(15, 1)), sg.Combo(['Choose Configuration', 'Change of Location', 'Change of User', 'Add User', 'Remove User'], key='Configurations', enable_events=True)],
    [sg.Button('Open Configuration')],
    [sg.Submit(), sg.Exit()],
]

window = sg.Window('Editable Form Entry', layout)

def retrieve_item_details(serial_number):
    global sub_fields_visible  # Declare as global
    item_details = collection.find_one({'Serial Number': serial_number})
    if item_details:
        sg.popup('Item Details:', item_details)
        if not sub_fields_visible:
            window['Asset Tag'].update(item_details.get('Asset Tag', ''))
            window['Type of Model'].update(item_details.get('Type of Model', ''))
            window['Equipment'].update(item_details.get('Equipment', ''))
            # Show and populate sub-fields
            window['Asset Tag'](visible=True)
            window['Type of Model'](visible=True)
            window['Equipment'](visible=True)
            sub_fields_visible = True
    else:
        sg.popup('Item not found.')

# Function to handle Change of Location
def change_location(serial_number):
    item_details = collection.find_one({'Serial Number': serial_number})
    if item_details:
        current_location = item_details.get('Location', '')
        new_location = sg.popup_get_text('Change Location to:', default_text=current_location)
        if new_location:
            if sg.popup_yes_no('Accept and Save changes?') == 'Yes':
                # Update both "Location" and "Hostname"
                collection.update_one({'Serial Number': serial_number}, {'$set': {'Location': new_location, 'Hostname': new_location}})
                sg.popup('Location and Hostname changed and saved to the database.')

def change_user(serial_number):
    item_details = collection.find_one({'Serial Number': serial_number})
    if item_details:
        current_user = item_details.get('Username', '')
        new_user = sg.popup_get_text('Change User to:', default_text=current_user)
        if new_user:
            if sg.popup_yes_no('Accept and Save changes?') == 'Yes':
                current_date = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                # Create a new user entry
                new_user_entry = {
                    'Equipment': item_details['Equipment'],
                    'Serial Number': item_details['Serial Number'],
                    'Username': new_user,
                    'DateOfIssue1': current_date,
                    'Remarks': "User changed"
                }
                # Update the 'Users' field
                users = item_details.get('Users', [])
                users.append(new_user_entry)
                collection.update_one({'Serial Number': serial_number}, {'$set': {'Users': users, 'Username': new_user}})
                sg.popup('User changed and saved to the database.')


# Function to handle Add User
def add_user(serial_number):
    item_details = collection.find_one({'Serial Number': serial_number})
    if item_details:
        current_users = [user['Username'] for user in item_details.get('Users', [])]
        new_user = sg.popup_get_text('Add User:', default_text="")
        if new_user:
            if new_user in current_users:
                sg.popup("Username already exists.")
            else:
                current_date = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                new_user_entry = {
                    'Equipment': item_details['Equipment'],
                    'Serial Number': item_details['Serial Number'],
                    'Username': new_user,
                    'DateOfIssue1': current_date,
                    'Remarks': "User added"
                }
                current_users.append(new_user)
                # Update the 'Users' field and the 'Username' field
                users = item_details.get('Users', [])
                users.append(new_user_entry)
                collection.update_one({'Serial Number': serial_number}, {'$set': {'Users': users, 'Username': new_user}})
                sg.popup('User added and saved to the database.')


# Function to handle Remove User
def remove_user(serial_number):
    item_details = collection.find_one({'Serial Number': serial_number})
    if item_details:
        current_users = [user['Username'] for user in item_details.get('Users', [])]
        if not current_users:
            sg.popup("No User Present")
        else:
            remove_user = sg.popup_get_text('Remove User:', default_text="")
            if remove_user in current_users:
                if sg.popup_yes_no('Are you sure you want to delete the user?') == 'Yes':
                    # Remove the user from the 'Users' field
                    updated_users = [user for user in item_details.get('Users', []) if user['Username'] != remove_user]
                    # Update the 'Users' field and the 'Username' field
                    updated_username = updated_users[0]['Username'] if updated_users else ''
                    collection.update_one({'Serial Number': serial_number}, {'$set': {'Users': updated_users, 'Username': updated_username}})
                    sg.popup('User removed and saved to the database.')
                else:
                    sg.popup('User not deleted.')
            else:
                sg.popup('No Username like that.')



while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Submit':
        # Check if any of the mandatory fields are empty
        mandatory_fields = ['Equipment', 'Serial Number', 'Username', 'DateOfIssue1']
        if any(values[field] == '' for field in mandatory_fields):
            sg.popup_error('Please fill in all mandatory fields except Date of Return.')
        else:
            form_data = {
                'Equipment': values['Equipment'],
                'Serial Number': values['Serial Number'],
                'Username': values['Username'],
                'DateOfIssue1': values['DateOfIssue1'],
                'Remarks': values['Remarks'],
            }

            # Update the DataFrame with the new data
            new_data = pd.DataFrame([form_data])
            df = pd.concat([df, new_data], ignore_index=True)

            # Save the updated DataFrame to the Excel file
            df.to_excel(EXCEL_FILE, index=False)
            sg.popup('Data saved to Excel sheet')

            # Save the data to MongoDB
            collection.update_one({'Serial Number': values['Serial Number']}, {'$push': {'Users': form_data}}, upsert=True)
            sg.popup("Data saved to MongoDB.")

            # Clear the form and return it to its original state
            window['Equipment']('')
            window['Serial Number']('')
            window['Username']('')
            window['DateOfIssue1'](datetime.date.today().strftime('%Y-%b-%d'))
            window['Remarks']('')
            sub_fields_visible = False
            window['Asset Tag'](visible=False)
            window['Type of Model'](visible=False)
            window['Equipment'](visible=False)

    if event == 'Retrieve Details':
        retrieve_serial_number = values['Serial Number']
        if retrieve_serial_number:
            retrieve_item_details(retrieve_serial_number)
    if event == 'Open Configuration':
        configuration = values['Configurations']
        if configuration == 'Change of Location':
            change_location(values['Serial Number'])
        elif configuration == 'Change of User':
            change_user(values['Serial Number'])
        elif configuration == 'Add User':
            add_user(values['Serial Number'])
        elif configuration == 'Remove User':
            remove_user(values['Serial Number'])

window.close()

# Close the MongoDB connection
client.close()
