import PySimpleGUI as sg
import subprocess
import pandas as pd
import pymongo

sg.theme("DarkPurple1")

# Connect to your MongoDB database
client = pymongo.MongoClient('mongodb+srv://Masterskimaru:Masterkimaru@equipment.rld5bgd.mongodb.net/')
db = client['EquipmentDataEntry']
collection = db['Equipment']

# Dictionary that maps form names to data files
form_data_files = {
    'Equipment Entry Form': 'NewEntry.xlsx',
    'Editable Form': 'EditableForm.xlsx',
    'Disposal Form': 'DisposalForm.xlsx',
}

# Additional features
info_text = sg.Text('Select a Form to Open', size=(30, 1))
status_text = sg.Text('', size=(30, 1))
help_button = sg.Button('Help', size=(30, 1))
clear_button = sg.Button('Clear', size=(30, 1))
exit_button = sg.Button('Exit', size=(30, 1))
view_data_button = sg.Button('View Form Data', size=(30, 1))
generate_report_button = sg.Button('Generate Report', size=(30, 1))
count_button = sg.Button('Count', size=(30, 1))

# Create a Multiline element for displaying data
data_display = sg.Multiline('', size=(100, 40), key='-DATA-', autoscroll=True, text_color='white', background_color='black')

layout = [
    [info_text],
    [sg.Button('Open Equipment Entry Form', size=(30, 1))],
    [sg.Button('Open Editable Form', size=(30, 1))],
    [sg.Button('Open Disposal Form', size=(30, 1))],
    [status_text],
    [view_data_button],
    [generate_report_button],
    [count_button],
    [help_button, clear_button, exit_button],
    [data_display],  # Add the Multiline element to the layout
]

window = sg.Window('Search Form', layout, finalize=True, size=(800, 600))

# Report generation functions and MongoDB connection code go here...

def generate_full_report():
    sg.popup('Generating Full Report...')
    full_report_data = list(collection.find({}))  # Retrieve all documents from MongoDB

    if full_report_data:
        full_report_df = pd.DataFrame(full_report_data)
        full_report_df.to_excel('FullReport.xlsx', index=False)
        sg.popup('Full Report generated and saved as FullReport.xlsx')
    else:
        sg.popup('No data found for Full Report.')

def generate_users_report():
    sg.popup('Generating Users Report...')
    user_fields = ['Username', 'Equipment', 'DateOfIssue1', 'Hostname']
    users_report_data = list(collection.find({}, projection={field: 1 for field in user_fields}))

    if users_report_data:
        users_report_df = pd.DataFrame(users_report_data)
        users_report_df.to_excel('UsersReport.xlsx', index=False)
        sg.popup('Users Report generated and saved as UsersReport.xlsx')
    else:
        sg.popup('No data found for Users Report.')

def generate_equipment_report():
    sg.popup('Generating Equipment Report...')
    equipment_fields = ['Equipment', 'Type of Model', 'Asset Tag', 'Supplier', 'Total Cost']
    equipment_report_data = list(collection.find({}, projection={field: 1 for field in equipment_fields}))

    if equipment_report_data:
        equipment_report_df = pd.DataFrame(equipment_report_data)
        equipment_report_df.to_excel('EquipmentReport.xlsx', index=False)
        sg.popup('Equipment Report generated and saved as EquipmentReport.xlsx')
    else:
        sg.popup('No data found for Equipment Report.')

def generate_disposal_report():
    sg.popup('Generating Disposal Report...')
    disposal_fields = ['Equipment', 'Serial Number', 'Type of Model']
    disposal_report_data = list(collection.find({}, projection={field: 1 for field in disposal_fields}))

    if disposal_report_data:
        disposal_report_df = pd.DataFrame(disposal_report_data)
        disposal_report_df.to_excel('DisposalReport.xlsx', index=False)
        sg.popup('Disposal Report generated and saved as DisposalReport.xlsx')
    else:
        sg.popup('No data found for Disposal Report.')
# Event loop with the "Count" button handling code...
while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    if event == 'Open Equipment Entry Form':
        subprocess.run(["python", "Equipment_entry.py"])
        status_text.update('Equipment Entry Form is open.')

    if event == 'Open Editable Form':
        subprocess.run(["python", "EditableForm.py"])
        status_text.update('Editable Form is open.')

    if event == 'Open Disposal Form':
        subprocess.run(["python", "DisposalForm.py"])
        status_text.update('Disposal Form is open.')

    if event == 'View Form Data':
        # Show a dialog to select which form's data to view
        form_selection = sg.popup_get_text('Select a Form to View Data', title='Select Form', default_text='Equipment Entry Form', size=(30, 1))

        if form_selection:
            selected_form = form_selection
            data_file = form_data_files.get(selected_form)

            if data_file:
                data = pd.read_excel(data_file)

                data_text = f'Viewing Data for: {selected_form}\n\n'
                for index, row in data.iterrows():
                    data_text += f'{row}\n'

                # Update the Multiline element with the data
                data_display.update(data_text)
                
                
    if event == 'Count':
        # Connect to your MongoDB collection
        user_count = collection.count_documents({})  # Total number of users

        equipment_count = collection.count_documents({"Equipment": {"$exists": True}})  # Total number of equipment

        # Group by "Type of Model" and count the number of equipment for each type
        equipment_type_counts = collection.aggregate([
            {"$match": {"Equipment": {"$exists": True}}},
            {"$group": {"_id": "$Type of Model", "count": {"$sum": 1}}}
        ])

        count_dialog_text = f"Total Number of Users: {user_count}\n"
        count_dialog_text += f"Total Number of Equipment: {equipment_count}\n"
        count_dialog_text += "Total Number of Equipment by Type:\n"

        for type_count in equipment_type_counts:
            count_dialog_text += f"{type_count['_id']}: {type_count['count']}\n"

        sg.popup("Count Information", count_dialog_text)


    if event == 'Generate Report':
        report_selection = sg.popup_get_text('Select a Report Type', title='Select Report',
                                            default_text='Full Report|Users Report|Equipment Report|Disposal Report', size=(30, 1))

        if report_selection:
            selected_report = report_selection.lower()

            if selected_report == 'full report':
                generate_full_report()
            elif selected_report == 'users report':
                generate_users_report()
            elif selected_report == 'equipment report':
                generate_equipment_report()
            elif selected_report == 'disposal report':
                generate_disposal_report()

    if event == 'Help':
        sg.popup('Help Information:\n\n1. Select a form to open.\n2. Click "Exit" to close the application.\n3. Click "View Form Data" to view data from a specific form.')

    if event == 'Clear':
        info_text.update('Select a Form to Open')
        status_text.update('')
        window.FindElement('Open Equipment Entry Form').Update(disabled=False)
        window.FindElement('Open Editable Form').Update(disabled=False)
        window.FindElement('Open Disposal Form').Update(disabled=False)


# Close the MongoDB connection
client.close()

window.close()
