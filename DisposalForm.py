import PySimpleGUI as sg
import pymongo
import pandas as pd

sg.theme("DarkPurple1")

# Connect to your MongoDB database
client = pymongo.MongoClient('mongodb+srv://Masterskimaru:Masterkimaru@equipment.rld5bgd.mongodb.net/')
db = client['EquipmentDataEntry']
collection = db['Equipment']

# Connect to an Excel file
EXCEL_FILE = 'DisposalForm.xlsx'
df = pd.read_excel(EXCEL_FILE)

disposal_options = ['Auction', 'Re-sell in parts', 'Take for Recycling', 'Repair']  # Available disposal options

# Define the layout, including the hidden subfields for Asset Tag and Type of Model
layout = [
    [sg.Text('Fill out the following fields of the Disposal form')],
    [sg.Text('Serial Number', size=(15, 1)), sg.InputText(key='Serial Number'), sg.Button('Retrieve Details')],
    [sg.Text('Equipment', size=(15, 1)), sg.InputText(key='Equipment', visible=False)],
        # Titles for the hidden subfields
    [sg.Text('Asset Tag', size=(15, 1)),  sg.Text('', size=(30, 1), key='Asset Tag')],
    [sg.Text('Type of Model', size=(15, 1)), sg.Text('', size=(30, 1), key='Type of Model')],    
    
    [sg.Text('Replace', size=(15, 1)), sg.Checkbox('Replace', key='Replace')],
    [sg.Text('Disposal', size=(15, 1)), sg.DropDown(disposal_options, key='Disposal', readonly=True)],
    [sg.Text('Remarks', size=(15, 1)), sg.Multiline(key='Remarks', size=(30, 3))],


    
    [sg.Submit(), sg.Button('Clear'), sg.Exit()],
]

window = sg.Window('Disposal Form Entry', layout)

def retrieve_item_details(serial_number):
    item_details = collection.find_one({'Serial Number': serial_number})
    if item_details:
        sg.popup('Item Details:', item_details)
        populate_form_fields(item_details)

        # Show and populate the subfields and their titles
        window['Asset Tag'].update(visible=True)
        window['Type of Model'].update(visible=True)
        window['Equipment'].update(visible=True)

        window['Asset Tag'].update(item_details.get('Asset Tag', ''))
        window['Type of Model'].update(item_details.get('Type of Model', ''))
        window['Equipment'].update(item_details.get('Equipment', ''))

        # Update the DataFrame with MongoDB data
        update_data_to_dataframe(item_details)
    else:
        sg.popup('Item not found.')

def populate_form_fields(item_details):
    equipment = item_details.get('Equipment', '')
    replace = 'YES' if item_details.get('Replace') == 'YES' else 'NO'
    disposal = item_details.get('Disposal', '')
    

    window['Equipment'].update(equipment)
    window['Replace'].update(replace)
    window['Disposal'].update(disposal)
   

# Create a function to update the DataFrame with MongoDB data
def update_data_to_dataframe(mongodb_data):
    global df
    new_data = {
        'Equipment': mongodb_data.get('Equipment', ''),
        'Serial Number': mongodb_data.get('Serial Number', ''),
        'Replace': 'YES' if mongodb_data.get('Replace') == 'YES' else 'NO',
        'Disposal': mongodb_data.get('Disposal', ''),
        'Remarks': mongodb_data.get('Remarks', ''),
        'Asset Tag': mongodb_data.get('Asset Tag', ''),
        'Type of Model': mongodb_data.get('Type of Model', ''),
    }
    df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

# Rest of the code remains the same

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Submit':
        form_data = {
            'Equipment': values['Equipment'],
            'Serial Number': values['Serial Number'],
            'Replace': 'YES' if values['Replace'] else 'NO',
            'Disposal': values['Disposal'],
            'Remarks': values['Remarks'],
            'Asset Tag': values['Asset Tag'],
            'Type of Model': values['Type of Model'],
        }

        collection.update_one({'Serial Number': values['Serial Number']}, {'$set': form_data}, upsert=True)
        sg.popup('Data saved in database')
        update_data_to_dataframe(form_data)
        sg.popup('Data saved in Excel sheet')

    if event == 'Retrieve Details':
        retrieve_serial_number = values['Serial Number']
        if retrieve_serial_number:
            retrieve_item_details(retrieve_serial_number)

    if event == 'Clear':
        window['Equipment'].update('')
        window['Serial Number'].update('')
        window['Replace'].update(False)
        window['Disposal'].update('')
        window['Remarks'].update('')
        window['Asset Tag'].update(visible=False)
        window['Type of Model'].update(visible=False)
        window['Equipment'].update(visible=False)

# Close the MongoDB connection
client.close()

window.close()
