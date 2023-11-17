import PySimpleGUI as sg
import datetime
import pymongo
import pandas as pd

sg.theme("DarkPurple1")

# Connect to your MongoDB database (make sure to use the same database name and collection name in both scripts)
client = pymongo.MongoClient('mongodb+srv://Masterskimaru:Masterkimaru@equipment.rld5bgd.mongodb.net/')  # Replace with your MongoDB connection string
db = client['EquipmentDataEntry']  # Replace with your shared database name
collection = db['Equipment']  # Replace with your shared collection name

# Define the Excel file name
EXCEL_FILE = 'NewEntry.xlsx'

# Specify the 'openpyxl' engine when reading the Excel file
df = pd.read_excel(EXCEL_FILE, engine='openpyxl')

# Initialize equipment type counters
equipment_type_counters = {}  # For tracking the equipment type and associated counter

# Define default dates
default_lpo_date = datetime.date.today().strftime('%Y-%b-%d')
default_supplier_date = datetime.date.today().strftime('%Y-%b-%d')

# Initialize the entry_counter to the maximum value in the existing data + 1
if not df.empty:
    max_existing_entries = df['No:(numbers)'].str.extract(r'(\d+)', expand=False)
    max_existing_entries = pd.to_numeric(max_existing_entries, errors='coerce')  # Convert to numeric, handling NaN
    max_existing_entries = max_existing_entries.dropna()
    if not max_existing_entries.empty:
        max_existing_entry = max_existing_entries.max()
        entry_counter = f'A{int(max_existing_entry) + 1:03d}'

    else:
        entry_counter = 'A001'
else:
    entry_counter = 'A001'


# Function to format the date in 'YYYY-MMM-DD' format
def format_date(date):
    return date.strftime('%Y-%b-%d')

def check_serial_number_uniqueness(serial_number):
    return collection.find_one({'Serial Number': serial_number}) is None

def clear_input():
    global entry_counter  # Declare entry_counter as a global variable

    # Store the default values for 'LPO Date' and 'Supplier Date'
    default_lpo_date = datetime.date.today().strftime('%Y-%b-%d')
    default_supplier_date = datetime.date.today().strftime('%Y-%b-%d')

    # Clear all other fields except for 'LPO Date' and 'Supplier Date'
    for key in values:
        if key != 'ChargerLaptop' and key not in ['LPO Date', 'Supplier Date']:
            # Set empty string for other fields
            window[key]('')

    # Set the default values for 'LPO Date' and 'Supplier Date'
    window['LPO Date'](default_lpo_date)
    window['Supplier Date'](default_supplier_date)

    window['Currency']('Ksh')
    window['VAT']('16%')
    window['Total Cost']('')

    # Increment the 'No:(numbers)' field for the next entry
    prefix, number = entry_counter[:2], int(entry_counter[2:])
    number += 1
    entry_counter = f'{prefix}{number:03d}'
    window['No:(numbers)'].update(entry_counter)

    return None

# Function to update the entry_counter based on the equipment type
def update_entry_counter(equipment_type):
    global entry_counter  # Declare entry_counter as a global variable

    if equipment_type not in equipment_type_counters:
        # If it's a new equipment type, initialize the counter for that type
        equipment_type_counters[equipment_type] = 1
    else:
        # If the equipment type already exists, increment its counter
        equipment_type_counters[equipment_type] += 1

    entry_counter = f'{equipment_type[:2]}{equipment_type_counters[equipment_type]:03d}'  # Update the global entry_counter

    return entry_counter

def equipment_field_callback(event, values):
    if event == 'Equipment':
        equipment = values['Equipment']
        equipment = equipment.upper()  # Convert the text to uppercase
        window['Equipment'].update(equipment)  # Update the field with the uppercase text

layout = [
    [sg.Text('Equipment Data Entry')],
    [sg.Text('No:(numbers)', size=(15, 1)), sg.Text(entry_counter, key='No:(numbers)'),
     sg.Text('', key='Status', text_color='red')],
    [sg.Text('Equipment', size=(15, 1)), sg.InputText(key='Equipment', enable_events=True, change_submits=True, size=(20, 1))],
    [sg.Text('Type of Model', size=(15, 1)), sg.InputText(key='Type of Model')],
    [sg.Text('Specifications', size=(15, 2)), sg.InputText(key='Specifications')],
    [sg.Text('Serial Number', size=(15, 1)), sg.InputText(key='Serial Number')],
    [sg.Text('Asset Tag', size=(15, 1)), sg.InputText(key='Asset Tag')],
    [sg.Text('Hostname or Location', size=(15, 2)), sg.InputText(key='Hostname', default_text='')],
    [sg.Text('Charger/Laptop', size=(15, 1)), sg.Checkbox('Both Taken', key='ChargerLaptop')],
    [sg.Text('LPO', size=(15, 2)), sg.InputText(key='LPO', default_text='')],
    [sg.Text('LPO Date', size=(15, 1)), sg.InputText(key='LPO Date', default_text=default_lpo_date),
     sg.CalendarButton('', target='LPO Date', format='%Y-%b-%d')],
    [sg.Text('Supplier', size=(15, 1)), sg.InputText(key='Supplier')],
    [sg.Text('Supplier Date', size=(15, 1)), sg.InputText(key='Supplier Date', default_text=default_supplier_date),
     sg.CalendarButton('', target='Supplier Date', format='%Y-%b-%d')],
    [sg.Text('Currency', size=(15, 1),
             tooltip='Select currency from the list'), sg.Combo(['Ksh', 'USD', 'EUR', 'GBP', 'JPY', 'CNY'], default_value='Ksh', key='Currency')],
    [sg.Text('Cost', size=(15, 1)), sg.Input(key='Cost')],
    [sg.Text('VAT', size=(15, 1),
             tooltip='Enter VAT as a percentage'), sg.Input(key='VAT', default_text='16%', enable_events=True, justification='left', size=(15, 1), pad=((20, 0), 0), text_color='black', background_color='white')],
    [sg.Text('Total Cost', size=(15, 1)), sg.Input(key='Total Cost')],
    [sg.Text('Remarks', size=(15, 1)), sg.Multiline(key='Remarks', size=(30, 3))],
    [sg.Text('Warranty', size=(15, 1)), sg.InputText(key='Warranty')],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

# Create the column
column = sg.Column(layout, vertical_scroll_only=True, scrollable=True, size=(600, 600))

# Create the window and add the column
window = sg.Window('Equipment data entry form', [[column]])

def calculate_total_cost(values):
    try:
        cost = float(values['Cost'])
        vat = float(values['VAT'].rstrip('%')) / 100
        total_cost = cost + (cost * vat)
        window['Total Cost'].update(f'{total_cost:.2f}')
    except (ValueError, ZeroDivisionError):
        window['Total Cost'].update('Invalid Input')

values = {}  # Initialize values

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == "Exit":
        break

    if event == 'Clear':
        clear_input()

    if event == "Submit":
        serial_number = values['Serial Number']
        equipment_type = values['Equipment']

        # Check if any of the mandatory fields are empty
        mandatory_fields = ['Equipment', 'Type of Model', 'Specifications', 'Serial Number', 'Asset Tag', 'Supplier', 'Warranty', 'Remarks', 'Total Cost', 'Cost']
        if any(values[field] == '' for field in mandatory_fields):
            window['Status'].update('Please fill in all mandatory fields.', text_color='red')
        else:
            if not check_serial_number_uniqueness(serial_number):
                window['Status'].update('Serial number already in use. Please use a different serial number.', text_color='red')
            else:
                # Update the 'No:(numbers)' field based on equipment type and unique number
                entry_counter = update_entry_counter(equipment_type)
                window['No:(numbers)'].update(entry_counter)

                # Exclude the 'Calendar' values when adding to the DataFrame
                data_to_add = {key: value if key != 'LPO Date' and key != 'Supplier Date' else default_lpo_date for key, value in values.items() if key != 'Calendar'}
                data_to_add['No:(numbers)'] = entry_counter  # Add the entry_counter to the data

                df = pd.concat([df, pd.DataFrame([data_to_add])], ignore_index=True)

                # Remove unwanted columns: 'Calendar0', unnamed column, and '0'
                df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
                df = df.drop(['Calendar0', '0'], axis=1, errors='ignore')

                # Save the updated DataFrame to the Excel file, including the No:(numbers)
                df.to_excel(EXCEL_FILE, index=False)

                form_data = {
                    'No:(numbers)': equipment_type_counters[equipment_type],  # Update the MongoDB entry with the correct No:(numbers)
                    'Serial Number': serial_number,
                    'Equipment': values['Equipment'],
                    'Type of Model': values['Type of Model'],
                    'Specifications': values['Specifications'],
                    'Asset Tag': values['Asset Tag'],
                    'Hostname': values['Hostname'],
                    'Charger/Laptop': values['ChargerLaptop'],
                    'LPO': values['LPO'],
                    'LPO Date': values['LPO Date'],
                    'Supplier': values['Supplier'],
                    'Supplier Date': values['Supplier Date'],
                    'Currency': values['Currency'],
                    'Cost': values['Cost'],
                    'VAT': values['VAT'],
                    'Total Cost': values['Total Cost'],
                    'Remarks': values['Remarks'],
                    'Warranty': values['Warranty'],
                }

                # Use 'update_one' to ensure the Serial Number is unique
                collection.update_one({'Serial Number': serial_number}, {'$set': form_data}, upsert=True)
                sg.popup('Data saved to MongoDB')

                clear_input()

    if event in ['Cost', 'VAT']:
        calculate_total_cost(values)

window.close()
