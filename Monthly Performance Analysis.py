import pyodbc, datetime, os, time, re, warnings, sys
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from sklearn.linear_model import LinearRegression
import matplotlib.pyplot as plt
from matplotlib.widgets import Slider

from bs4 import BeautifulSoup

from icecream import ic
#Login to Google Script Group
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError



SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
paSheet = '1EHQgMAh9q3vt0ZiMXDkdQxZ8oEQJ-L1EVUhbLHa7QhI'
woSheet = '1bxGKjFCgN-T1hxuEm3NpC7W5eyQpJo6nV8WPldHD184'

testpaSheet = '1wfEawAJZ9KvLNijJQPvLGmStq8yw6DeIPikx11oU1oc'
testwoSheet = '1aasuRuI9YT8RZvi8GqhV0K_FgmiBdiYBqjUpxW1ckrI'


credentials = None
if os.path.exists("PyPA-token.json"):
    credentials = Credentials.from_authorized_user_file("PyPA-token.json", SCOPES)
if not credentials or not credentials.valid:
    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(r"G:\Shared drives\O&M\NCC Automations\Daily Automations\NCC-AutomationCredentials.json", SCOPES)
        credentials = flow.run_local_server(port=0)
    with open("PyPA-token.json", "w+") as token:
        token.write(credentials.to_json())

service = build('sheets', 'v4', credentials=credentials)
#End Group




#Stores the data we put into the report, Is global
dfs = {}


# Suppress specific warnings from openpyxl
# WARNING FOLLOWS:
#C:\Users\OMOPS.AzureAD\AppData\Local\Programs\Python\Python312-32\Lib\site-packages\openpyxl\worksheet\_reader.py:329: UserWarning: Unknown extension is not supported and will be removed
#  warn(msg)
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl') #Prevents warning from printing

pvsystDB = r"G:\Shared drives\O&M\NCC Automations\Performance Reporting\PVsyst (Josephs Edits).accdb"
xl_sheets = ['Bluebird Solar', 'Cardinal', 'Cherry Blossom Solar, LLC', 'Cougar Solar, LLC', 'Harrison Solar', 'Hayes', 'Hickory Solar, LLC', 'Violet Solar, LLC', 'Wellons Solar, LLC']
#intialize dictionaries
for sheet in xl_sheets:
    dfs[sheet] = {}

report_sheets = ['Bluebird', 'Cardinal', 'Cherry', 'Cougar', 'Harrison', 'Hayes', 'Hickory', 'Violet', 'Wellons']
wo_Only_sites = ['Bulloch 1A', 'Bulloch 1B', 'Elk', 'Freight Line', 'Gray Fox', 'Harding', 'Holly Swamp', 'Mclean', 'PG Solar', 'Richmond Cadle', 'Shorthorn', 'Sunflower', 'Upson', 'Warbler', 'Washington', 'Whitehall', 'Whitetail']
for site in wo_Only_sites:
    dfs[site] = {}
wo_sites = ['Bluebird', 'Cardinal', 'Cherry Blossom', 'Cougar', 'Harrison', 'Hayes', 'Hickory', 'Violet', 'Wellons Farm']
def dbcnxn():
    global db, connect_db, c
    #Connect to DB
    db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\O&M\NCC Automations\Performance Reporting\PVsyst (Josephs Edits).accdb;'
    connect_db = pyodbc.connect(db)
    c = connect_db.cursor()


def plot_poa_data(dates, predicted_poa, actual_poa, modified_poa, site):
    fig, ax = plt.subplots(figsize=(10, 6))
    plt.subplots_adjust(bottom=0.25)

    # Plot the data
    ax.plot(dates, predicted_poa, label='Predicted POA', color='blue')
    ax.plot(dates, actual_poa, label='Actual Found', color='orange')
    ax.plot(dates, modified_poa, label="Modified POA", color='green')

    ax.set_title(f"POA Data for {site}")
    ax.set_xlabel("Date")
    ax.set_ylabel("POA Value")
    ax.legend()

    # Create a horizontal slider to scroll through the data
    ax_slider = plt.axes([0.1, 0.1, 0.8, 0.03], facecolor='lightgoldenrodyellow')
    slider = Slider(ax_slider, 'Scroll', 0, len(dates) - 1, valinit=0, valstep=1)

    def update(val):
        pos = slider.val
        ax.set_xlim(dates[int(pos)], dates[min(int(pos) + 30, len(dates) - 1)])
        fig.canvas.draw_idle()

    slider.on_changed(update)

    plt.show()





def show_poa_selection(predicted_poa_sum, sum_poa, modified_poa_sum, site, dates, predicted_poa_data, modified_poa_data, actual_poa_data):
    if plot_option_var.get() == True:
        plot_poa_data(dates, predicted_poa_data, actual_poa_data, modified_poa_data, site)


    def on_ok():
        selected_value = var.get()
        global sum_poa_chosen
        sum_poa_chosen = selected_value
        poa_win.destroy()

    poa_win = tk.Toplevel()
    poa_win.title(f"POA: {site}")
    poa_win.geometry('300x100+100-700')

    var = tk.DoubleVar()

    # Determine the largest value
    largest_value = max(predicted_poa_sum, sum_poa, modified_poa_sum)
    var.set(largest_value)

    # Create check buttons
    check_buttons = [
        tk.Checkbutton(poa_win, text=f"Predicted POA: {round(predicted_poa_sum, 0)}", variable=var, onvalue=predicted_poa_sum),
        tk.Checkbutton(poa_win, text=f"Actual Found: {round(sum_poa, 0)}", variable=var, onvalue=sum_poa),
        tk.Checkbutton(poa_win, text=f"0's and -'s replaced by Predicted: {round(modified_poa_sum, 0)}", variable=var, onvalue=modified_poa_sum)
    ]
    # Set the background color of the check button with the largest value to light greenYEAR(TODAY())
    for check_button in check_buttons:
        if check_button.cget("onvalue") == largest_value:
            check_button.config(bg="light green")
        check_button.pack(anchor='w')
    # Create OK button
    tk.Button(poa_win, text="OK", command=on_ok).pack(fill='x')
    poa_win.wait_window()


def get_sheet_id(sheet_name):
    # Retrieve the spreadsheet metadata to get the sheet ID
    if sheet_name in wo_Only_sites or any(sheet_name == site + ' Quarterly' for site in wo_Only_sites):
        spreadass = woSheet
    else:
        spreadass = paSheet
    spreadsheet = service.spreadsheets().get(spreadsheetId=spreadass).execute()
    sheets = spreadsheet.get('sheets', '')

    for sheet in sheets:
        if sheet.get("properties", {}).get("title") == sheet_name:
            return sheet.get("properties", {}).get("sheetId")

    raise ValueError(f"Sheet {sheet_name} not found.")

def clear_cells(sheet_name):
    # Clear specific cells if the month is February
    current_month = datetime.datetime.now().month
    if current_month == 2:
        # Define the ranges to clear
        ranges_to_clear = [f"{sheet_name}!B20:B31", f"{sheet_name}!C20:C31", f"{sheet_name}!E20:E31", f"{sheet_name}!K20:K31"]

        # Prepare the requests to clear the specified ranges
        requests = [{
            'updateCells': {
                'range': {
                    'sheetId': get_sheet_id(sheet_name),
                    'startRowIndex': 22,  # 0-based index, so 20th row is index 19
                    'endRowIndex': 35,    # 32 is exclusive, so it clears up to 31
                    'startColumnIndex': col_index,
                    'endColumnIndex': col_index + 1
                },
                'fields': 'userEnteredValue'
            }
        } for col_index in [1, 2, 4, 10]]  # B, C, E, K are columns 1, 2, 4, 10 in 0-based index

        # Execute the batch update to clear the cells
        body = {
            'requests': requests
        }

        service.spreadsheets().batchUpdate(
            spreadsheetId=paSheet,
            body=body
        ).execute()
        time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)

def delete_rows(sheet_name, woVpa):
    if woVpa == True:
        spread = paSheet
        begin_row = 36
        s_row = 40
    else:
        spread = woSheet
        begin_row = 7
        s_row = 12
    # Define the range to search for the start and end markers
    search_range = f"{sheet_name}!G{begin_row}:K"

    # Read the data from the specified range
    result = service.spreadsheets().values().get(
        spreadsheetId=spread,
        range=search_range
    ).execute()
    time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)
    values = result.get('values', [])

    start_row = None
    end_row = None

    # Find the start and end rows
    for i, row in enumerate(values):
        if 'How Repaired' in row:
            start_row = i
        elif 'End of Reporting Record' in row:
            end_row = i
            break

    if start_row is not None and end_row is not None:
        # Calculate the number of rows to delete
        num_rows_to_delete = end_row - start_row

        if num_rows_to_delete > 0:
            # Use batchUpdate to delete rows
            requests = [{
                'deleteDimension': {
                    'range': {
                        'sheetId': get_sheet_id(sheet_name),
                        'dimension': 'ROWS',
                        'startIndex': start_row+begin_row,
                        'endIndex': end_row+begin_row-1
                    }
                }
            }]

            body = {
                'requests': requests
            }

            service.spreadsheets().batchUpdate(
                spreadsheetId=spread,
                body=body
            ).execute()
            time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)




    # Define the range to search for the start and end markers for the second deletion
    search_range_2 = f"{sheet_name}!E{s_row}:K"

    # Read the data from the specified range
    result_2 = service.spreadsheets().values().get(
        spreadsheetId=spread,
        range=search_range_2
    ).execute()
    time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)


    values_2 = result_2.get('values', [])

    start_row_2 = None
    end_row_2 = None

    # Find the start and end rows for the second deletion
    for i, row in enumerate(values_2):
        if 'Problem Description' in row:
            start_row_2 = i  # Start deleting from the row after "Problem Description"
        elif 'End of Reporting Record' in row:
            end_row_2 = i
            break

    if start_row_2 is not None and end_row_2 is not None:
        # Calculate the number of rows to delete
        num_rows_to_delete_2 = end_row_2 - start_row_2

        if num_rows_to_delete_2 > 0:
            # Use batchUpdate to delete rows
            requests_2 = [{
                'deleteDimension': {
                    'range': {
                        'sheetId': get_sheet_id(sheet_name),
                        'dimension': 'ROWS',
                        'startIndex': start_row_2+s_row,
                        'endIndex': end_row_2+s_row-1
                    }
                }
            }]

            body_2 = {
                'requests': requests_2
            }

            service.spreadsheets().batchUpdate(
                spreadsheetId=spread,
                body=body_2
            ).execute()
            time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)


def predict_poa_from_meter(df):
    meter = 'EGrid_KWH'
    poa = 'GlobInc_WHSQM'
    # Ensure there are no missing values in the columns
    df = df.dropna(subset=[meter, poa])
    
    # Reshape the data for sklearn
    X = df[meter].values.reshape(-1, 1)
    y = df[poa].values
    
    # Create and fit the model
    model = LinearRegression()
    model.fit(X, y)
    
    # Get the slope (coefficient) and intercept
    slope = model.coef_[0]
    intercept = model.intercept_

    # Function to predict POA based on meter value
    def predict_poa(meter_value):
        prediction = slope * meter_value + intercept
        if prediction > 2000:
            return 0
        else:
            return prediction 
    
    return predict_poa, slope, intercept





def insert_rows(sheet_name, start_row_index, num_rows, oVc, woVpa):
    if oVc == True:
        start_col = 6
    else:
        start_col = 4
    if woVpa == True:
        spreadch = woSheet
    else:
        spreadch = paSheet
    # Insert a new row at the specified index
    if num_rows > 0:
        sheet_id = get_sheet_id(sheet_name)
        end_row_index = start_row_index + num_rows

        requests = [{
            'insertDimension': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': start_row_index - 1,
                    'endIndex': end_row_index - 1
                },
                'inheritFromBefore': False
            }
        }]

        # Add formatting and merge requests for each new row
        for i in range(num_rows):
            current_row_index = start_row_index - 1 + i
            requests.append({
                'updateCells': {
                    'range': {'sheetId': sheet_id, 'startRowIndex': current_row_index, 'endRowIndex': current_row_index + 1, 'startColumnIndex': 0, 'endColumnIndex': 11},
                    'rows': [{'values': [{'userEnteredFormat': {'wrapStrategy': 'WRAP', 'horizontalAlignment': 'CENTER', 'verticalAlignment': 'MIDDLE'}}] * 11}],
                    'fields': 'userEnteredFormat(wrapStrategy,horizontalAlignment,verticalAlignment)'
                }
            })
            requests.append({
                'mergeCells': {
                    'range': {'sheetId': sheet_id, 'startRowIndex': current_row_index, 'endRowIndex': current_row_index + 1, 'startColumnIndex': start_col, 'endColumnIndex': 11},
                    'mergeType': 'MERGE_ALL'
                }
            })

        body = {'requests': requests}

        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadch,
            body=body
        ).execute()
        time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)





def input_data_to_Reports():
    global start_time

    # Define the font, fill, and border styles (not applicable in Google Sheets)
    # Format of groups['data'] = [df, sum_c_meter, total_kwh, sum_ghi, sum_poa, inv_availability, report_date, p50_kwh]
    for sites, groups in dfs.items():
        match = re.match(r'\b\w+\b', sites)
        sheet_name = match.group()
        if sheet_name in report_sheets:
            #Yearly Reset
            clear_cells(sheet_name)
            #Monthly WO Reset
            delete_rows(sheet_name, True)

            # Clear Old Data
            # Note: Clearing cells in Google Sheets is different and may require batch updates

            # Prepare data to be inserted
            reporting_row = 22 + groups['data'][6].month
            values = [
                [
                    round(groups['data'][1], 0), 
                    round(groups['data'][2], 0), 
                    None,
                    round(groups['data'][4], 0), 
                    None, 
                    None, 
                    None, 
                    None, 
                    None, 
                    groups['data'][5] if pd.notna(groups['data'][5]) else 0
                ]
            ]

            # Insert data into Google Sheets
            body = {
                'values': values
            }
            range_name = f"{sheet_name}!B{reporting_row}:K{reporting_row}"
            service.spreadsheets().values().update(
                spreadsheetId=paSheet,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
            time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time) (Rate 60 executes per minute)

            # Batch insert for closed_wo
            closed_wo_df = groups['closed_wo']
            if not closed_wo_df.empty:
                closed_wo_start_row = 39
                insert_rows(sheet_name, closed_wo_start_row, len(closed_wo_df), True, False)

                values = []
                for _, row in closed_wo_df.iterrows():
                    values.append([
                        value.strftime("%m/%d/%Y") if isinstance(value, (datetime.datetime, datetime.date)) and value is not pd.NaT else (value if pd.notna(value) else '')
                        for value in row
                    ])
                
                range_co = f"{sheet_name}!A{closed_wo_start_row}:K{closed_wo_start_row + len(closed_wo_df) - 1}"
                service.spreadsheets().values().update(
                    spreadsheetId=paSheet, range=range_co, valueInputOption='RAW', body={'values': values}
                ).execute()
                time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)

            # Batch insert for open_wo
            open_wo_df = groups['open_wo']
            if not open_wo_df.empty:
                open_wo_start_row = 39 + len(closed_wo_df) + 4 # Start after closed WOs and header
                insert_rows(sheet_name, open_wo_start_row, len(open_wo_df), False, False)

                values = []
                for _, row in open_wo_df.iterrows():
                    values.append([
                        value.strftime("%m/%d/%Y") if isinstance(value, (datetime.datetime, datetime.date)) and value is not pd.NaT else (value if pd.notna(value) else '')
                        for value in row
                    ])
                range_op = f"{sheet_name}!A{open_wo_start_row}:K{open_wo_start_row + len(open_wo_df) - 1}"
                service.spreadsheets().values().update(
                    spreadsheetId=paSheet, range=range_op, valueInputOption='RAW', body={'values': values}
                ).execute()
                time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)

            # Update outage count
            outage_count = sum(any(word in str(row[4]).lower() for word in ['outage', 'trip', 'curtail', 'curtailment']) for row in groups['closed_wo'].itertuples())
            outage_count += sum(any(word in str(row[4]).lower() for word in ['outage', 'trip', 'curtail', 'curtailment']) for row in groups['open_wo'].itertuples())
            if outage_count > 0:
                range_out = f"{sheet_name}!J18"
                service.spreadsheets().values().update(
                    spreadsheetId=paSheet,
                    range=range_out,
                    valueInputOption='RAW',
                    body={'values': [[outage_count]]}
                ).execute()


            #Quarterly Reports
            if datetime.datetime.now().month in [4, 7, 10, 1]:
                delete_rows(f"{sheet_name} Quarterly", True)
                # Batch insert for quarterly_closed_wo
                q_closed_wo_df = groups.get('quarterly_closed_wo')
                if q_closed_wo_df is not None and not q_closed_wo_df.empty:
                    qclosed_wo_start_row = 38
                    insert_rows(f"{sheet_name} Quarterly", qclosed_wo_start_row, len(q_closed_wo_df), True, False)

                    values = []
                    for _, row in q_closed_wo_df.iterrows():
                        values.append([
                            value.strftime("%m/%d/%Y") if isinstance(value, (datetime.datetime, datetime.date)) and value is not pd.NaT else (value if pd.notna(value) else '')
                            for value in row
                        ])
                    
                    range_co = f"{sheet_name} Quarterly!A{qclosed_wo_start_row}:K{qclosed_wo_start_row + len(q_closed_wo_df) - 1}"
                    service.spreadsheets().values().update(
                        spreadsheetId=paSheet, range=range_co, valueInputOption='RAW', body={'values': values}
                    ).execute()
                    time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)

                # Batch insert for open_wo in quarterly report
                q_open_wo_df = groups['open_wo']
                if not q_open_wo_df.empty:
                    qopen_wo_start_row = 38 + (len(q_closed_wo_df) if q_closed_wo_df is not None else 0) + 4
                    insert_rows(f"{sheet_name} Quarterly", qopen_wo_start_row, len(q_open_wo_df), False, False)

                    values = []
                    for _, row in q_open_wo_df.iterrows():
                        values.append([
                            value.strftime("%m/%d/%Y") if isinstance(value, (datetime.datetime, datetime.date)) and value is not pd.NaT else (value if pd.notna(value) else '')
                            for value in row
                        ])
                    range_op = f"{sheet_name} Quarterly!A{qopen_wo_start_row}:K{qopen_wo_start_row + len(q_open_wo_df) - 1}"
                    service.spreadsheets().values().update(
                        spreadsheetId=paSheet, range=range_op, valueInputOption='RAW', body={'values': values}
                    ).execute()
                    time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)

                # Update outage count
                outage_count = sum(any(word in str(row[4]).lower() for word in ['outage', 'trip', 'curtail', 'curtailment']) for row in groups['quarterly_closed_wo'].itertuples())
                outage_count += sum(any(word in str(row[4]).lower() for word in ['outage', 'trip', 'curtail', 'curtailment']) for row in groups['open_wo'].itertuples())
                if outage_count > 0:
                    range_out = f"{sheet_name} Quarterly!G25"
                    service.spreadsheets().values().update(
                        spreadsheetId=paSheet,
                        range=range_out,
                        valueInputOption='RAW',
                        body={'values': [[outage_count]]}
                    ).execute()


            print(f"Finished: {sheet_name:<20} |  Time: {round((time.time()-start_time)/60, 2):<4} Minutes")
    input_WO_only_reports()




def input_WO_only_reports():
    global start_time
    # Define the font, fill, and border styles (not applicable in Google Sheets)
    # Format of groups['data'] = [df, sum_c_meter, p50_kwh, sum_ghi, sum_poa, inv_availability, report_date]
    for sites, groups in dfs.items():        
        if sites in wo_Only_sites:
            print(f"    Site: {sites}")
            delete_rows(sites, False)

            # Batch insert for monthly_closed_wo
            monthly_closed_df = groups.get('monthly_closed_wo')
            if monthly_closed_df is not None and not monthly_closed_df.empty:
                closed_woonly_start_row = 10
                insert_rows(sites, closed_woonly_start_row, len(monthly_closed_df), True, True)
                
                values = []
                for _, row in monthly_closed_df.iterrows():
                    values.append([value.strftime("%m/%d/%Y") if isinstance(value, (datetime.datetime, datetime.date)) and value is not pd.NaT else (value if pd.notna(value) else '') for value in row])
                
                range_co = f"{sites}!A{closed_woonly_start_row}:K{closed_woonly_start_row + len(monthly_closed_df) - 1}"
                service.spreadsheets().values().update(
                    spreadsheetId=woSheet, range=range_co, valueInputOption='RAW', body={'values': values}
                ).execute()
                time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)

            # Batch insert for open_wo
            open_wo_df = groups.get('open_wo')
            if open_wo_df is not None and not open_wo_df.empty:
                open_woonly_start_row = 10 + (len(monthly_closed_df) if monthly_closed_df is not None else 0) + 4
                insert_rows(sites, open_woonly_start_row, len(open_wo_df), False, True)

                values = []
                for _, row in open_wo_df.iterrows():
                    values.append([value.strftime("%m/%d/%Y") if isinstance(value, (datetime.datetime, datetime.date)) and value is not pd.NaT else (value if pd.notna(value) else '') for value in row])

                range_op = f"{sites}!A{open_woonly_start_row}:K{open_woonly_start_row + len(open_wo_df) - 1}"
                service.spreadsheets().values().update(
                    spreadsheetId=woSheet, range=range_op, valueInputOption='RAW', body={'values': values}
                ).execute()
                time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)

            print(f"Finished: {sites+" Monthly ":<20} |  Time: {round((time.time()-start_time)/60, 2):<4} Minutes")

            #Quarterly Pages
            site_name = f'{sites} Quarterly'
            delete_rows(site_name, False)
            
            # Batch insert for quarterly_closed_wo
            quarterly_closed_df = groups.get('quarterly_closed_wo')
            if quarterly_closed_df is not None and not quarterly_closed_df.empty:
                closed_woQ_start_row = 10
                insert_rows(site_name, closed_woQ_start_row, len(quarterly_closed_df), True, True)
                values = [ [value.strftime("%m/%d/%Y") if isinstance(value, (datetime.datetime, datetime.date)) and value is not pd.NaT else (value if pd.notna(value) else '') for value in row] for _, row in quarterly_closed_df.iterrows() ]
                range_co = f"{site_name}!A{closed_woQ_start_row}:K{closed_woQ_start_row + len(quarterly_closed_df) - 1}"
                service.spreadsheets().values().update(
                    spreadsheetId=woSheet, range=range_co, valueInputOption='RAW', body={'values': values}
                ).execute()
                time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)

            # Batch insert for open_wo in quarterly
            if open_wo_df is not None and not open_wo_df.empty:
                open_woQ_start_row = 10 + (len(quarterly_closed_df) if quarterly_closed_df is not None else 0) + 4
                insert_rows(site_name, open_woQ_start_row, len(open_wo_df), False, True)
                values = [ [value.strftime("%m/%d/%Y") if isinstance(value, (datetime.datetime, datetime.date)) and value is not pd.NaT else (value if pd.notna(value) else '') for value in row] for _, row in open_wo_df.iterrows() ]
                range_op = f"{site_name}!A{open_woQ_start_row}:K{open_woQ_start_row + len(open_wo_df) - 1}"
                service.spreadsheets().values().update(
                    spreadsheetId=woSheet, range=range_op, valueInputOption='RAW', body={'values': values}
                ).execute()
                time.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)

            print(f"Finished: {sites:<20} |  Time: {round((time.time()-start_time)/60, 2):<4} Minutes")

            
    end_time = time.time()
    print(f"Time Taken: {round((end_time-start_time)/60, 2)} Minutes")
    messagebox.showinfo(title="Verify Pages", message="Check each Site sheet to ensure that the data is correct as well as input the Manual Entry items for this month.")
    root.destroy()

def update_close_out_date(row): #Puts the Completion Notes End Date in If no Completion date
    # Check if 'Completion Date' is empty and 'Close Out Date' is not
    if pd.isnull(row['Completion Date']) and pd.notnull(row.get('Close Out Date')):
        # Assign 'Close Out Date' to 'Completion Date'
        row['Completion Date'] = row['Close Out Date']
    elif pd.isnull(row['Completion Date']) and pd.notnull(row['Completion Notes']):
        match = re.search(r'(\d{1,2}[-/]\d{1,2})([-/](\d{2}|\d{4}))?', row['Completion Notes'])
        if match:
            month_day = match.group(1)
            year = match.group(3)
            if not year:
                current_year = datetime.datetime.now().year
                year = str(current_year)
            elif len(year) == 2:
                current_year = datetime.datetime.now().year
                century = current_year // 100 * 100
                year = str(century + int(year))

            date_str = f"{month_day}/{year}"
            
            try:
                return pd.to_datetime(date_str, format='%m/%d/%Y')
            except ValueError:
                return row['Completion Date']
    return row['Completion Date']

def parse_wo(file):
    global dfs
    try:
        wo = pd.read_excel(file)
    except UnicodeDecodeError:
        print("Error: The file encoding is not supported.")
        return
    # Get the current date and calculate the first and last day of the previous month
    today = datetime.datetime.today()
    first_day_of_current_month = today.replace(day=1)
    last_day_of_previous_month = first_day_of_current_month - datetime.timedelta(days=1)
    first_day_of_previous_month = last_day_of_previous_month.replace(day=1)
    someday_of_previous_q = first_day_of_previous_month - datetime.timedelta(days=40)
    first_day_of_previous_q = someday_of_previous_q.replace(day=1)


    for site, sheet in zip(wo_sites, xl_sheets):
        wo['Completion Date'] = pd.to_datetime(wo['Completion Date'], errors='coerce')
        wo['WO Date'] = pd.to_datetime(wo['WO Date'], errors='coerce')

        wo['Completion Date'] = wo.apply(update_close_out_date, axis=1)

        closed_wo = wo[
            (wo['Site'] == site) &
            (wo['Completion Date'] >= first_day_of_previous_month) & (wo['Completion Date'] <= last_day_of_previous_month)
        ]
        closed_woCopy = closed_wo.copy() 
        def parse_html(text):
            """Parse text with BeautifulSoup if it contains HTML tags."""
            if pd.notna(text) and '<' in text and '>' in text:
                return BeautifulSoup(text, "html.parser").get_text()
            return text

        # Apply parsing to both 'Completion Notes' and 'Work Description'
        closed_woCopy['Completion Notes'] = closed_woCopy['Completion Notes'].apply(parse_html)
        closed_woCopy['Work Description'] = closed_woCopy['Work Description'].apply(parse_html)

        # Create a new column 'Final Description' based on the condition
        closed_woCopy.loc[:,'Final Description'] = closed_woCopy.apply(
            lambda row: row['Completion Notes'] if pd.notna(row['Completion Notes']) and row['Completion Notes'] != '' else row['Work Description'], 
            axis=1
        )


        open_wo = wo[
            (wo['Site'] == site) &
            (pd.isna(wo['Completion Date'])) &
            (pd.isna(wo['Completion Notes'])) &
            (wo['WO Date'] <= last_day_of_previous_month)
        ]

        if last_day_of_previous_month.month in [3, 6, 9, 12]:
            q_closed_wo = wo[
                (wo['Site'] == site) &
                (wo['Completion Date'] >= first_day_of_previous_q) & (wo['Completion Date'] <= last_day_of_previous_month)
            ]   
            q_closed_woCopy = q_closed_wo.copy() 

            # Apply parsing to both 'Completion Notes' and 'Work Description'
            q_closed_woCopy['Completion Notes'] = q_closed_woCopy['Completion Notes'].apply(parse_html)
            q_closed_woCopy['Work Description'] = q_closed_woCopy['Work Description'].apply(parse_html)

            q_closed_woCopy.loc[:,'Final Description'] = q_closed_woCopy.apply(
                lambda row: row['Completion Notes'] if pd.notna(row['Completion Notes']) and row['Completion Notes'] != '' else row['Work Description'], 
                axis=1
            )
            q_closed_wo_selected = q_closed_woCopy[['WO No.', 'Asset Description', 'WO Date', 'Completion Date', 'Brief Description', 'Final Description']]
            dfs[sheet]['quarterly_closed_wo'] = q_closed_wo_selected
        
        closed_wo_selected = closed_woCopy[['WO No.', 'Asset Description', 'WO Date', 'Completion Date', 'WO Type', 'Brief Description', 'Final Description']]
        open_wo_selected = open_wo[['WO No.', 'Asset Description', 'WO Date', 'Sched. Completion Date', 'Brief Description']]

        # Save the DataFrames to the dictionary
        dfs[sheet]['closed_wo'] = closed_wo_selected
        dfs[sheet]['open_wo'] = open_wo_selected
    
    # Define a mapping for special cases
    site_aliases = {
        'Gray Fox': ['Gray fox', 'Gray Fox'],
        'Upson': ['Upson', 'Upson Ranch'],
        'Bulloch 1A': ['Bulloch 1A', 'Bulloch 1A & 1B'],
        'Bulloch 1B': ['Bulloch 1B', 'Bulloch 1A & 1B'],
    }

    #Monthly WO Only Reports
    for site in wo_Only_sites:
        if site in site_aliases.keys():
            possible_names = site_aliases[site]
        else:
            possible_names = [site]
        # Filter the DataFrame based on the conditions
        wo['Completion Date'] = pd.to_datetime(wo['Completion Date'], errors='coerce')
        wo['WO Date'] = pd.to_datetime(wo['WO Date'], errors='coerce')

        wo['Completion Date'] = wo.apply(update_close_out_date, axis=1)

        closed_wo = wo[
            (wo['Site'].isin(possible_names)) &
            (wo['Completion Date'] >= first_day_of_previous_month) & (wo['Completion Date'] <= last_day_of_previous_month)
        ]
        closed_woCopy = closed_wo.copy() 
                # Apply parsing to both 'Completion Notes' and 'Work Description'
        closed_woCopy['Completion Notes'] = closed_woCopy['Completion Notes'].apply(parse_html)
        closed_woCopy['Work Description'] = closed_woCopy['Work Description'].apply(parse_html)


        # Create a new column 'Final Description' based on the condition
        closed_woCopy.loc[:,'Final Description'] = closed_woCopy.apply(
            lambda row: row['Completion Notes'] if pd.notna(row['Completion Notes']) and row['Completion Notes'] != '' else row['Work Description'], 
            axis=1
        )


        open_wo = wo[
            (wo['Site'].isin(possible_names)) &
            (pd.isna(wo['Completion Date'])) &
            (pd.isna(wo['Completion Notes'])) &
            (wo['WO Date'] <= last_day_of_previous_month)
        ]

        closed_wo_selected = closed_woCopy[['WO No.', 'Asset Description', 'WO Date', 'Completion Date', 'WO Type', 'Brief Description', 'Final Description']]
        open_wo_selected = open_wo[['WO No.', 'Asset Description', 'WO Date', 'Sched. Completion Date', 'Brief Description']]

        # Save the DataFrames to the dictionary
        dfs[site]['monthly_closed_wo'] = closed_wo_selected
        dfs[site]['open_wo'] = open_wo_selected
    
    #Sets the Quarterly Timespan
    if first_day_of_previous_month.month > 2:
        first_day_three_months_ago = first_day_of_previous_month.replace(month=(first_day_of_previous_month.month - 2))
    else:
        first_day_three_months_ago = first_day_of_previous_month.replace(month=(first_day_of_previous_month.month+10), year=(today.year - 1))
        
    # Calculate the first day of the month three months ago
    

    # Quarterly WO Reports
    for site in wo_Only_sites:
        if site in site_aliases.keys():
            possible_names = site_aliases[site]
        else:
            possible_names = [site]

        # Filter the DataFrame based on the conditions
        wo['Completion Date'] = pd.to_datetime(wo['Completion Date'], errors='coerce')
        wo['WO Date'] = pd.to_datetime(wo['WO Date'], errors='coerce')

        wo['Completion Date'] = wo.apply(update_close_out_date, axis=1)

        closed_wo = wo[
            (wo['Site'].isin(possible_names)) &
            (wo['Completion Date'] >= first_day_three_months_ago) & 
            (wo['Completion Date'] <= last_day_of_previous_month)
        ]

        closed_woCopy = closed_wo.copy() 
        # Apply parsing to both 'Completion Notes' and 'Work Description'
        closed_woCopy['Completion Notes'] = closed_woCopy['Completion Notes'].apply(parse_html)
        closed_woCopy['Work Description'] = closed_woCopy['Work Description'].apply(parse_html)

        # Create a new column 'Final Description' based on the condition
        closed_woCopy.loc[:,'Final Description'] = closed_woCopy.apply(
            lambda row: row['Completion Notes'] if pd.notna(row['Completion Notes']) and row['Completion Notes'] != '' else row['Work Description'], 
            axis=1
        )
        closed_wo_selected = closed_woCopy[['WO No.', 'Asset Description', 'WO Date', 'Completion Date', 'WO Type', 'Brief Description', 'Final Description']]
        # Save the DataFrames to the dictionary
        dfs[site]['quarterly_closed_wo'] = closed_wo_selected

    input_data_to_Reports()
    #input_WO_only_reports()
  
def process_xl(file, wo_file):
    # Dictionary to hold DataFrames for each sheet
    global dfs
    
    
    # Loop through each sheet name in the list
    for sheet in xl_sheets:
        if sheet != 'Charlotte Airport':
            df = pd.read_excel(file, sheet_name=sheet, skiprows=2) # Read the sheet into a DataFrame
            dfd = df.drop(index=0) #Drop the Units Row
            dfd.dropna(subset=[dfd.columns[0]], inplace=True)
            inv_end_col = -3 
            dbcnxn()
            #different W/S check pioints for different sites
            #Need to Add moving GHI and POA columns data. 
            #lb and upb are lower and upper bounds for poa vs ghi check to make sure poa is tracking.
            if sheet == "Bluebird Solar":
                c_meter = 4
                c_poa = 3
                c_ghi = 2
                lb = 30000
                upb = 55000
                top = 4000
            elif sheet == "Cardinal":
                c_meter = 5
                c_poa = 3
                c_ghi = 2
                lb = 30000
                upb = 55000
                top = 8020
            elif sheet == "Cherry Blossom Solar, LLC":
                c_meter = 4
                c_poa = 2
                c_ghi = 3
                lb = 30000
                upb = 45000
                top = 12000
            elif sheet == "Cougar Solar, LLC Site":
                c_meter = 4
                c_poa = 3
                c_ghi = 2
                lb = -100
                upb = 5000
                top = 4000
            elif sheet == "Harrison Solar":
                c_meter = 5
                c_poa = 3
                c_ghi = 2
                lb = 5000
                upb = 45000
                top = 6000
            elif sheet == "Hayes":
                c_meter = 5
                c_poa = 3
                c_ghi = 4
                lb = 90000
                upb = 166000
                top = 4000
            elif sheet == "Hickory Solar, LLC":
                c_meter = 4
                c_poa = 2
                c_ghi = 3
                lb = 5000
                upb = 12000
                top = 6000
            elif sheet == "Violet Solar, LLC":
                c_meter = 4
                c_poa = 2
                c_ghi = 3
                lb = 30000
                upb = 55000
                top = 9000
            elif sheet == "Wellons Solar, LLC":
                c_meter = 5
                c_poa = 2
                c_ghi = 3
                lb = 25000
                upb = 50000
                top = 6000

            # Checks for missing POA data, which needs to be filled by user from PVsyst
            poa_col_name = dfd.columns[c_poa]
            # Check for gaps (NaN/null values) in the POA data
            if dfd[poa_col_name].isnull().any():
                # Find the DataFrame index of the first gap
                first_gap_index = dfd[dfd[poa_col_name].isnull()].index[0]

                # Calculate the corresponding row number in the Excel sheet.
                # read_excel(skiprows=2) skips Excel rows 1 and 2.
                # df.drop(index=0) removes the units row (originally Excel row 3).
                # So, dfd index 0 corresponds to Excel row 4.
                excel_row_number = first_gap_index + 4
                if messagebox.askokcancel(
                    title="POA Data Gap Detected",
                    message=f"Gaps detected in POA data for {sheet}.\n"
                            f"First gap found at Excel row: {excel_row_number}.\n\n"
                            f"Ok - Launches external tools for review and editing.\n"
                            "Cancel - Continues on with Performance Analysis Process"
                ) == True:
                    os.startfile(r"G:\Shared drives\O&M\NCC Automations\Performance Reporting\PVsyst (Josephs Edits).accdb")
                    os.startfile(file)
                    root.destroy()
                    sys.exit()


            dfd.iloc[1, 0] = pd.to_datetime(dfd.iloc[1, 0])
            report_date = dfd.iloc[1, 0]

            #Turn unrealistic values into Null values so that we can glean inv sum as meter
            dfd.iloc[:, c_meter] = dfd.iloc[:, c_meter].where(dfd.iloc[:, c_meter] <= top, np.nan)

            # Iterate through each cell in the c_meter column
            for index, cell in dfd.iloc[:, c_meter].items():
                if pd.isnull(cell):
                    if sheet == 'Wellons Solar, LLC': #This should be removed as soon as 1-2 starts communicating.
                        missing_inv_data = dfd.loc[index, 'Inverter 3-1 - SMA 800CP-US, Pac'] #Average of the Communicating Inverters
                        sum_inverters = missing_inv_data*4
                    else:
                        sum_inverters = dfd.iloc[index-1, c_meter+1:inv_end_col].sum()  # Sum the inverters in the same row
                    dfd.iloc[index-1, c_meter] = sum_inverters  # Place the sum in the blank cell
                    timestamp = dfd.iloc[index-1, 0]  # Assuming the first column contains the timestamps
                    

                    #maybe add an if statement for if sum_inverters = 0 don't send messagebox as it is likely a site outage. has been so far. 
                    if sum_inverters != 0:
                        messagebox.showinfo(
                            title="Performance Analysis: Meter Data Checks",
                            message=f"Found missing Meter Data in {sheet} at:\nTime: {timestamp}, Replaced Value: {sum_inverters}"
                        )

            sum_c_meter = dfd.iloc[:, c_meter].sum()
            sum_net_meter = dfd.iloc[:, -2].sum()
            sum_meter = max(sum_c_meter, sum_net_meter)

            # Define the SQL query
            site = sheet.upper().replace(", LLC", "").replace("SOLAR", "").replace("SITE", "").strip()
            query = "SELECT [GlobInc_WHSQM], [EGrid_KWH] FROM [PVsystHourly] WHERE [PlantName] = ?"

            # Execute the query and read into a DataFrame
            slope_df = pd.read_sql_query(query, connect_db, params=[site])

            if sheet == "Wellons Solar, LLC":
                #Converting from W to kW
                slope_df['EGrid_KWH'] = slope_df['EGrid_KWH'] / 1000

            connect_db.close()


            #The Following section is untested:
            # Predict POA from meter values
            predict_poa, slope, intercept = predict_poa_from_meter(slope_df)

            print("Slope Formula: ", site, f" | POA = {slope}x + {intercept}")

             # Ends at the Fourth column from last and includes it. This is only for the Filer, Select all inv to chekc for nan or 0 and removes the rows.
            filtered_df = dfd[~((dfd.iloc[:, c_meter+1:inv_end_col].lt(1) | dfd.iloc[:, c_meter+1:inv_end_col].isna()).all(axis=1))] # Filter the DataFrame to remove rows where inverters are only producing less than 1 kw for Inv Availability Calc.
            
            inv_availability = (filtered_df.iloc[:, inv_end_col].dropna().mean())/100
            p50_kwh, total_kwh = degradation_calc(report_date, sheet)
            if sheet == "Wellons Solar, LLC":
                #Adjusting Units W to kW
                p50_kwh = p50_kwh/1000
                total_kwh = total_kwh/1000



            # Filter the DataFrame to include only values greater than 0 for GHI and POA Calc
            df_filtered_C = dfd[dfd.iloc[:, c_ghi] > 0]  # Assuming column C is the 3rd column (index 2)
            df_filtered_D = dfd[dfd.iloc[:, c_poa] > 0]  # Assuming column D is the 4th column (index 3)
            sum_ghi = df_filtered_C.iloc[:, c_ghi].sum()
            sum_poa = df_filtered_D.iloc[:, c_poa].sum()

            #POA v. GHI Check. Making sure the POA sensor is tracking. 
            if not (lb <= (sum_poa - sum_ghi) <= upb):
                messagebox.showwarning(title="Performance Analysis: W/S Check", message=f"{sheet} POA is not within {lb}-{upb} more than the GHI data. Check for Issues. GHI: {sum_ghi} POA: {sum_poa} \n{sheet}")

            dfd['Predicted_POA'] = dfd.iloc[:, -2].apply(predict_poa)
            predicted_poa_sum = dfd['Predicted_POA'].sum()
            

            #Get POA col Name based on column Number
            poa_col_name = dfd.columns[c_poa]

            og_poa_data = dfd[poa_col_name].copy() # Use the column name for consistency

            # This fills NaN, 0, and Negative values in the target POA column with values from 'Predicted_POA'.
            dfd[poa_col_name] = dfd[poa_col_name].where(dfd[poa_col_name] > 0, dfd['Predicted_POA'])
            
            modified_poa_data = dfd[poa_col_name].copy() # Use the column name for consistency

            # Remove all values less than 0 in the c_poa column
            # Use .clip(lower=0) for an efficient way to set negative values to 0.
            dfd[poa_col_name] = dfd[poa_col_name].clip(lower=0)
            # Calculate the sum of the modified c_poa column
            modified_poa_sum = dfd[poa_col_name].sum()
            dates = dfd['Timestamp'].copy()
            show_poa_selection(predicted_poa_sum, sum_poa, modified_poa_sum, site, dates, dfd['Predicted_POA'], modified_poa_data, og_poa_data)
           
            #Save Data to Dict
            dfs[sheet]['data'] = ['dfd', sum_meter, total_kwh, sum_ghi, sum_poa_chosen, inv_availability, report_date, p50_kwh]   # Store the DataFrame in the dictionary
        else:
            continue
    parse_wo(wo_file)
        





def browse_files():
    global start_time

    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    AE_PA_file_path = filedialog.askopenfilename(initialdir=downloads_path, title="Select the Also Energy Performance Analysis File", filetypes=[("Excel files", "*.xlsx *.xls")])

    wo_File = filedialog.askopenfilename(initialdir=downloads_path, title="Select the Emaint WO Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])  
    start_time = time.time()
    
    #parse_wo(wo_File)
    process_xl(AE_PA_file_path, wo_File)




def degradation_calc(date, sheet):
    dbcnxn()
    # Strip "Solar" or "LLC" and change to all caps
    sitename = re.sub(r'\b(Solar|, LLC|Site)\b', '', sheet, flags=re.IGNORECASE).strip().upper()
    query = """
SELECT EGrid_KWH
FROM PVsystHourly
WHERE PlantName = ? AND MonthCode = ? AND EGrid_KWH > 0
    """


    c.execute(query, (sitename, date.month))
    results = c.fetchall()
    results = [float(row[0]) for row in results]
    total_kwh = sum(results)



    if isinstance(date, str):
        date = datetime.datetime.strptime(date, '%m/%d/%Y')


    if sitename == "BLUEBIRD":
        com_date_str = '11/1/2020'
        com_date = datetime.datetime.strptime(com_date_str, '%m/%d/%Y')
    elif sitename == "CARDINAL":
        com_date_str = '1/1/2021'
        com_date = datetime.datetime.strptime(com_date_str, '%m/%d/%Y')
    elif sitename == "CHERRY BLOSSOM":
        com_date_str = '1/1/2020'
        com_date = datetime.datetime.strptime(com_date_str, '%m/%d/%Y')
    elif sitename == "COUGAR":
        com_date_str = '5/1/2018'
        com_date = datetime.datetime.strptime(com_date_str, '%m/%d/%Y')
    elif sitename == "HARRISON":
        com_date_str = '1/1/2021'
        com_date = datetime.datetime.strptime(com_date_str, '%m/%d/%Y')
    elif sitename == "HAYES":
        com_date_str = '1/1/2021'
        com_date = datetime.datetime.strptime(com_date_str, '%m/%d/%Y')
    elif sitename == "HICKORY":
        com_date_str = '1/1/2020'
        com_date = datetime.datetime.strptime(com_date_str, '%m/%d/%Y')
    elif sitename == "VIOLET":
        com_date_str = '1/1/2020'
        com_date = datetime.datetime.strptime(com_date_str, '%m/%d/%Y')
    elif sitename == "WELLONS":
        com_date_str = '6/1/2016'
        com_date = datetime.datetime.strptime(com_date_str, '%m/%d/%Y')     


    difference_in_days = (date - com_date).days
    difference_in_years = difference_in_days / 365.25  # Using 365.25 to account for leap years
    degradation_percentage = difference_in_years * 0.005
    if sitename == 'WELLONS':
        print('Wellons p50: ', total_kwh, ' | ', degradation_percentage)
    p50_kwh = total_kwh * (1 - degradation_percentage)
    return (p50_kwh, total_kwh)




root = tk.Tk()
root.title("NARENCO Performance Analysis")
lbl = tk.Label(root, text="Please select the file from Also Energy 1st, then WO's")
plot_option_var = tk.BooleanVar()
plot_option = tk.Checkbutton(root, text='Select to view POA Charts', variable=plot_option_var)
browse = tk.Button(root, text="Browse Files", command= lambda: browse_files(), width= 50, height=5)

lbl.pack()
plot_option.pack()
browse.pack()



root.mainloop()