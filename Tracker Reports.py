import pandas as pd
from datetime import datetime, timedelta
import os, sys
import tkinter as tk
from tkinter import filedialog
import re
from bs4 import BeautifulSoup
import time as ty

#my Package
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(parent_dir)
from PythonTools import get_google_credentials, EMAILS, CREDS

# Google Sheets API imports
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Google Auth imports
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.auth import exceptions as auth_exceptions
from google_auth_oauthlib.flow import InstalledAppFlow

# For emailing and PDF handling
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from googleapiclient.http import MediaIoBaseUpload
import io
import requests


SHEET_ID = "1MlL1QKwOyOaNV9k0SJ59H0NsqvUVJfWXxocKxWhnmWQ"

# Define the necessary scopes for Sheets and Drive
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/drive']

# Site assignments
JOSEPH_SITES = {"Lily", "Bluebird", "Bulloch 1A", "Bulloch 1B", "CDIA", "Cougar", "Harding", "Harrison", "Holly Swamp", "JEFFERSON", "LongLeaf Pine Solar",
                "Marshall", "Mclean", "Sunflower", "Thunderhead", "Upson", "Violet", "Wayne II", "Wayne III", "Wellons Farm", "Whitehall", "Whitetail"}
JACOB_SITES = {"Lily", "Hayes", "Hickory", "BISHOPVILLE", "Cardinal", "Cherry Blossom", "Conetoe 1", "Duplin", "Elk",  "Freight Line", "Gray Fox",
               "HICKSON", "OGBURN", "PG Solar", "Richmond Cadle", "Shorthorn", "Tedder", "Van Buren", "Warbler", "Washington", "Wayne I", "WILLIAMS"}

# Placeholder folder IDs
JOSEPH_FOLDER_ID = '1KNf8yrqW58M4I_CuWrbP7LMn5cfQmKhA'
JACOB_FOLDER_ID = '17_JtiQRzJu-iPNx8DxT35WMT1S5Y1C1q'


def get_sheet_id(service, spreadsheet_id, sheet_name):
    """Helper function to get sheet ID."""
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in sheet_metadata.get('sheets', ''):
        if sheet.get("properties", {}).get("title") == sheet_name:
            return sheet.get("properties", {}).get("sheetId")
    return None

def find_or_create_folder(service, parent_folder_id, folder_name):
    """Finds a folder by name in a parent folder, or creates it if it doesn't exist."""
    try:
        # Search for the folder
        query = f"'{parent_folder_id}' in parents and name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        response = service.files().list(q=query,
                                        spaces='drive',
                                        fields='files(id, name)',
                                        supportsAllDrives=True,
                                        includeItemsFromAllDrives=True).execute()
        files = response.get('files', [])

        if files:
            # Folder found
            folder_id = files[0].get('id')
            print(f"Found folder '{folder_name}' with ID: {folder_id}")
            return folder_id
        else:
            # Folder not found, create it
            print(f"Folder '{folder_name}' not found. Creating it...")
            file_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_folder_id]
            }
            folder = service.files().create(body=file_metadata,
                                            fields='id',
                                            supportsAllDrives=True).execute()
            folder_id = folder.get('id')
            print(f"Created folder '{folder_name}' with ID: {folder_id}")
            return folder_id
    except HttpError as error:
        print(f'An error occurred while creating/finding the monthly folder: {error}')
        return None

def get_day_with_ordinal_suffix(d):
    """Returns the day of the month with its ordinal suffix (e.g., 1st, 2nd, 3rd, 4th)."""
    # Handles 11th, 12th, 13th
    if 11 <= d <= 13:
        return f"{d}th"
    suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(d % 10, 'th')
    return f"{d}{suffix}"

def email_pdf_reports():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except auth_exceptions.RefreshError as e:
                print(f"Error refreshing token: {e}")
                print("This is likely due to a change in scopes. Deleting token.json and re-authenticating.")
                os.remove('token.json')
                script_dir = os.path.dirname(os.path.abspath(__file__))
                credentials_path = os.path.join(script_dir, 'NCC-AutomationCredentials.json')
                flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
                creds = flow.run_local_server(port=0)
        else:
            # Make sure you have a 'credentials.json' file from Google Cloud in the same directory
            script_dir = os.path.dirname(os.path.abspath(__file__))
            credentials_path = os.path.join(script_dir, 'NCC-AutomationCredentials.json')
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    drive_service = build('drive', 'v3', credentials=creds)
    sheets_service = build('sheets', 'v4', credentials=creds)

    # Instead of deleting, create/find the monthly folder
    folder_name = datetime.now().strftime("%B %Y") # e.g., "January 2026"
    
    print(f"Ensuring subfolder '{folder_name}' exists for reports...")
    joseph_monthly_folder_id = find_or_create_folder(drive_service, JOSEPH_FOLDER_ID, folder_name)
    jacob_monthly_folder_id = find_or_create_folder(drive_service, JACOB_FOLDER_ID, folder_name)
    print("Subfolder check complete.")
    spreadsheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    all_sheets = spreadsheet_metadata.get('sheets', [])

    joseph_reports_generated = False
    jacob_reports_generated = False

    for sheet in all_sheets:
        sheet_properties = sheet.get('properties', {})
        sheet_title = sheet_properties.get('title')
        sheet_id = sheet_properties.get('sheetId')

        if sheet_title not in JOSEPH_SITES and sheet_title not in JACOB_SITES:
            continue

        print(f"Generating PDF for {sheet_title}...")

        try:
            # Use requests with auth headers to download the PDF content of a single sheet
            headers = {'Authorization': 'Bearer ' + creds.token}
            url = f'https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=pdf&gid={sheet_id}'
            res = requests.get(url, headers=headers)
            res.raise_for_status()  # Raise an exception for bad status codes
            pdf_content = res.content

            # Determine the correct folder
            folder_id = None
            if sheet_title in JOSEPH_SITES:
                folder_id = joseph_monthly_folder_id
            elif sheet_title in JACOB_SITES:
                folder_id = jacob_monthly_folder_id

            if not folder_id:
                print(f"No valid monthly folder ID found for {sheet_title}. Skipping upload.")
                continue

            # Add day to filename
            day_of_month = datetime.now().day
            day_str = get_day_with_ordinal_suffix(day_of_month)
            file_name = f'{sheet_title} Tracker Report {day_str}.pdf'

            # Upload the PDF to Google Drive
            file_metadata = {
                'name': file_name,
                'mimeType': 'application/pdf',
                'parents': [folder_id]
            }
            media = MediaIoBaseUpload(io.BytesIO(pdf_content), mimetype='application/pdf', resumable=True)
            
            created_file = drive_service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id, webViewLink',
                supportsAllDrives=True
            ).execute()
            
            file_link = created_file.get('webViewLink')
            print(f"Uploaded {sheet_title} report. Link: {file_link}")

            # Add link to the appropriate list
            if sheet_title in JOSEPH_SITES:
                joseph_reports_generated = True
            if sheet_title in JACOB_SITES:
                jacob_reports_generated = True

            # Add a delay to avoid hitting API rate limits
            ty.sleep(5)
        except requests.exceptions.RequestException as e:
            print(f"Error downloading PDF for {sheet_title}: {e}")
        except HttpError as e:
            print(f"Error uploading PDF for {sheet_title} to Google Drive: {e}")

    # Send the email with the links
    if not joseph_reports_generated and not jacob_reports_generated:
        print("No reports were generated. Skipping email.")
        return
    
    #return #pause email send
    sender_email = EMAILS['NCC Desk']
    recipients = EMAILS['Administrators Only']
    smtp_password = CREDS['shiftsumEmail'] # Assuming this is the correct password key
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = f"Weekly Tracker Reports - {datetime.now().strftime('%Y-%m-%d')}"

    html_body = "<html><body><p>Hello Team,</p><p>Please find the Weekly tracker reports linked below:</p>"

    if joseph_reports_generated:
        joseph_folder_link = f"https://drive.google.com/drive/folders/{JOSEPH_FOLDER_ID}"
        html_body += f'<h3><a href="{joseph_folder_link}">Joseph\'s Sites Reports</a></h3>'
    
    if jacob_reports_generated:
        jacob_folder_link = f"https://drive.google.com/drive/folders/{JACOB_FOLDER_ID}"
        html_body += f'<h3><a href="{jacob_folder_link}">Jacob\'s Sites Reports</a></h3>'

    html_body += "<p>Thank you,<br>NCC Automation</p></body></html>"
    msg.attach(MIMEText(html_body, 'html'))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, smtp_password)
            server.send_message(msg)
        print(f"Email sent successfully to {', '.join(recipients)}!")
    except Exception as e:
        print(f"Failed to send email: {e}")

def delete_rows(sheet_name):
    creds = get_google_credentials()
    service = build('sheets', 'v4', credentials=creds)

    begin_row = 12
    s_row = 16
    # Define the range to search for the start and end markers
    search_range = f"{sheet_name}!G{begin_row}:K"

    # Read the data from the specified range
    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=search_range
    ).execute()
    ty.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)
    values = result.get('values', [])

    sheet_id = get_sheet_id(service, SHEET_ID, sheet_name)
    if sheet_id is None:
        print(f"Could not find sheetId for {sheet_name}")
        return

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
                        'sheetId': sheet_id,
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
                spreadsheetId=SHEET_ID,
                body=body
            ).execute()
            ty.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)




    # Define the range to search for the start and end markers for the second deletion
    search_range_2 = f"{sheet_name}!F{s_row}:K"

    # Read the data from the specified range
    result_2 = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=search_range_2
    ).execute()
    ty.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)


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
                        'sheetId': sheet_id,
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
                spreadsheetId=SHEET_ID,
                body=body_2
            ).execute()
            ty.sleep(0.7) #Slow the executes to prevent being kicked by Google (playing with the minimum time)

    print(f"Deleted {num_rows_to_delete} rows at {start_row+begin_row} to {end_row+begin_row-1} from {sheet_name}.")    
    print(f"Deleted {num_rows_to_delete_2} rows at {start_row_2+s_row} to {end_row_2+s_row-1} from {sheet_name} (second deletion).")


def insert_rows(service, spreadsheet_id, sheet_id, start_row_index, num_rows, openVclosed):
    """Inserts rows into a sheet."""
    if num_rows <= 0:
        return

    if openVclosed:
        start_col = 6  # Column G
    else:
        start_col = 5  # Column F
    requests = [{
        'insertDimension': {
            'range': {
                'sheetId': sheet_id,
                'dimension': 'ROWS',
                'startIndex': start_row_index - 1, # 0-indexed
                'endIndex': start_row_index - 1 + num_rows
            },
            'inheritFromBefore': False 
        }
    }]

    # Add merge requests for each new row
    for i in range(num_rows):
        current_row_index = start_row_index - 1 + i
        requests.append({
            'updateCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': current_row_index,
                    'endRowIndex': current_row_index + 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': 11
                },
                'rows': [{'values': [{'userEnteredFormat': {
                    'borders': {
                        'top': {'style': 'SOLID'},
                        'bottom': {'style': 'SOLID'},
                        'left': {'style': 'SOLID'},
                        'right': {'style': 'SOLID'}
                    },
                    'wrapStrategy': 'WRAP',
                    'horizontalAlignment': 'CENTER',
                    'verticalAlignment': 'MIDDLE'
                }}] * 11}],
                'fields': 'userEnteredFormat(borders,wrapStrategy,horizontalAlignment,verticalAlignment)'
            }
        })
        requests.append({
            'mergeCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': current_row_index,
                    'endRowIndex': current_row_index + 1,
                    'startColumnIndex': start_col,
                    'endColumnIndex': 11  # Column K is index 10, endIndex is exclusive
                },
                'mergeType': 'MERGE_ALL'
            }
        })
    body = {'requests': requests}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    ty.sleep(0.7)


def input_to_spreadsheet(site, issue_log_time, log_repair_time, issue_repair_time, average_open, longest_open, longest_open_wo, fastest_repair, fastest_repair_wo, open_wos, completed_wos, stow_count):
    print(f"Updating Sheet: {site}")
    
    open_wo_count = len(open_wos)
    creds = get_google_credentials()
    service = build('sheets', 'v4', credentials=creds)
    sheet_service = service.spreadsheets()

    spreadsheet_metadata = sheet_service.get(spreadsheetId=SHEET_ID).execute()
    sheet_id = None
    for s in spreadsheet_metadata.get('sheets', []):
        if s.get('properties', {}).get('title') == site:
            sheet_id = s.get('properties', {}).get('sheetId')
            break
    
    if sheet_id is None:
        print(f"Sheet '{site}' not found. Skipping update.")
        return

    def format_timedelta(td):
        if pd.isna(td) or td is None:
            return "N/A"
        days = td.days
        hours, remainder = divmod(td.seconds, 3600)
        minutes, _ = divmod(remainder, 60)
        return f"{days} days, {hours:02}:{minutes:02}"

    # Convert numpy types to standard Python types to avoid JSON serialization errors
    longest_open_wo_val = int(longest_open_wo) if pd.notna(longest_open_wo) else "N/A"
    fastest_repair_wo_val = int(fastest_repair_wo) if pd.notna(fastest_repair_wo) else "N/A"

    metrics_data = [
        [
            f"{issue_log_time:.0f} days" if issue_log_time is not None else "N/A",
            f"{log_repair_time:.0f} days" if log_repair_time is not None else "N/A",
            "", 
            format_timedelta(issue_repair_time),
            int(open_wo_count),
            format_timedelta(average_open),
            format_timedelta(longest_open),
            longest_open_wo_val,
            format_timedelta(fastest_repair),
            fastest_repair_wo_val,
            int(stow_count)
        ]
    ]
    
    sheet_service.values().update(spreadsheetId=SHEET_ID, range=f"'{site}'!A9:K9", valueInputOption='USER_ENTERED', body={'values': metrics_data}).execute()
    print(f"Updated metrics for {site}\nDeleting old data...")
    delete_rows(site)


    # Write new "Completed Work Orders" data.
    if not completed_wos.empty:
        completed_wos_to_sheet = completed_wos[['WO No.', 'Start Date Time', 'WO Date', 'End Date Time', 'Duration', 'Brief Description', 'Work Description']].copy()
        completed_wos_to_sheet['WO Date'] = pd.to_datetime(completed_wos_to_sheet['WO Date']).dt.strftime('%Y-%m-%d')
        completed_wos_to_sheet['Start Date Time'] = pd.to_datetime(completed_wos_to_sheet['Start Date Time']).dt.strftime('%Y-%m-%d %H:%M')
        completed_wos_to_sheet['End Date Time'] = pd.to_datetime(completed_wos_to_sheet['End Date Time']).dt.strftime('%Y-%m-%d %H:%M')
        completed_wos_to_sheet['Duration'] = completed_wos_to_sheet['Duration'].apply(format_timedelta)
        
        values = completed_wos_to_sheet.fillna('N/A').values.tolist()
        insert_rows(service, SHEET_ID, sheet_id, 13, len(values), True)
        sheet_service.values().update(spreadsheetId=SHEET_ID, range=f"'{site}'!A13", valueInputOption='USER_ENTERED', body={'values': values}).execute()

    # Determine the dynamic starting row for the "Open Work Orders" section.
    if not completed_wos.empty:
        num_completed_rows = len(completed_wos.index)
        # Layout: Header at 13, then N data rows. Then EndMarker, Blank, Title, Header.
        open_wo_header_row = 13 + num_completed_rows + 4
    else:
        # If no completed WOs, user specified data starts at 17, so header is at 16.
        open_wo_header_row = 17
    
    # Write new "Open Work Orders" data.
    if not open_wos.empty:
        open_wos_to_sheet = open_wos[['WO No.', 'Start Date Time', 'WO Date', 'Sched. Completion Date', 'Duration', 'Brief Description']].copy()
        open_wos_to_sheet['WO Date'] = pd.to_datetime(open_wos_to_sheet['WO Date']).dt.strftime('%Y-%m-%d')
        open_wos_to_sheet['Start Date Time'] = pd.to_datetime(open_wos_to_sheet['Start Date Time']).dt.strftime('%Y-%m-%d %H:%M')
        open_wos_to_sheet['Sched. Completion Date'] = pd.to_datetime(open_wos_to_sheet['Sched. Completion Date']).dt.strftime('%Y-%m-%d')
        open_wos_to_sheet['Duration'] = open_wos_to_sheet['Duration'].apply(format_timedelta)

        values = open_wos_to_sheet.fillna('N/A').values.tolist()
        insert_rows(service, SHEET_ID, sheet_id, open_wo_header_row, len(values), False)
        sheet_service.values().update(spreadsheetId=SHEET_ID, range=f"'{site}'!A{open_wo_header_row}", valueInputOption='USER_ENTERED', body={'values': values}).execute()

    print(f"Successfully updated sheet for {site}")



def process_wos(file_path):
    creds = get_google_credentials()
    service = build('sheets', 'v4', credentials=creds)
    spreadsheet_metadata = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    valid_sites = {s.get('properties', {}).get('title') for s in spreadsheet_metadata.get('sheets', [])}

    df = pd.read_excel(file_path, sheet_name='Sheet1')

    for site, site_df in df.groupby('Site'):
        if site not in valid_sites:
            continue
        print(f"Processing site: {site}")
        
        # Count WOs with 'stow' in the brief description, case-insensitive
        stow_count = site_df['Brief Description'].str.lower().str.contains('stow', na=False).sum()

        completed_wos = site_df[site_df['Job Status'].isin(['Complete', 'Closed'])].copy()
        open_wos = site_df[~site_df['Job Status'].isin(['Complete', 'Closed'])].copy()

        durations = []
        log_delays = []
        repair_delays = []
        start_datetimes = []
        end_datetimes = []
        completed_clean_descriptions = []

        def fix_time(t_match):
            if not t_match: return None
            t = t_match.group(1)
            if ':' in t: return t
            if len(t) == 3: return f"0{t[0]}:{t[1:]}"
            if len(t) == 4: return f"{t[:2]}:{t[2:]}"
            return t
        
        def fix_date(d_match):
            if not d_match: return None
            d = d_match.group(1)
            if re.search(r'\d{4}', d): return d
            if re.search(r'[-/]\d{2}$', d): 
                 parts = re.split(r'[-/]', d)
                 return f"{parts[0]}/{parts[1]}/20{parts[2]}"
            return f"{d}/{datetime.now().year}"

        for index, row in completed_wos.iterrows():
            desc = row['Work Description']
            s_dt, e_dt = None, None
            clean_desc = desc
            if pd.isna(desc):
                start_datetimes.append(s_dt)
                end_datetimes.append(e_dt)
                durations.append(None)
                log_delays.append(None)
                repair_delays.append(None)
                completed_clean_descriptions.append(None)
                continue
            
            try:
                clean_text = BeautifulSoup(str(desc), "html.parser").get_text()
                clean_desc = clean_text
                
                start_date = re.search(r'Start Date:\s*(\d{1,2}[-/.]\d{1,2}(?:[-/.]\d{2,4})?)', clean_text, re.IGNORECASE)
                start_time = re.search(r'Start Time:\s*(\d{1,2}[:;]?\d{1,2}|\d{3,4})', clean_text, re.IGNORECASE)
                end_date = re.search(r'End Date:\s*(\d{1,2}[-/.]\d{1,2}(?:[-/.]\d{2,4})?)', clean_text, re.IGNORECASE)
                end_time = re.search(r'End Time:\s*(\d{1,2}[:;]?\d{1,2}|\d{3,4})', clean_text, re.IGNORECASE)

                s_date_str = fix_date(start_date)
                s_time_str = fix_time(start_time)
                e_date_str = fix_date(end_date)
                e_time_str = fix_time(end_time)

                if s_date_str and s_time_str:
                    s_dt = pd.to_datetime(f"{s_date_str} {s_time_str}", errors='coerce')
                if e_date_str and e_time_str:
                    e_dt = pd.to_datetime(f"{e_date_str} {e_time_str}", errors='coerce')

            except Exception:
                s_dt, e_dt = None, None
            
            completed_clean_descriptions.append(clean_desc)
            start_datetimes.append(s_dt)
            end_datetimes.append(e_dt)

            if pd.notna(s_dt) and pd.notna(e_dt):
                durations.append(e_dt - s_dt)
                wo_date = pd.to_datetime(row['WO Date'], errors='coerce')
                if pd.notna(wo_date):
                    log_delays.append((wo_date.date() - s_dt.date()).days)
                    repair_delays.append((e_dt.date() - wo_date.date()).days)
                else:
                    log_delays.append(None)
                    repair_delays.append(None)
            else:
                durations.append(None)
                log_delays.append(None)
                repair_delays.append(None)

        completed_wos['Work Description'] = completed_clean_descriptions
        completed_wos['Duration'] = durations
        completed_wos['Log Delay (Days)'] = log_delays
        completed_wos['Repair Delay (Days)'] = repair_delays
        completed_wos['Start Date Time'] = start_datetimes
        completed_wos['End Date Time'] = end_datetimes

        open_start_datetimes = []
        open_durations = []
        open_clean_descriptions = []
        now = datetime.now()
        for index, row in open_wos.iterrows():
            desc = row['Work Description']
            s_dt = None
            clean_desc = desc
            if pd.notna(desc):
                try:
                    clean_text = BeautifulSoup(str(desc), "html.parser").get_text()
                    clean_desc = clean_text
                    start_date_match = re.search(r'Start Date:\s*(\d{1,2}[-/]\d{1,2}(?:[-/]\d{2,4})?)', clean_text, re.IGNORECASE)
                    start_time_match = re.search(r'Start Time:\s*(\d{1,2}[:;]?\d{1,2}|\d{3,4})', clean_text, re.IGNORECASE)
                    
                    s_date_str = fix_date(start_date_match)
                    s_time_str = fix_time(start_time_match)

                    if s_date_str and s_time_str:
                        s_dt = pd.to_datetime(f"{s_date_str} {s_time_str}", errors='coerce')
                except Exception:
                    s_dt = None
            open_clean_descriptions.append(clean_desc)
            open_start_datetimes.append(s_dt)
            if pd.notna(s_dt):
                open_durations.append(now - s_dt)
            else:
                open_durations.append(None)
        
        open_wos['Work Description'] = open_clean_descriptions
        open_wos['Start Date Time'] = open_start_datetimes
        open_wos['Duration'] = open_durations
        
        avg_log_delay = None
        avg_repair_delay = None
        avg_duration = None
        max_duration = None
        max_duration_wo = None
        min_duration = None
        min_duration_wo = None

        if not completed_wos['Duration'].dropna().empty:
            max_duration_idx = completed_wos['Duration'].idxmax()
            min_duration_idx = completed_wos['Duration'].idxmin()
            max_duration_wo = completed_wos.loc[max_duration_idx, 'WO No.']
            min_duration_wo = completed_wos.loc[min_duration_idx, 'WO No.']
            avg_duration = completed_wos['Duration'].mean()
            max_duration = completed_wos['Duration'].max()
            min_duration = completed_wos['Duration'].min()
        
        if not completed_wos['Log Delay (Days)'].dropna().empty:
            avg_log_delay = completed_wos['Log Delay (Days)'].mean()

        if not completed_wos['Repair Delay (Days)'].dropna().empty:
            avg_repair_delay = completed_wos['Repair Delay (Days)'].mean()
        
        if not open_wos['Duration'].dropna().empty:
            avg_open_duration = open_wos['Duration'].mean()

        input_to_spreadsheet(
            site=site,
            issue_log_time=avg_log_delay,    
            log_repair_time=avg_repair_delay,
            issue_repair_time=avg_duration,
            average_open=avg_open_duration,
            longest_open=max_duration,
            longest_open_wo=max_duration_wo,
            fastest_repair=min_duration,
            fastest_repair_wo=min_duration_wo,
            open_wos=open_wos,
            completed_wos=completed_wos,
            stow_count=stow_count
        )
    if datetime.now().weekday() < 2: #Monday and Tuesday
        email_pdf_reports()
    #Close the Tkinter window after processing
    root.destroy()



def browse_file():
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    file_path = filedialog.askopenfilename(title="Select Tracker WO'S Excel File", filetypes=[("Excel files", "*.xlsx")], initialdir=downloads_folder)
    if file_path:
        process_wos(file_path)









#This is where the Script Starts. Predefined Shit is Above.
# Create the main window
root = tk.Tk()
root.title("Tracker Reports")
try:
    root.iconbitmap(r"G:\Shared drives\O&M\NCC Automations\Icons\tracker_3KU_icon.ico")
except Exception as e:
    print(f"Error loading icon: {e}")

browse_button = tk.Button(root, text="Select Tracker WO file", command=browse_file, width= 50, height= 10)
browse_button.pack(fill='both')
frame = tk.Frame(root)
frame.pack()

root.mainloop()
