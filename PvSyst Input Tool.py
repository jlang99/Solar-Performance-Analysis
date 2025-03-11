import pyodbc
import re, csv
import pandas as pd
import tkinter as tk
from datetime import datetime, date, time, timedelta
from tkinter import filedialog, messagebox, simpledialog

def dbcnxn():
    global db, connect_db, c
    # Connect to DB
    db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\O&M\NCC Automations\Performance Reporting\PVsyst (Josephs Edits).accdb;'
    connect_db = pyodbc.connect(db)
    c = connect_db.cursor()
def extract_date(cell_value):
    # Use regex to extract the date part from the cell value
    match = re.match(r'(\d{2}/\d{2}/\d{2})', cell_value)
    match = re.match(r'(\d{2}/\d{2}/\d{2})', cell_value)
    if match:
        # Parse the extracted date
        date_obj = datetime.strptime(match.group(1), '%d/%m/%y')
        # Format the date as mm/dd/yyyy
        formatted_date = date_obj.strftime('%m/%d/%Y')
        return formatted_date
    return None

def import_csv_to_db(csv_file_path, table_name, plant_name):
    with open(csv_file_path, mode='r', encoding='iso-8859-1') as file:
        csv_reader = csv.reader(file, delimiter=';')
        for i, row in enumerate(csv_reader):
            if i == 6:
                print(row)
                simdate = row[2]

    #######
    #######
    #######
    #######
    #Needs to Match the DB, Date object needs to be organized for mm/dd/yyyy

    simulation_date = extract_date(simdate) 
    #######
    #######
    #######
    #######
    
    
    df = pd.read_csv(csv_file_path, skiprows=10, encoding='iso-8859-1', delimiter=';')
    # Remove the first two rows after the header
    df = df.iloc[2:]
    # Extract date components and create new columns
    df['DateTime'] = pd.to_datetime(df['date'], format='%d/%m/%y %H:%M')
    df['MonthCode'] = df['DateTime'].dt.month
    df['DayCode'] = df['DateTime'].dt.day
    df['HourCode'] = df['DateTime'].dt.hour
    df['DateTimeCode'] = df['DateTime'].dt.strftime('%m%d%H')

    # Define ID as PlantName + DateTimeCode
    df['ID'] = plant_name + '_' + df['DateTimeCode']




    # Map CSV columns to database columns
    csv_to_db_mapping = {
        'E_Grid': 'EGrid_KWH',
        'GlobInc': 'GlobInc_WHSQM'
    }

    # Filter the DataFrame to include only the desired columns
    df_filtered = df[list(csv_to_db_mapping.keys())]
    df_filtered = df_filtered.rename(columns=csv_to_db_mapping)
    
    df_filtered['ID'] = df['ID']
    df_filtered['PlantName'] = plant_name
    df_filtered['SimulationDate'] = simulation_date
    df_filtered['DateTimeCode'] = df['DateTimeCode']
    df_filtered['MonthCode'] = df['MonthCode']
    df_filtered['DayCode'] = df['DayCode']
    df_filtered['HourCode'] = df['HourCode']
    

    # Insert the DataFrame into the database
    for index, dfrow in df_filtered.iterrows():
        sql_insert = f"INSERT INTO {table_name} ({', '.join(df_filtered.columns)}) VALUES ({', '.join(['?'] * len(df_filtered.columns))})"
        try:
            c.execute(sql_insert, tuple(dfrow))
        except pyodbc.DataError as e:
            print(f"DataError: {e}")
            print(f"Row causing error: {dfrow}")
    connect_db.commit()

    messagebox.showinfo("Success", "Data imported successfully.")

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        table_name = 'PVsystHourly'  # Replace with your table name in the Access DB
        plant_name = simpledialog.askstring("Input", "Enter the Plant Name:")
        if plant_name:
            import_csv_to_db(file_path, table_name, plant_name)

def main():
    dbcnxn()
    root = tk.Tk()
    root.title("PvSyst Input Tool")
    root.geometry('300x100')
    select_button = tk.Button(root, text="Select CSV File", command=select_file, width = 50, height= 15)
    select_button.pack(pady=1)

    root.mainloop()

if __name__ == "__main__":
    main()