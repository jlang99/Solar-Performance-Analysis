import pyodbc, datetime, os, time, re, warnings, openpyxl
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

from icecream import ic


met_station_file = r"G:\Shared drives\O&M\NCC Automations\In Progress\Performance Analysis - Python\Brandon\Python Switch\Lily\Lily Report Data.xlsx"

teletrans_dfs = {}
meteo_dfs = {}

data_conversion = {
    'NumEstacion': {
        3879: 'WEATHER STATION 1',
        3880: 'WEATHER STATION 2',
        3881: 'WEATHER STATION 3',
        3882: 'WEATHER STATION 4'
    },
    'NumParametro': {
        1: 'Battery Level',
        2: 'Relative Humidity',
        4: 'Rain',
        6: 'Air Temperature',
        7: 'Wind Speed',
        8: 'Wind Direction',
        308: 'Surface Temperature 1',
        309: 'Surface Temperature 2',
        537: 'External Voltage Reference',
        590: 'External Voltage Reference 2',
        591: 'Cell Radiation 1',
        592: 'Cell Radiation 2',
        851: 'Global Radiation 1',
        853: 'Global Radiation 2'
    },
    'NumFuncion': {
        1: 'Average',
        2: 'Accumulation',
        4: 'Maximum',
        5: 'Minimum',
        6: 'Standard Deviation'
    }
}





def dbcnxn():
    global db, connect_db, c
    #Connect to DB
    db = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=G:\Shared drives\O&M\NCC Automations\In Progress\Performance Analysis - Python\Brandon\Python Switch\Also Energy Sites\PVsyst (Josephs Edits).accdb;'
    connect_db = pyodbc.connect(db)
    c = connect_db.cursor()

def save_data_excel(lilyscada_df, production_df, soiling_df, invall_df):
    # Load the Excel file and clear the contents of the sheet except for the first row
    wb = openpyxl.load_workbook(met_station_file)
    for ws in wb:
        if ws != 'SCADA DATA':
            for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=ws.max_column):
                for cell in row:
                    cell.value = None
        else:
            for row in ws.iter_rows(min_row=ws.min_row, max_row=ws.max_row, max_col=ws.max_column):
                for cell in row:
                    cell.value = None
    wb.save(met_station_file)
    #IRRADIANCE TELETRANS Page        
    for station, station_df in teletrans_dfs.items(): 
        # Write the DataFrame to the Excel sheet
        with pd.ExcelWriter(met_station_file, mode='a', if_sheet_exists='overlay') as writer:
            # Debugging: Print the DataFrame to check its content
            print(f"Irradiance Data for {station}:")
            print(station_df.head())
            station_df.to_excel(writer, sheet_name=station, index=False, startrow=2, startcol=8)
    #METEO TELETRANS Page        
    for station, station_df in meteo_dfs.items(): 
        # Write the DataFrame to the Excel sheet
        with pd.ExcelWriter(met_station_file, mode='a', if_sheet_exists='overlay') as writer:
            # Debugging: Print the DataFrame to check its content
            print(f"Meteo Data for {station}:")
            print(station_df.head())
            station_df.to_excel(writer, sheet_name=station, index=False, startrow=2)
    #SCADA DATA Page        
    with pd.ExcelWriter(met_station_file, mode='a', if_sheet_exists='overlay') as writer:
        # Debugging: Print the DataFrame to check its content
        print(f"SCADA DATA")
        print(lilyscada_df.head())
        lilyscada_df.to_excel(writer, sheet_name='SCADA DATA', index=False)
    #Production      
    with pd.ExcelWriter(met_station_file, mode='a', if_sheet_exists='overlay') as writer:
        # Debugging: Print the DataFrame to check its content
        print(f"PRODUCTION")
        print(production_df.head())
        production_df.to_excel(writer, sheet_name='PRODUCTION', index=False, startrow=3)
    #Soiling    
    with pd.ExcelWriter(met_station_file, mode='a', if_sheet_exists='overlay') as writer:
        # Debugging: Print the DataFrame to check its content
        print(f"SOILING")
        print(soiling_df.head())
        soiling_df.to_excel(writer, sheet_name='SOILING', index=False)
    #INVALL    
    with pd.ExcelWriter(met_station_file, mode='a', if_sheet_exists='overlay') as writer:
        # Debugging: Print the DataFrame to check its content
        print(f"INVALL")
        print(invall_df.head())
        invall_df.to_excel(writer, sheet_name='INVALL', index=False)


    end = time.time()
    print(round(end-start_time, 2))
    os.startfile(met_station_file)
    




def process_lily_scada_data(callcellscsv, inv3csv, invallcsv, ppccsv, subcsv, submfmtcsv, production):
    callcelldf = pd.read_csv(callcellscsv, header=1, delimiter=';')
    inv3df = pd.read_csv(inv3csv, header=1, delimiter=';')
    invalldf = pd.read_csv(invallcsv, header=1, delimiter=';')
    ppcdf = pd.read_csv(ppccsv, header=1, delimiter=';')
    subdf = pd.read_csv(subcsv, header=1, delimiter=';')
    submfmtdf = pd.read_csv(submfmtcsv, header=1, delimiter=';')
    productiondf = pd.read_csv(production, header=1)

    #Production Sum Values
    proddf = productiondf.iloc[:, 1:3]
    # Function to clean and convert values to float\
    print('Pre Cleaning:', proddf.head())

    def clean_and_convert(value):
        if isinstance(value, str):
            # Remove leading apostrophes
            val = value.lstrip("'")
            return float(val)
        else:
            return value
        
    proddf = proddf.map(clean_and_convert)
    print('Post Cleaning:', proddf.head())
    
    col_sum = proddf.sum()
    # Create a new DataFrame with the sum
    sum_df = pd.DataFrame([col_sum], columns=['KWH Delivered', 'KWH Received'])
    print('Final:', sum_df.head())

    #Soiling Data
    # Select specific columns by their names
    soiling_columns = callcelldf[['Date/time', 'US LIL CALIBRATED CELL 1.02 RADIATION CELL 1 (W/m2)', 
                       'US LIL CALIBRATED CELL 1.03 RADIATION CELL 2 (W/m2)', 'US LIL CALIBRATED CELL 2.02 RADIATION CELL 1 (W/m2)', 
                       'US LIL CALIBRATED CELL 2.03 RADIATION CELL 2 (W/m2)', 'US LIL CALIBRATED CELL 3.02 RADIATION CELL 1 (W/m2)', 
                       'US LIL CALIBRATED CELL 3.03 RADIATION CELL 2 (W/m2)', 'US LIL CALIBRATED CELL 4.02 RADIATION CELL 1 (W/m2)', 
                       'US LIL CALIBRATED CELL 4.03 RADIATION CELL 2 (W/m2)', 'US LIL CALIBRATED CELL 5.02 RADIATION CELL 1 (W/m2)', 
                       'US LIL CALIBRATED CELL 5.03 RADIATION CELL 2 (W/m2)', 'US LIL CALIBRATED CELL 6.02 RADIATION CELL 1 (W/m2)', 
                       'US LIL CALIBRATED CELL 6.03 RADIATION CELL 2 (W/m2)']]



    # Select the first four columns from ppcdf
    ppc_selected = ppcdf.iloc[:, :4]
    # Select columns 2, 4, and 3 from subdf (Note: iloc is zero-indexed)
    sub_selected = subdf.iloc[:, [1, 3, 2]]
    submfmt_selected = submfmtdf.iloc[:, [1, 3, 2, 4]]
    inv3_selected = inv3df.iloc[:, [1, 2]]

    # Concatenate the selected columns horizontally
    lilyscada_df = pd.concat([ppc_selected, sub_selected, submfmt_selected, inv3_selected], axis=1)

    save_data_excel(lilyscada_df, proddf, soiling_columns, invalldf)

def process_met_stations(datos_file):
    # Read the CSV file into a DataFrame, setting the first row as the header
    df = pd.read_csv(datos_file, header=0)
    
    # Convert data in the DataFrame using the data_conversion dictionary
    for column, conversion_dict in data_conversion.items():
        if column in df.columns:
            df[column] = df[column].map(conversion_dict).fillna(df[column])
    
    # Convert 'Fecha' column to datetime and format it to mm/dd/yyyy 00:00:00
    if 'Fecha' in df.columns:
        df['Fecha'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y %H:%M:%S').dt.strftime('%m/%d/%Y %H:%M:%S')
    
    # Drop unnecessary columns
    df.drop(columns=['Control', 'Tipo', 'Calidad1', 'Calidad2'], inplace=True)
    
    # Create teletrans_df for each weather station
    weather_stations = df['NumEstacion'].unique()
    
    
    for station in weather_stations:
        station_df = df[(df['NumEstacion'] == station) & (df['NumFuncion'] == 'Average')].copy()
        station_df.rename(columns={'Fecha': 'DateTime'}, inplace=True)


        # Filter and rename columns based on NumParametro
        parameters = ['Cell Radiation 1', 'Cell Radiation 2', 'Global Radiation 1', 'Global Radiation 2']
        for param in parameters:
            filtered_df = station_df[station_df['NumParametro'] == param][['DateTime', 'Valor']].copy()
            filtered_df.rename(columns={'Valor': param}, inplace=True)
            if 'DateTime' in station_df.columns:
                station_df = station_df.merge(filtered_df, on='DateTime', how='left')
        
        # Select only the 'DateTime' and new parameter columns
        columns_to_keep = ['DateTime'] + parameters
        station_df = station_df[columns_to_keep]
        
        # Store the DataFrame in a dictionary with the station name as the key
        teletrans_dfs[station] = station_df

    #Meteo Data
    for station in weather_stations:
        station_df = df[(df['NumEstacion'] == station) & (df['NumFuncion'].isin(['Accumulation', 'Average']))].copy()
        station_df.rename(columns={'Fecha': 'DateTime'}, inplace=True)
        parameters = ['Wind Speed', 'Wind Direction', 'Air Temperature', 'Surface Temperature 1', 'Surface Temperature 2', 'Rain']
        for param in parameters:
            filtered_df = station_df[station_df['NumParametro'] == param][['DateTime', 'Valor']].copy()
            filtered_df.rename(columns={'Valor': param}, inplace=True)
            if 'DateTime' in station_df.columns:
                station_df = station_df.merge(filtered_df, on='DateTime', how='left')
        
        # Select only the 'DateTime' and new parameter columns
        columns_to_keep = ['DateTime'] + parameters
        station_df = station_df[columns_to_keep]

        # Store the DataFrame in a dictionary with the station name as the key
        meteo_dfs[station] = station_df

    #root.destroy()


def browse_files():
    global start_time
    start_time = time.time()
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    datos_file_path = filedialog.askopenfilename(initialdir=downloads_path, title="Select the Datos File", filetypes=[("Excel files", "*.xlsx *.xls")])
    production_file = filedialog.askopenfilename(initialdir=downloads_path, title="Select the Monthly Billing File", filetypes=[("*.csv")])
    #datos_file_path = r"G:\Shared drives\O&M\NCC Automations\In Progress\Performance Analysis - Python\Brandon\Python Switch\Lily\DatosOCTOBER2024.csv"
    process_met_stations(datos_file_path)
    calcells_file = filedialog.askopenfilename(initialdir=downloads_path, title="Select the Call Cells CSV File", filetypes=[("*.csv")])  
    inv3_file = filedialog.askopenfilename(initialdir=downloads_path, title="Select the Lily Inv3 CSV File", filetypes=[("*.csv")])  
    invall_file = filedialog.askopenfilename(initialdir=downloads_path, title="Select the INV ALL CSV File", filetypes=[("*.csv")])  
    ppc_file = filedialog.askopenfilename(initialdir=downloads_path, title="Select the PPC CSV File", filetypes=[("*.csv")])  
    sub_file = filedialog.askopenfilename(initialdir=downloads_path, title="Select the Sub CSV File", filetypes=[("*.csv")])  
    submfmt_file = filedialog.askopenfilename(initialdir=downloads_path, title="Select the Sub MFMT CSV File", filetypes=[("*.csv")])  
    process_lily_scada_data(calcells_file, inv3_file, invall_file, ppc_file, sub_file, submfmt_file, production_file)
    



root = tk.Tk()
root.title("Lily Performance Analysis")
lbl = tk.Label(root, text="Please select the Datos File")
browse = tk.Button(root, text="Browse Files", command= lambda: browse_files(), width= 50, height=5)
lbl.pack()
browse.pack()

root.mainloop()