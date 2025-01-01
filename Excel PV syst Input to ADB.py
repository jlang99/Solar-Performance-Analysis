import pandas as pd
import pyodbc

def insert_excel_to_access(excel_file_path, access_db_path, table_name):
    # Read the Excel file
    df = pd.read_excel(excel_file_path)
    
    # Create a connection to the Access database
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=' + access_db_path + ';'
    )
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    
    # Insert data into the Access database table
    for index, row in df.iterrows():
        # Create a list of values to insert
        values = tuple(row)
        
        # Create a SQL insert statement
        placeholders = ', '.join(['?'] * len(values))
        sql = f"INSERT INTO {table_name} VALUES ({placeholders})"
        
        # Execute the SQL statement
        cursor.execute(sql, values)
    
    # Commit the transaction
    conn.commit()
    
    # Close the connection
    cursor.close()
    conn.close()

# Example usage
excel_file_path = r"C:\Users\OMOPS.AzureAD\Downloads\Wellons PV syst.xlsx"
access_db_path = r"G:\Shared drives\O&M\NCC Automations\In Progress\Performance Analysis - Python\Brandon\Python Switch\Also Energy Sites\PVsyst (Josephs Edits).accdb"
table_name = 'PVsystHourly'


insert_excel_to_access(excel_file_path, access_db_path, table_name)
print("Done!")