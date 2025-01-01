import pandas as pd

def csv_to_excel(csv_file_path, excel_file_path):
    # Read the CSV file into a DataFrame
    df = pd.read_csv(csv_file_path, encoding='latin1', delimiter=';')
    
    # Write the DataFrame to an Excel file
    df.to_excel(excel_file_path, index=False)

if __name__ == "__main__":
    # Define the path to the CSV file and the desired Excel file

    w_path_og = r"C:\Users\OMOPS.AzureAD\Downloads\Selma Wellons_Project_HourlyRes_7.CSV"
    w_path = r"C:\Users\OMOPS.AzureAD\Downloads\Wellons PV syst.xlsx"

    #Convert CSV to Excel
    csv_to_excel(w_path_og, w_path)
    
    print(f"CSV file '{w_path_og}' has been successfully parsed into Excel file '{w_path}'")