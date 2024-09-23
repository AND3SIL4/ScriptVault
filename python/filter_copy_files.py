import pandas as pd
import os

def main(params):
    try:
        # Set variables
        file_path = params.get("file_path")
        sheet_name = params.get("sheet_name")
        temp_folder_path = params.get("temp_folder_path")
        temp_file_name = params.get("temp_file_name")
        cut_off_date_input = params.get("cut_off_date")
        column_index = int(params.get("column_index"))
        start_date_input = params.get("start_date_input")

        # Validate parameters
        if not all([file_path, sheet_name, temp_folder_path, column_index, start_date_input, cut_off_date_input]):
            return "Error: Missing one or more parameters."

        print(file_path)

        # Check if the file exists
        if not os.path.exists(file_path):
            return f"Error: File {file_path} not found."

        # Convert dates to datetime
        cut_off_date = pd.to_datetime(cut_off_date_input, format='%d/%m/%Y')
        start_date = pd.to_datetime(start_date_input, format='%d/%m/%Y')

        # Open file using pandas
        df = pd.read_excel(file_path, engine="openpyxl", sheet_name=sheet_name)

        # Convert the column to datetime using the column index
        df.iloc[:, column_index] = pd.to_datetime(df.iloc[:, column_index], format='%d/%m/%Y')

        # Filter from start_date to cut_off_date
        filter_file = df[(df.iloc[:, column_index] >= start_date) & (df.iloc[:, column_index] <= cut_off_date)]

        new_file = os.path.join(temp_folder_path, temp_file_name)

        # Create temp folder if it doesn't exist
        if not os.path.exists(temp_folder_path):
            os.makedirs(temp_folder_path)

        # Copy to temp file to make validations
        filter_file.to_excel(new_file, index=False, sheet_name=sheet_name)

        return "File copied successfully"

    except Exception as e:
        return f"Error: {e}"


if  __name__ == "__main__":
    params = {
        # Set variables
        "file_path": r"C:/ProgramData/AutomationAnywhere/Bots/Logs/AD_RCSN_SabanaPagosYBasesParaSinestralidad/InputFolder/BASE DE REPARTO 2024.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "temp_folder_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder",
        "temp_file_name": "BASE REPARTO 2024.xlsx",
        "cut_off_date": "01/01/2024",
        "column_index": "24",
        "start_date_input": "30/07/2024"
    }

    print(main(params))