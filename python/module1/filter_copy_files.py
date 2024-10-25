import pandas as pd  # type: ignore
import os


def main(params):
    try:
        # Set variables
        file_path = params.get("file_path")
        sheet_name = params.get("sheet_name")
        temp_file = params.get("temp_file")
        cut_off_date_input = params.get("cut_off_date")
        column_index = int(params.get("column_index"))
        start_date_input = params.get("start_date_input")

        # Validate parameters
        if not all(
            [
                file_path,
                sheet_name,
                temp_file,
                column_index,
                start_date_input,
                cut_off_date_input,
            ]
        ):
            return "ERROR: Missing one or more parameters."

        # Check if the file exists
        if not os.path.exists(file_path):
            return f"ERROR: File {file_path} not found."

        # Convert dates to datetime
        cut_off_date = pd.to_datetime(cut_off_date_input, format="%d/%m/%Y")
        start_date = pd.to_datetime(start_date_input, format="%d/%m/%Y")

        # Open file using pandas
        df = pd.read_excel(file_path, engine="openpyxl", sheet_name=sheet_name)
        otros_gastos = pd.read_excel(file_path, sheet_name="OGDS", engine="openpyxl")
        ## Assign the columns of the first data frame
        otros_gastos.columns = df.columns

        # Convert the column to datetime using the column index
        df.iloc[:, column_index] = pd.to_datetime(
            df.iloc[:, column_index], format="%d/%m/%Y"
        )

        # Filter from start_date to cut_off_date
        filter_file = df[
            (df.iloc[:, column_index] >= start_date)
            & (df.iloc[:, column_index] <= cut_off_date)
        ]

        final_data_frame: pd.DataFrame = pd.concat(
            [filter_file, otros_gastos], ignore_index=True
        )
        # Copy to temp file to make validations
        final_data_frame.to_excel(temp_file, index=False, sheet_name=sheet_name)
        return "SUCCESS: file copied successfully"

    except Exception as e:
        return f"Error: {e}"


if __name__ == "__main__":
    params = {
        # Set variables
        "file_path": r"C:/ProgramData/AutomationAnywhere/Bots/Logs/AD_RCSN_SabanaPagosYBasesParaSinestralidad/InputFolder/BASE DE REPARTO 2024.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "temp_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
        "start_date_input": "01/01/2024",
        "column_index": "24",
        "cut_off_date": "24/10/2024",
    }

    print(main(params))
