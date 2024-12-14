import pandas as pd #type: ignore
import os
import re

def sudameris(params: dict):
    try:
        ## Set initial variables
        sudameris_bank = params.get("sudameris_bank")
        base_reparto = params.get("base_reparto")
        sheet_sudameris = params.get("sheet_sudameris")
        sheet_reparto = params.get("sheet_reparto")
        initial_date = params.get("initial_date")
        cut_off_date = params.get("cut_off_date")
        date_col_idx = int(params.get("date_col_idx"))
        vs_col = int(params.get("vs_col"))
        in_file = params.get("in_file")

        ## Validate if the dictionary values are present
        if not all([
            sudameris_bank, base_reparto, sheet_sudameris, sheet_reparto, 
            initial_date, cut_off_date, in_file
        ]):
            return "Error: an input param required is missing"
        
        ## Converts dates to datetime format
        initial_date = pd.to_datetime(initial_date, format="%d/%m/%Y")
        cut_off_date = pd.to_datetime(cut_off_date, format="%d/%m/%Y")

        ## Read work books
        sudameris: pd.DataFrame = pd.read_excel(
            sudameris_bank, sheet_name=sheet_sudameris, engine="openpyxl"
        )
        reparto: pd.DataFrame = pd.read_excel(
            base_reparto, sheet_name=sheet_reparto, engine="openpyxl"
        )
        
        ## Filter workbooks based on date column
        filtered_sudameris = sudameris[(sudameris.iloc[:, date_col_idx] >= initial_date) & 
                                       (sudameris.iloc[:, date_col_idx] <= cut_off_date)]

        filtered_sudameris = filtered_sudameris.copy()
        ## Check if all values from Sudameris are in Reparto
        filtered_sudameris["is_in_reparto"] = filtered_sudameris.iloc[:, vs_col].isin(reparto.iloc[:, vs_col])
        ## Find rows in Sudameris that are not in Reparto
        not_in_reparto: pd.DataFrame = filtered_sudameris[~filtered_sudameris["is_in_reparto"]].copy()

        not_in_reparto["is_valid"] = not_in_reparto.iloc[:, vs_col].astype(str).apply(
            lambda x: validate_comments(x)
        )

        inconsistencies = not_in_reparto[~not_in_reparto["is_valid"].astype(bool)].copy()

        ##!Unit tests
        print(inconsistencies.iloc[:, -1])
        #print(reparto.iloc[:, vs_col])

        ## Store and return the inconsistencies
        if inconsistencies.empty:
            return "All values from 'Banco Sudameris' are present in 'Base Reparto'"
        else:
            inconsistencies['COORDENADAS'] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(vs_col + 1)}{row.name + 2}", axis=1
            )       
            ## Create a sheet name to store the inconsistencies
            new_sheet = "BancoSudameris"
            
            if os.path.exists(in_file):
                with pd.ExcelFile(in_file, engine="openpyxl") as xls:
                    if new_sheet in xls.sheet_names:
                        existing_file = pd.read_excel(xls, sheet_name=new_sheet, engine="openpyxl")
                        inconsistencies = pd.concat([existing_file, inconsistencies], ignore_index=True)

            with pd.ExcelWriter(in_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                inconsistencies.to_excel(writer, index=False, sheet_name=new_sheet)
                return "Inconsistencies registered successfully"

    except Exception as e:
        return f"Error: {e}"

def agrario(params: dict):
    try:
        ## Set initial variables
        agrario_bank = params.get("agrario_bank")
        base_reparto = params.get("base_reparto")
        sheet_reparto = params.get("sheet_reparto")
        initial_date = params.get("initial_date")
        cut_off_date = params.get("cut_off_date")
        date_col_idx = int(params.get("date_col_idx"))
        vs_col = int(params.get("vs_col"))
        in_file = params.get("in_file")

        ## Validate if the dictionary values are present
        if not all(
            [
                agrario_bank, base_reparto, sheet_reparto, initial_date, cut_off_date, in_file
            ]
        ):
            return "Error: an input param is missing"
        
        ## Converts dates to datetime format
        initial_date = pd.to_datetime(initial_date, format="%d/%m/%Y")
        cut_off_date = pd.to_datetime(cut_off_date, format="%d/%m/%Y")

        ## Read work books
        agrario_sheets = ["DEUDORES - LINEA GENERAL", "EMPLEADOS BANCO AGRARIO", "TARJETAS  BANCO AGRARIO"]
        data_frames = (pd.read_excel(agrario_bank, sheet_name=sheet) for sheet in agrario_sheets)
        agrario = pd.concat(data_frames, ignore_index=True)
        reparto = pd.read_excel(base_reparto, sheet_name=sheet_reparto, engine="openpyxl")
        
        ## Filter workbooks based on date column
        filtered_agrario = agrario[(agrario.iloc[:, date_col_idx] >= initial_date) & 
                                    (agrario.iloc[:, date_col_idx] <= cut_off_date)]

        ## Check if all values from agrario are in Reparto
        filtered_agrario = filtered_agrario.copy()
        filtered_agrario["is_in_reparto"] = filtered_agrario.iloc[:, vs_col].isin(reparto.iloc[:, vs_col])

        ## Find rows in agrario that are not in Reparto
        not_in_reparto: pd.DataFrame = filtered_agrario[~filtered_agrario["is_in_reparto"]].copy()

        not_in_reparto["is_valid"] = not_in_reparto.iloc[:, vs_col].astype(str).apply(
            lambda x: validate_comments(x)
        )

        inconsistencies = not_in_reparto[~not_in_reparto["is_valid"]].copy()

        ##!For debugging
        print(inconsistencies)

        ## Store and return the inconsistencies
        if inconsistencies.empty:
            return "All values from 'Banco agrario' are present in 'Base Reparto'"
        else:
            inconsistencies['COORDENADAS'] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(vs_col + 1)}{row.name + 2}", axis=1
            )

            ## Create a sheet name to store the inconsistencies
            new_sheet = "BancoAgrario"
            
            if os.path.exists(in_file):
                with pd.ExcelFile(in_file, engine="openpyxl") as xls:
                    if new_sheet in xls.sheet_names:
                        existing_file = pd.read_excel(xls, sheet_name=new_sheet, engine="openpyxl")
                        inconsistencies = pd.concat([existing_file, inconsistencies], ignore_index=True)

            with pd.ExcelWriter(in_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                inconsistencies.to_excel(writer, index=False, sheet_name=new_sheet)
                return "Inconsistencies registered successfully"

    except Exception as e:
        return f"Error: {e}"

def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ''
    while n > 0:
        n, remainder = divmod(n-1, 26)
        result = chr(65 + remainder) + result
    return result


def validate_comments(string: str) -> bool:
    """This method works to validate is the input has any string character
    and know if is a comment"""
    try:
        return bool(re.search(r"[^0-9]", string))
    except:
        return False
        
if __name__ == "__main__":
    params = {
        "agrario_bank": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\BANCO AGRARIO 2024.xlsx",
        "base_reparto": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
        "sheet_reparto": "CASOS NUEVOS",
        "initial_date": "01/01/2024",
        "cut_off_date": "30/06/2024",
        "date_col_idx": "24",
        "vs_col": "0",    
        "in_file": "C:/ProgramData/AutomationAnywhere/Bots/Logs/AD_RCSN_SabanaPagosYBasesParaSinestralidad/OutputFolder/Inconsistencias/InconBaseReparto.xlsx",
    }

    print(agrario(params))