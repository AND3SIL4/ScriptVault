import pandas as pd
import os
import re

def main(params: dict):
    try:    
        ##Set initial variables
        file_path = params.get("file_path")
        inconsistencias_file = params.get("inconsistencias_file")
        sheet_name = params.get("sheet_name")
        col_name = params.get("col_name")

        ##Validate if a para input is missing
        if not all([file_path, inconsistencias_file, sheet_name, col_name]):
            return "Error: An input param is missing"

        ##Load and read work book
        df = pd.read_excel(file_path, engine="openpyxl", sheet_name=sheet_name)

        ##Filter
        df["col"] = df[col_name].apply(lambda x: clean(str(x)))
        
        ##Store the 
        inconsistencies = df[df[col_name] != df["col"]].copy()
        
        print(inconsistencies) ##!Comment to deploy
        if inconsistencies.empty:
            return "Validación realizada correctamente, no se encontraron inconsistencias"
        else:
            #Add a column with Excel coordinates (e.g., A2, B3) of the inconsistent cells
            col_index = df.columns.get_loc(col_name) + 1  # Get column number (1-based)
            inconsistencies['COORDENADAS'] = inconsistencies.apply(
              lambda row: f"{get_excel_column_name(col_index)}{row.name + 2}", axis=1
            )

            new_sheet = "CaracteresEspeciales"
            if os.path.exists(inconsistencias_file):
                with pd.ExcelFile(inconsistencias_file, engine="openpyxl") as xls:
                    if new_sheet in xls.sheet_names:
                        existing_file = pd.read_excel(xls, engine="openpyxl", sheet_name=new_sheet)
                        existing_file = existing_file.dropna(how="all", axis=1)
                        inconsistencies = pd.concat([existing_file, inconsistencies], ignore_index=True)
            
            with pd.ExcelWriter(inconsistencias_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                inconsistencies.to_excel(writer, index=False, sheet_name=new_sheet)
                return "Inconsistencias registradas correctamente"
        
    except Exception as e:
        return f"Error: {e}"

def clean(string: str):
    ##Dictionary to change the special characters
    replacements = {
        'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
        'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U',
        ##'ñ': 'n', 'Ñ': 'N' ##!Uncomment to validate if "ñ" is not valid
    }

    ##Replace the string with accented charts to non accented chars
    for accented_char, non_accented_chat in replacements.items():
        string = str(string).replace(accented_char, non_accented_chat)

    ##Clean white spaces of string with a regex expression
    string = re.sub(r"\s+", " ", string).strip()
    ##Delete other special characters
    string = re.sub(r"[^a-zA-Z0-9\sñÑ]", "", string)
    ##Return the result
    return string

##!NEW FUNCTION HERE!!
def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ''
    while n > 0:
        n, remainder = divmod(n-1, 26)
        result = chr(65 + remainder) + result
    return result

if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
        "inconsistencias_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "col_name": "RIESGO / ASEGURADO",
    }

    print(main(params))
