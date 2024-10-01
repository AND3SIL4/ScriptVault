import pandas as pd
import os

def main(params: dict):
    try:
        ##Set the variables
        file_path: str = params.get("file_path")
        sheet_name: str = params.get("sheet_name")
        col_idx1: int = int(params.get("col_idx1"))
        col_idx2: int  = int(params.get("col_idx2"))
        inconsistencies_file: str = params.get("in_file")
        new_sheet: str = params.get("new_sheet")
        need_iaxis: bool = bool(params.get("need_iaxis"))
        list_file: str = params.get("list_file")
        except_idx = int(params.get("except_idx"))
        sheet_name_list = params.get("sheet_name_list")

        ##Local variables
        validated_values: dict = {}

        ##Validate if all the required inputs are present
        if not all([
            file_path, 
            sheet_name,
            new_sheet,
            inconsistencies_file,
            list_file,
            sheet_name_list
        ]):
            return "Error: an input param is missing"

        ##Read the book and load it
        df: pd.DataFrame = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
        list_df: pd.DataFrame = pd.read_excel(list_file, sheet_name=sheet_name_list, engine="openpyxl")

        ##Apply validation
        df["is_valid"] = df.apply(
            lambda row: validate(
                str(row.iloc[col_idx1]), ##Key
                str(row.iloc[col_idx2]), ##Value
                validated_values, ##Dictionary as a global variable
                list_df.iloc[:, except_idx].dropna().astype(str).values.tolist()
            ), 
            axis=1
        ).astype(bool)

        ##Make a column to know if is "TRUE" or "FALSE"
        filtered_file = df[~df["is_valid"]].copy()
        print(filtered_file)

        if not need_iaxis:
            if not filtered_file.empty:
                ##Get the coordinates
                filtered_file["COORDINATE_1"]  = filtered_file.apply(
                    lambda row: f"{get_excel_column_name(col_idx1 + 1)}{row.name + 2}", axis=1
                )
                filtered_file["COORDINATE_2"] = filtered_file.apply(
                    lambda row: f"{get_excel_column_name(col_idx2 + 1)}{row.name + 2}", axis=1
                )
                ##Store the inconsistencies into the inconsistencies file
                return append_inconsistencias(inconsistencies_file, new_sheet, filtered_file)                
            else:
                return "Validacion realizada, no se encontraron inconsistencias"
        else:
            ##Get the inconsistencies values
            inconsistencies_col = filtered_file.iloc[:, col_idx1]
            new_df = df[df.iloc[:, col_idx1].isin(inconsistencies_col)].copy()
            ##Get the coordinates
            if not new_df.empty:
                new_df["COORDINATE_1"] = new_df.apply(
                    lambda row: f"{get_excel_column_name(col_idx1 + 1)}{row.name + 2}", axis=1
                )
                new_df["COORDINATE_2"] = new_df.apply(
                    lambda row: f"{get_excel_column_name(col_idx2 + 1)}{row.name + 2}", axis=1
                )
                ##Store the inconsistencies into the inconsistencies file
                append_inconsistencias(inconsistencies_file, new_sheet, new_df)
                return True
            else:
                return False
    except Exception as e:
        return f"Error: {e}"

def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ''
    while n > 0:
        n, remainder = divmod(n-1, 26)
        result = chr(65 + remainder) + result
    return result

def validate(key: str, value: str, validated: dict[str], exception_list: list[str]) -> bool:
    ##Set the local variables
    if key in validated:
        if validated[key] == value or key in exception_list:
            return True
    else:
        validated[key] = value
        return True


def append_inconsistencias(file_path: str, new_sheet: str, data_frame) -> None:
    """This function get the inconsistencies data frame and append it into the inconsistencies file"""
    if os.path.exists(file_path):
        with pd.ExcelFile(file_path, engine="openpyxl") as xls:
            if new_sheet in xls.sheet_names:
                existing = pd.read_excel(xls, sheet_name=new_sheet, engine="openpyxl")
                data_frame = pd.concat([existing, data_frame], ignore_index=True)

        with pd.ExcelWriter(
            file_path, engine="openpyxl", if_sheet_exists="replace", mode="a"
        ) as writer:
            data_frame.to_excel(writer, index=False, sheet_name=new_sheet)        
            return "Inconsistencias registradas correctamente"

if __name__ == "__main__":
    params = {
        "file_path": "C:/ProgramData/AutomationAnywhere/Bots/Logs/AD_RCSN_SabanaPagosYBasesParaSinestralidad/TempFolder/BASE DE REPARTO 2024.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "col_idx1": "0",
        "col_idx2": "18",
        "in_file": "C:/ProgramData/AutomationAnywhere/Bots/Logs/AD_RCSN_SabanaPagosYBasesParaSinestralidad/OutputFolder/Inconsistencias/InconBaseReparto.xlsx",
        "new_sheet": "SiniestroUnicaPoliza",
        "need_iaxis": True,
        "list_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\Listas - BOT.xlsx",
        "except_idx": "2",
        "sheet_name_list": "EXCEPCIONES COPARACION SINIES"
    }

    print(main(params))
