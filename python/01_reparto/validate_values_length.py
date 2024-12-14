import pandas as pd  # type: ignore
import os
import re


def main(params: dict):
    try:
        ##Set initial variables
        file_path = params.get("file_path")
        in_file = params.get("in_file")
        sheet_name = params.get("sheet_name")
        col_name = params.get("col_name")
        in_sheet = params.get("in_sheet")
        list_file = params.get("list_file")
        list_col = params.get("list_col")

        ##Validate if all the input params are present
        if not all([file_path, in_file, sheet_name, col_name, in_sheet, list_file]):
            return "Error: an input param is missing"

        ##Read the work book
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
        list_df = pd.read_excel(list_file, sheet_name="LISTAS", engine="openpyxl")

        lst = list_df[list_col].dropna().astype(str).to_list()

        ##Validate the length in a specific column
        is_in = df[~df[col_name].isin(lst)].copy()

        print(is_in)

        if is_in.empty:
            return "Validación realizada, no se encontraron inconsistencias"
        else:
            col_index = df.columns.get_loc(col_name) + 1  # Get column number (1-based)
            is_in["COORDENADAS"] = is_in.apply(
                lambda row: f"{get_excel_column_name(col_index)}{row.name + 2}", axis=1
            )

            ##Register into the inconsistencies file
            new_sheet = in_sheet
            ##Validate is the inconsistencies file exist
            if os.path.exists(in_file):
                with pd.ExcelFile(in_file, engine="openpyxl") as xls:
                    if new_sheet in xls.sheet_names:
                        existing_file = pd.read_excel(
                            xls, engine="openpyxl", sheet_name=new_sheet
                        )
                        existing_file = existing_file.dropna(
                            how="all", axis=1
                        )  ##Does not include the empty columns
                        is_in = pd.concat([existing_file, is_in], ignore_index=True)

            ##Store data
            with pd.ExcelWriter(
                in_file, engine="openpyxl", mode="a", if_sheet_exists="replace"
            ) as writer:
                is_in.to_excel(writer, sheet_name=new_sheet, index=False)
                return "Inconsistencias registradas correctamente"

    except Exception as e:
        return f"Error: {e}"


def clean(string):
    """Validate the white spaces in strings."""
    string = re.sub(r"\s+", " ", string).strip()
    return string


def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


if __name__ == "__main__":
    params = {
        "file_path": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
        "in_file": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "col_name": "RAMO",
        "in_sheet": "ValidacionRamos",
    }

    print(main(params))

"HOMOLOGACION AMPAROS"