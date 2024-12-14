import pandas as pd #type: ignore
import os
import re
from typing import List


def main(params: dict):
    try:
        ##Set initial variables
        file_path = params.get("file_path")
        inconsistencias_file = params.get("inconsistencias_file")
        sheet_name = params.get("sheet_name")
        col_idx = int(params.get("col_idx"))
        list_file = params.get("list_file")
        list_col = int(params.get("list_col"))

        ##Validate if a para input is missing
        if not all([file_path, inconsistencias_file, sheet_name, list_file]):
            return "Error: An input param is missing"

        ##Load and read work book
        df: pd.DataFrame = pd.read_excel(
            file_path, engine="openpyxl", sheet_name=sheet_name
        )
        list_df = pd.read_excel(
            list_file, engine="openpyxl", sheet_name="CARACTERES ESPECIALES"
        )

        ##Apply validation to the file
        df["is_valid"] = (
            df.iloc[:, col_idx]
            .astype(str)
            .apply(
                lambda x: clean(
                    x, list_df.iloc[:, list_col].dropna().astype(str).values.tolist()
                )
            )
        )

        ##Filter the file and store the inconsistencies
        inconsistencies = df[~df["is_valid"]].copy()

        # print(inconsistencies.iloc[:, col_idx])  ##!Comment to deploy
        print(inconsistencies)

        if inconsistencies.empty:
            return (
                "Validación realizada correctamente, no se encontraron inconsistencias"
            )
        else:
            # Add a column with Excel coordinates (e.g., A2, B3) of the inconsistent cells
            inconsistencies["COORDENADAS"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(col_idx + 1)}{row.name + 2}",
                axis=1,
            )

            new_sheet = "CaracteresEspeciales"
            if os.path.exists(inconsistencias_file):
                with pd.ExcelFile(inconsistencias_file, engine="openpyxl") as xls:
                    if new_sheet in xls.sheet_names:
                        existing_file = pd.read_excel(
                            xls, engine="openpyxl", sheet_name=new_sheet
                        )
                        existing_file = existing_file.dropna(how="all", axis=1)
                        inconsistencies = pd.concat(
                            [existing_file, inconsistencies], ignore_index=True
                        )

            with pd.ExcelWriter(
                inconsistencias_file,
                engine="openpyxl",
                mode="a",
                if_sheet_exists="replace",
            ) as writer:
                inconsistencies.to_excel(writer, index=False, sheet_name=new_sheet)
                return "Inconsistencias registradas correctamente"

    except Exception as e:
        return f"Error: {e}"


def clean(string: str, exception_list: List[str]) -> bool:
    # Validate if string contains any does not allowed character
    if string in exception_list:
        return True

    if (
        re.search(r"[^a-zA-Z0-9\sñÑ]", string)
        or "  " in string
        or string.startswith(" ")
        or string.endswith(" ")
    ):
        return False

    return True


def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
        "inconsistencias_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "col_idx": "17",
        "list_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\Listas - BOT.xlsx",
        "list_col": "0",
    }

    print(main(params))
