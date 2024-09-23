import pandas as pd
import os


def main(params: dict):
    try:
        ##Set initial variables
        file_path: str = params.get("file_path")
        col_idx: int = int(params.get("col_name"))
        inconsistencias_file: str = params.get("inconsistencias_file")
        sheet_name: str = "CASOS NUEVOS"
        mandatory = params.get("mandatory")
        alpha_numeric = params.get("alpha_numeric")

        if not all(
            [
                ##Set initial variables
                file_path,
                col_idx,
                inconsistencias_file,
                sheet_name,
                mandatory,
                alpha_numeric,
            ]
        ):
            return "ERROR: An input required param is missing"

        ##Read book using pandas
        df: pd.DataFrame = pd.read_excel(
            file_path, sheet_name=sheet_name, engine="openpyxl"
        )

        ##Filter information and validate if the current data if number type
        df["is_valid"] = df.apply(
            lambda row: is_valid(
                str(row.iloc[col_idx]), mandatory, alpha_numeric, str(row.iloc[15])
            ),
            axis=1,
        )
        ##Add inconsistencies to a filtered data frame
        filtered_file = df[~df["is_valid"]].copy()

        ##Return and store the result
        if filtered_file.empty:
            return "Validación correcta, no se encontraron inconsistencias"
        else:
            filtered_file["COORDENADAS"] = filtered_file.apply(
                lambda row: f"{get_excel_column_name(col_idx + 1)}{row.name + 2}",
                axis=1,
            )

            new_sheet_name = "TipoNumeroBancoW"
            if os.path.exists(inconsistencias_file):
                with pd.ExcelFile(inconsistencias_file, engine="openpyxl") as xls:
                    if new_sheet_name in xls.sheet_names:
                        exiting_df = pd.read_excel(
                            xls, sheet_name=new_sheet_name, engine="openpyxl"
                        )
                        filtered_file = pd.concat(
                            [exiting_df, filtered_file], ignore_index=True
                        )

            with pd.ExcelWriter(
                inconsistencias_file,
                engine="openpyxl",
                mode="a",
                if_sheet_exists="replace",
            ) as writer:
                filtered_file.to_excel(writer, sheet_name=new_sheet_name, index=False)
                return "Inconsistencias registradas correctamente"

    except Exception as e:
        return f"Error: {e}"


def is_valid(value: str, mandatory: list[str], alpha_numeric: list[str], tomador: str):
    if value.strip() == "" or value.lower() == "nan":
        if tomador in mandatory:
            return False ##!Should return False
        else: 
            return True
    else:
      try:
        #Parse the value into string
        int(value)
        return True
      except:
        if tomador in alpha_numeric and value.isalnum():  
          return True
        print(tomador)
        return False

def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

if __name__ == "__main__":
    params = {
        "file_path": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\BASE DE REPARTO 2024.xlsx",
        "col_name": "98",
        "inconsistencias_file": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
        "mandatory": [
            "BANCO W SA",
            "BANCO AGRARIO DE COLOMBIA SA",
            "BANCO GNB SUDAMERIS",
        ],
        "alpha_numeric": ["BANCO W SA", "BAN"],
    }

    print(main(params))
