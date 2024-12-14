import pandas as pd  # type: ignore
from datetime import datetime
import os


def main(params: dict) -> None:
    try:
        ##Set local variables
        path_file: str = params.get("file_path")
        sheet_name: str = params.get("sheet_name")
        inconsistencies_file: str = params.get("inconsistencias_file")
        latest_file: str = params.get("latest_file")
        sheet_latest_name: str = params.get("sheet_latest_name")
        col_idx: int = int(params.get("col_idx"))
        cut_date: str = params.get("cut_date")

        ##Validate if all the required files are present
        if not all(
            [
                path_file,
                sheet_name,
                inconsistencies_file,
                latest_file,
                sheet_latest_name,
            ]
        ):
            return "ERROR: an input required param is missing"

        ##Set the dates for filtering the files
        year = datetime.today().year
        date = f"01/01/{year}"
        initial_date = pd.to_datetime(date, format="%d/%m/%Y")
        cut = pd.to_datetime(cut_date, format="%d/%m/%Y")
        cut_off_date = cut - pd.DateOffset(months=1)

        ##Load the work books needed
        ##Current base reparto
        path_file_df: pd.DataFrame = pd.read_excel(
            path_file, sheet_name=sheet_name, engine="openpyxl"
        )
        file_filtered: pd.DataFrame = path_file_df[
            (path_file_df.iloc[:, col_idx] > initial_date)
            & (path_file_df.iloc[:, col_idx] < cut_off_date)
        ]

        # Base reparto latest month
        latest_file_df: pd.DataFrame = pd.read_excel(
            latest_file, sheet_name=sheet_latest_name, engine="openpyxl"
        )
        latest_filtered: pd.DataFrame = latest_file_df[
            (latest_file_df.iloc[:, col_idx] > initial_date)
            & (latest_file_df.iloc[:, col_idx] < cut_off_date)
        ]
        file_filtered: pd.DataFrame = file_filtered.iloc[:, :111]

        latest_filtered.columns = file_filtered.columns

        ##Key name
        key_name = "SINIESTRO+RADICADO+AMPARO+RESERVA"

        file_filtered[key_name] = (
            file_filtered.iloc[:, 0].astype(str)
            + "-"
            + file_filtered.iloc[:, 2].astype(str)
            + "-"
            + file_filtered.iloc[:, 32].astype(str)
            + "-"
            + file_filtered.iloc[:, 34].astype(str)
        )

        latest_filtered[key_name] = (
            latest_filtered.iloc[:, 0].astype(str)
            + "-"
            + latest_filtered.iloc[:, 2].astype(str)
            + "-"
            + latest_filtered.iloc[:, 32].astype(str)
            + "-"
            + latest_filtered.iloc[:, 34].astype(str)
        )

        file_not_in_latest = file_filtered[
            ~file_filtered[key_name].isin(latest_filtered[key_name])
        ].copy()
        if not file_not_in_latest.empty:
            file_not_in_latest["FILE_TO_FIND"] = "ARCHIVO ACTUAL"

        latest_not_in_file = latest_filtered[
            ~latest_filtered[key_name].isin(file_filtered[key_name])
        ].copy()

        # Initialize inconsistencies as an empty DataFrame
        inconsistencies = pd.DataFrame()

        if not latest_not_in_file.empty:
            latest_not_in_file["FILE_TO_FIND"] = "ARCHIVO ANTERIOR"

        if not file_not_in_latest.empty and not latest_not_in_file.empty:
            # Combine the results of both mismatches
            inconsistencies = pd.concat([file_not_in_latest, latest_not_in_file])
        elif not file_not_in_latest.empty and latest_not_in_file.empty:
            inconsistencies = file_not_in_latest
        elif file_not_in_latest.empty and not latest_not_in_file.empty:
            inconsistencies = latest_not_in_file

        if not inconsistencies.empty:
            # Add a column with Excel coordinates (e.g., A2, B3) of the inconsistent cells
            inconsistencies["COORDENADAS_1"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(1)}{row.name + 2}",
                axis=1,
            )
            inconsistencies["COORDENADAS_2"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(3)}{row.name + 2}",
                axis=1,
            )
            inconsistencies["COORDENADAS_3"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(33)}{row.name + 2}",
                axis=1,
            )
            inconsistencies["COORDENADAS_4"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(35)}{row.name + 2}",
                axis=1,
            )

            return append_inconsistencias(
                inconsistencies_file, "ReservaCorteAnterior", inconsistencies
            )
        else:
            return "Validacion realizada, no se encontraron inconsistencias"

    except Exception as e:
        return f"ERROR: {e}"


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
        "latest_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Historico base reparto\BASE DE REPARTO 092024.xlsx",
        "sheet_latest_name": "CASOS NUEVOS",
        "col_idx": "24",
        "cut_date": "25/10/2024",
    }

    print(main(params))
