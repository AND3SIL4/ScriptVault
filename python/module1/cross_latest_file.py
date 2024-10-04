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
        cut_off_date = pd.to_datetime(datetime.now().date(), format="%d/%m/%Y")
        cut_off_date = cut_off_date - pd.DateOffset(months=1)
        year = datetime.today().year
        date = f"01/01/{year}"
        initial_date = pd.to_datetime(date, format="%d/%m/%Y")

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
        ##Set the variables to store the values of column "Valor Reserva"
        reserva_latest_cut = file_filtered.iloc[:, 34].sum()
        reserva_current_cut = latest_filtered.iloc[:, 34].sum()

        # Validate if the values of "Reserva" are the same
        return reserva_latest_cut == reserva_current_cut
    except Exception as e:
        return f"ERROR: {e}"


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
        "inconsistencias_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "latest_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Historico base reparto\BASE DE REPARTO 2024.xlsx",
        "sheet_latest_name": "CASOS NUEVOS",
        "col_idx": "24",
    }

    print(main(params))
