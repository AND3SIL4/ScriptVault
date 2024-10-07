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
        cut_off_date = pd.to_datetime(cut_date, format="%d/%m/%Y")
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

        print(reserva_latest_cut)
        print(reserva_current_cut)

        # Validate if the values of "Reserva" are the same
        if reserva_latest_cut != reserva_current_cut:
            file_filtered["LLAVE_CRUCE"] = (
                file_filtered.iloc[:, 0].astype(str)
                + "-"
                + file_filtered.iloc[:, 2].astype(str)
                + "-"
                + file_filtered.iloc[:, 32].astype(str)
            )
            file_filtered["RESERVA_CRUCE"] = file_filtered.iloc[:, 34]

            latest_filtered["LLAVE_CRUCE"] = (
                latest_filtered.iloc[:, 0].astype(str)
                + "-"
                + latest_filtered.iloc[:, 2].astype(str)
                + "-"
                + latest_filtered.iloc[:, 32].astype(str)
            )

            ##Make crossover file
            merge_df: pd.DataFrame = latest_filtered.merge(
                file_filtered[["LLAVE_CRUCE", "RESERVA_CRUCE"]],
                on="LLAVE_CRUCE",
                how="left",
                suffixes=("_OLD", "_NEW"),
            )

            merge_df["VALIDATION"] = merge_df.iloc[:, 34] == merge_df.iloc[:, -1]
            inconsistencies = merge_df[~merge_df["VALIDATION"]].copy()

            inconsistencies["COORDINATE_1"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(35)}{row.name + 2}",
                axis=1,
            )

            return append_inconsistencias(
                inconsistencies_file, "ValidacionTotalReserva", inconsistencies
            )
        else:
            return "Validacion realizada, no se encontraron novedades por registrar"
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
        "latest_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Historico base reparto\BASE DE REPARTO 2024 ACTUALIZADO.xlsx",
        "sheet_latest_name": "CASOS NUEVOS",
        "col_idx": "24",
        "cut_date": "30/07/2024",
    }

    print(main(params))
