import pandas as pd  # type: ignore
from datetime import datetime
import re
import traceback
import os


def main(params: dict):
    try:
        ##Set initial variables
        file_path: str = params.get("file_path")
        sheet_name: str = params.get("sheet_name")
        latest_file: str = params.get("latest_file")
        col_idx: int = int(params.get("col_idx"))
        cut_off_date: str = params.get("cut_off_date")
        inconsistencias_file: str = params.get("inconsistencias_file")

        ##Validate if all the required inputs are present
        if not all([file_path, sheet_name, latest_file, inconsistencias_file]):
            return "ERROR: an required input is missing"

        year = datetime.today().year
        cut_date = pd.to_datetime(cut_off_date, format="%d/%m/%Y")
        date = f"01/01/{year}"
        initial_date = pd.to_datetime(date, format="%d/%m/%Y")

        current_df: pd.DataFrame = pd.read_excel(
            file_path, sheet_name=sheet_name, engine="openpyxl"
        )
        current_filtered: pd.DataFrame = current_df[
            (current_df.iloc[:, col_idx] > initial_date)
            & (current_df.iloc[:, col_idx] < cut_date)
        ].copy()

        latest_df: pd.DataFrame = pd.read_excel(
            latest_file, sheet_name=sheet_name, engine="openpyxl"
        )
        latest_filtered: pd.DataFrame = latest_df[
            (latest_df.iloc[:, col_idx] > initial_date)
            & (latest_df.iloc[:, col_idx] < cut_date)
        ].copy()
        ##Fix white spaces
        current_filtered["MES DE ASIGNACION"] = (
            current_filtered["MES DE ASIGNACION"]
            .astype(str)
            .apply(lambda x: clean_white_spaces(x))
        )
        latest_filtered["MES DE ASIGNACION"] = (
            latest_filtered["MES DE ASIGNACION"]
            .astype(str)
            .apply(lambda x: clean_white_spaces(x))
        )

        ##Make tables
        current_table = pd.pivot_table(
            current_filtered,
            values="VALOR RESERVA",
            index="RAMO",
            columns="MES DE ASIGNACION",
            aggfunc="sum",
            fill_value=0,
        )
        latest_table = pd.pivot_table(
            latest_filtered,
            values="VALOR RESERVA",
            index="RAMO",
            columns="MES DE ASIGNACION",
            aggfunc="sum",
            fill_value=0,
        )
        ##Convert values to integer
        current_table = current_table.astype(int)
        latest_table = latest_table.astype(int)

        # Sum the values of both tables
        latest_sum = latest_table.sum().astype(int)
        current_sum = current_table.sum().astype(int)

        if current_filtered.empty:
            return "ERROR: No hay datos después del filtrado en el archivo actual"

        if latest_filtered.empty:
            return "ERROR: No hay datos después del filtrado en el archivo más reciente"

        # Después de crear current_table
        if current_table.empty:
            return "ERROR: La tabla dinámica actual está vacía"
        
        # Add totals to both tables
        current_table.loc["TOTAL_ACTUAL"] = current_sum
        current_table.loc["TOTAL_ANTERIOR"] = latest_sum
        current_table.loc["VALIDACION"] = latest_sum == current_sum
        current_table.loc["VALIDACION"] = current_table.loc["VALIDACION"].astype(bool)
        # Validation of totals
        is_valid = current_sum.equals(latest_sum)
        if is_valid:
            current_df["MES DE ASIGNACION"] = (
                current_df["MES DE ASIGNACION"]
                .astype(str)
                .apply(lambda x: clean_white_spaces(x))
            )
            table = pd.pivot_table(
                current_df,
                values="VALOR RESERVA",
                index="RAMO",
                columns="MES DE ASIGNACION",
                aggfunc="sum",
                fill_value=0,
            )
            table = table.astype(int)
            # Sum the values of both tables
            latest_sum = latest_table.sum().astype(int)
            col_sum = table.sum().astype(int)

            # Add totals to both tables
            table.loc["TOTAL_ACTUAL"] = col_sum
            table.loc["TOTAL_ANTERIOR"] = latest_sum
            sorted_df = sort_month_columns(table)
            # If totals match, save the current table to the file
            save_to_file(sorted_df, file_path, "REPORTE")
            return True
        else:
            fixed_df = sort_month_columns(current_table)
            # If totals do not match, report inconsistencies
            append_inconsistencias(
                inconsistencias_file, "ValorReservaEjecucionAnterior", fixed_df
            )
            return False
    except Exception as e:
        return f"ERROR: {e} {traceback.format_exc()}"


def sort_month_columns(df: pd.DataFrame) -> pd.DataFrame:
    months = [
        "ENERO",
        "FEBRERO",
        "MARZO",
        "ABRIL",
        "MAYO",
        "JUNIO",
        "JULIO",
        "AGOSTO",
        "SEPTIEMBRE",
        "OCTUBRE",
        "NOVIEMBRE",
        "DICIEMBRE",
    ]
    columns_present = [mes for mes in months if mes in df.columns]
    sorted_df = df[columns_present]
    return sorted_df


def clean_white_spaces(string: str):
    value = re.sub(r"[\s]", "", string)
    return value


def save_to_file(data_frame: pd.DataFrame, file_path: str, sheet_name: str) -> None:
    """Function to save the DataFrame to an Excel file"""
    with pd.ExcelWriter(
        file_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        data_frame.to_excel(writer, sheet_name=sheet_name, index=True)
        return "Tabla guardada correctamente"


def append_inconsistencias(file_path: str, new_sheet: str, data_frame) -> None:
    """This function get the inconsistencies data frame and append it into the inconsistencies file"""
    if os.path.exists(file_path):
        with pd.ExcelFile(file_path, engine="openpyxl") as xls:
            if new_sheet in xls.sheet_names:
                existing = pd.read_excel(xls, sheet_name=new_sheet, engine="openpyxl")
                data_frame = pd.concat([existing, data_frame], ignore_index=False)

        with pd.ExcelWriter(
            file_path, engine="openpyxl", if_sheet_exists="replace", mode="a"
        ) as writer:
            data_frame.to_excel(writer, index=True, sheet_name=new_sheet)
            return "Inconsistencias registradas correctamente"
