import pandas as pd  # type: ignore
from datetime import datetime
import traceback
from typing import Optional


class Tables:
    """Clase para manejar la información de coaseguros"""

    def __init__(self, file_path: str, sheet_name: str):
        self.path_file = file_path
        self.sheet_name = sheet_name

    def read_excel(self, file_path: str, sheet_name: str) -> pd.DataFrame:
        """Method for returning a data frame"""
        return pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

    def create_pivot_table(
        self, df: pd.DataFrame, value_column: str, columna: str, aggfunc: str
    ) -> pd.DataFrame:
        """Create a pivot table for the given aggregation function (sum or count)."""
        return pd.pivot_table(
            df,
            values=value_column,
            index="RAMO",
            columns=columna,
            aggfunc=aggfunc,
            fill_value=0,
        ).astype(int)

    def add_total(self, df: pd.DataFrame) -> None:
        """Validate if the current table matches the latest table."""
        current_sum = df.sum().astype(float)
        ##Add validation to show into the pivot table
        df.loc["TOTAL_ACTUAL"] = current_sum

    def sort_month_columns(self, df: pd.DataFrame) -> pd.DataFrame:
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

    def save_to_file(
        self, data_frame: pd.DataFrame, file_path: str, sheet_name: str
    ) -> None:
        """Function to save the DataFrame to an Excel file"""
        with pd.ExcelWriter(
            file_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            data_frame.to_excel(writer, sheet_name=sheet_name, index=True)
            return "Tabla guardada correctamente"

    def get_month(self, date: str) -> str:
        months = {
            1: "ENERO",
            2: "FEBRERO",
            3: "MARZO",
            4: "ABRIL",
            5: "MAYO",
            6: "JUNIO",
            7: "JULIO",
            8: "AGOSTO",
            9: "SEPTIEMBRE",
            10: "OCTUBRE",
            11: "NOVIEMBRE",
            12: "DICIEMBRE",
        }

        date_variable: pd.Timestamp = pd.to_datetime(
            date, format="%d/%m/%Y", errors="coerce"
        )
        month = months.get(date_variable.month)
        return month


##* INITIALIZE THE VARIABLE TO INSTANCE THE MAIN CLASS
tables: Optional[Tables] = None


##* CALL THE MAIN FUNCTION WITH THE MAIN PARAMS
def main(params: dict) -> bool:
    try:
        global tables

        ## Get the variables
        file_path: str = params.get("file_path")
        sheet_name: str = params.get("sheet_name")

        ## Pass the values to the constructor in the main class
        tables = Tables(file_path, sheet_name)
        return True
    except Exception as e:
        return f"ERROR: {e}"


def table_valor_movimiento():
    try:
        pagos_data_frame: pd.DataFrame = tables.read_excel(
            tables.path_file, tables.sheet_name
        )
        ## Get the month depends on the "FECHA E MAIL ENVIO FINANCIERA"
        pagos_data_frame["MES_ENVIO_FINANCIERA"] = pagos_data_frame.iloc[:, 72].apply(
            lambda date: tables.get_month(date)
        )
        ## Create pivot table for the sum of "VALOR RESERVA" by "MES_ENVIO_FINANCIERA"
        valor_reserva = tables.create_pivot_table(
            pagos_data_frame,
            pagos_data_frame.columns[45],
            "MES_ENVIO_FINANCIERA",
            "sum",
        )
        tables.add_total(valor_reserva)
        valor_reserva = tables.sort_month_columns(valor_reserva)
        tables.save_to_file(valor_reserva, tables.path_file, "VALOR MOVIMIENTO")
        return "SUCCESS: Tabla generada correctamente"
    except Exception as e:
        return f"ERROR: {e}"


def table_cantidad_registros():
    try:
        pagos_data_frame: pd.DataFrame = tables.read_excel(
            tables.path_file, tables.sheet_name
        )
        ## Get the month depends on the "FECHA E MAIL ENVIO FINANCIERA"
        pagos_data_frame["MES_ENVIO_FINANCIERA"] = pagos_data_frame.iloc[:, 72].apply(
            lambda date: tables.get_month(date)
        )
        ## Create pivot table for the sum of "VALOR RESERVA" by "MES_ENVIO_FINANCIERA"
        valor_reserva = tables.create_pivot_table(
            pagos_data_frame,
            pagos_data_frame.columns[45],
            "MES_ENVIO_FINANCIERA",
            "count",
        )
        tables.add_total(valor_reserva)
        valor_reserva = tables.sort_month_columns(valor_reserva)
        tables.save_to_file(valor_reserva, tables.path_file, "CANTIDAD REGISTROS")
        return "SUCCESS: Tabla generada correctamente"
    except Exception as e:
        return f"ERROR: {e}"


def table_valor_coaseguro_positiva():
    try:
        pagos_data_frame: pd.DataFrame = tables.read_excel(
            tables.path_file, tables.sheet_name
        )
        ## Get the month depends on the "FECHA E MAIL ENVIO FINANCIERA"
        pagos_data_frame["MES_ENVIO_FINANCIERA"] = pagos_data_frame.iloc[:, 72].apply(
            lambda date: tables.get_month(date)
        )
        ## Create pivot table for the sum of "VALOR RESERVA" by "MES_ENVIO_FINANCIERA"
        valor_reserva = pd.pivot_table(
            pagos_data_frame,
            values=[pagos_data_frame.columns[49], pagos_data_frame.columns[51]],
            index="RAMO",
            aggfunc="sum",  # Aquí forzamos 2 decimales en la suma
            fill_value=0,
        )

        tables.add_total(valor_reserva)
        tables.save_to_file(valor_reserva, tables.path_file, "VR POSITIVA-COASEGURA")
        return "SUCCESS: Tabla generada correctamente"
    except Exception as e:
        return f"ERROR: {e}"


def table_valor_positiva():
    try:
        pagos_data_frame: pd.DataFrame = tables.read_excel(
            tables.path_file, tables.sheet_name
        )

        # Get the month depends on the "FECHA E MAIL ENVIO FINANCIERA"
        pagos_data_frame["MES_ENVIO_FINANCIERA"] = pagos_data_frame.iloc[:, 72].apply(
            lambda date: tables.get_month(date)
        )

        # Create pivot table for the sum of "VALOR RESERVA" by "MES_ENVIO_FINANCIERA"
        # y aplicamos el formato decimal directamente en la agregación
        valor_reserva = pd.pivot_table(
            pagos_data_frame,
            values=pagos_data_frame.columns[49],
            columns="MES_ENVIO_FINANCIERA",
            index="RAMO",
            aggfunc="sum",
            #aggfunc=lambda x: round(sum(x), 2),  # Aquí forzamos 2 decimales en la suma
            fill_value=0,
        )
        valor_reserva = valor_reserva.round(2)
        tables.add_total(valor_reserva)

        valor_reserva = tables.sort_month_columns(valor_reserva)

        # Guardamos asegurando el formato de 2 decimales
        with pd.ExcelWriter(tables.path_file, mode="a", engine="openpyxl") as writer:
            valor_reserva.to_excel(
                writer, sheet_name="VALOR POSITIVA", float_format="%.2f"
            )

        return "SUCCESS: Tabla generada correctamente"
    except Exception as e:
        return f"ERROR: {e}"


def table_valor_coaseguradora():
    try:
        pagos_data_frame: pd.DataFrame = tables.read_excel(
            tables.path_file, tables.sheet_name
        )

        # Get the month depends on the "FECHA E MAIL ENVIO FINANCIERA"
        pagos_data_frame["MES_ENVIO_FINANCIERA"] = pagos_data_frame.iloc[:, 72].apply(
            lambda date: tables.get_month(date)
        )
        # Create pivot table for the sum of "VALOR RESERVA" by "MES_ENVIO_FINANCIERA"
        # y aplicamos el formato decimal directamente en la agregación
        valor_reserva: pd.DataFrame = pd.pivot_table(
            pagos_data_frame,
            values=pagos_data_frame.columns[51],
            columns="MES_ENVIO_FINANCIERA",
            index="RAMO",
            aggfunc="sum",  # Aquí forzamos 2 decimales en la suma
            fill_value=0,
        )
        valor_reserva = valor_reserva.round(2)
        tables.add_total(valor_reserva)

        valor_reserva = tables.sort_month_columns(valor_reserva)

        # Guardamos asegurando el formato de 2 decimales
        with pd.ExcelWriter(tables.path_file, mode="a", engine="openpyxl") as writer:
            valor_reserva.to_excel(
                writer, sheet_name="VALOR COASEGURADORA", float_format="%.2f"
            )

        return "SUCCESS: Tabla generada correctamente"
    except Exception as e:
        return f"ERROR: {e}"


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE PAGOS.xlsx",
        "sheet_name": "PAGOS",
    }
    main(params)
    ## Call the function
    print(table_valor_coaseguro_positiva())
