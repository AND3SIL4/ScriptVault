import pandas as pd  # type: ignore
import numpy as np  # type: ignore
from typing import Optional
import os


class Coaseguro:
    """Clase para manejar la información de coaseguros"""

    def __init__(
        self,
        file_path: str,
        sheet_name: str,
        inconsistencies_file: str,
        exception_file: str,
    ):
        self.path_file = file_path
        self.sheet_name = sheet_name
        self.inconsistencies_file = inconsistencies_file
        self.exception_file = exception_file

    def read_excel(self, file_path: str, sheet_name: str) -> pd.DataFrame:
        """Method for returning a data frame"""
        return pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

    def save_inconsistencies_file(self, df: pd.DataFrame, new_sheet: str) -> bool:
        if os.path.exists(self.inconsistencies_file):
            with pd.ExcelFile(self.inconsistencies_file, engine="openpyxl") as xls:
                if new_sheet in xls.sheet_names:
                    existing = pd.read_excel(
                        xls, engine="openpyxl", sheet_name=new_sheet
                    )
                    df = pd.concat([existing, df], ignore_index=True)

            with pd.ExcelWriter(
                self.inconsistencies_file,
                engine="openpyxl",
                mode="a",
                if_sheet_exists="replace",
            ) as writer:
                df.to_excel(writer, sheet_name=new_sheet, index=False)
                return True
        else:
            return False

    def excel_col_name(self, number) -> str:
        """Method to convert (1-based) to Excel column name"""
        result = ""
        while number > 0:
            number, reminder = divmod(number - 1, 26)
            result = chr(65 + reminder) + result
        return result

    def validate_inconsistencies(
        self, df: pd.DataFrame, col_idx, sheet_name: str
    ) -> str:
        """Method to validate the inconsistencies before append in a inconsistencies file"""
        if not df.empty:
            df = df.copy()
            if isinstance(col_idx, int):
                df[f"COORDENADAS"] = df.apply(
                    lambda row: f"{self.excel_col_name(col_idx + 1)}{row.name + 2}",
                    axis=1,
                )
            else:
                for i in col_idx:
                    df[f"COORDENADAS_{i + 2}"] = df.apply(
                        lambda row: f"{self.excel_col_name(i + 1)}{row.name + 2}",
                        axis=1,
                    )
            self.save_inconsistencies_file(df, sheet_name)
            return "SUCCESS: Inconsistencies guardadas correctamente"
        else:
            return "INFO: Validacion realizada, no se encontraron inconsistencias"

    def is_coaseguro(self):
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        lista: list[str] = ["PREVISORA; MUNDIAL", "GENERAL", "MUNDIAL", "PREVISORA"]

        ## Sub function to validate if is coaseguro
        def is_coaseguro_helper(coaseguro: str, porcentaje_positiva: str) -> bool:
            if porcentaje_positiva != "1":
                return coaseguro in lista
            else:
                return coaseguro == "nan"

        data_frame["is_valid"] = data_frame.apply(
            lambda row: is_coaseguro_helper(str(row.iloc[47]), str(row.iloc[48])),
            axis=1,
        )

        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, 47, "ValidacionCoaseguro")

    def data_from_coaseguro_sheet(self) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        coaseguro_df: pd.DataFrame = self.read_excel(self.exception_file, "COASEGURO")
        coaseguro_df = coaseguro_df.iloc[:, 0:6]
        ## Merge data
        name_col_pagos: str = data_frame.columns[6]
        name_col_coaseguro: str = coaseguro_df.columns[0]
        merged_df: pd.DataFrame = data_frame.merge(
            coaseguro_df,
            left_on=name_col_pagos,
            right_on=name_col_coaseguro,
            how="left",
            suffixes=("_PAGOS", "_COASEGURO"),
        )

        def validate_equals(pagos: str, coaseguro: str) -> bool:
            return pagos == coaseguro

        merged_df["is_valid"] = merged_df.apply(
            lambda row: validate_equals(str(row.iloc[48]), str(row.iloc[114])),
            axis=1,
        )
        inconsistencies: pd.DataFrame = merged_df[~merged_df["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, 48, "ValidacionValorPositiva"
        )

    def positiva_calculados(self) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        ## Columns
        vr_movimiento: pd.Series = data_frame.iloc[:, 45]
        porcentaje_positiva: pd.Series = data_frame.iloc[:, 48]
        vr_positiva: pd.Series = data_frame.iloc[:, 49].astype(float).round(2)

        data_frame["POSITIVA_CALCULADOS"] = (
            (vr_movimiento * porcentaje_positiva).astype(float).round(2)
        )
        data_frame["VALIDACION"] = data_frame["POSITIVA_CALCULADOS"] == vr_positiva
        inconsistencies: pd.DataFrame = data_frame[~data_frame["VALIDACION"]]
        ## Save inconsistencies into file
        return self.validate_inconsistencies(
            inconsistencies, [49, 111], "ValidacionPositivaCalculados"
        )

    def coasegura_calculado(self) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        vr_100: pd.Series = data_frame.iloc[:, 45].astype(float).round(2)
        porcentaje_coaseguradora: pd.Series = data_frame.iloc[:, 50]
        vr_coaseguradora: pd.Series = data_frame.iloc[:, 51]

        ## Subfunction to validate and fix the coaseguradora percentage
        def fix_coaseguradora_percentage(value: str) -> float:
            try:
                # Eliminar caracteres innecesarios
                value = value.replace("%", "").replace(" ", "")
                # Si el valor contiene un ';', procesarlo como suma
                if ";" in value:
                    values = value.split(";")
                    # Convertir los dos valores y calcular el resultado
                    total_percentage = int(values[0]) + int(values[1])
                    result = total_percentage / 100.0
                else:
                    # Si no contiene ';', convertir directamente a flotante
                    result = float(value)
                return result
            except (ValueError, IndexError) as e:
                # Si hay algún error en la conversión o el formato, manejarlo aquí
                print(f"Error al procesar el valor '{value}': {e}")
                return 0.0  # Devolver un valor por defecto en caso de error

        data_frame = data_frame[data_frame.iloc[:, 42] == "COASEGURO"]

        data_frame["PORCENTAJE_COASEGURADORA"] = porcentaje_coaseguradora.astype(
            str
        ).apply(lambda value: fix_coaseguradora_percentage(value))

        data_frame["COASEGURADORA_CALCULADO"] = (
            data_frame["PORCENTAJE_COASEGURADORA"] * vr_100
        )

        ## Sub function to validate the belonging
        def validate_belonging(vr_coaseguro: float, coaseguro_calculado: float) -> bool:
            return round(vr_coaseguro, 2) == round(coaseguro_calculado, 2)

        data_frame["VR_COASEGURO_VS_COASEGURO_CALCULADO"] = data_frame.apply(
            lambda row: validate_belonging(float(row.iloc[51]), float(row.iloc[112])),
            axis=1,
        )
        inconsistencies: pd.DataFrame = data_frame[
            ~data_frame["VR_COASEGURO_VS_COASEGURO_CALCULADO"]
        ]
        return self.validate_inconsistencies(
            inconsistencies, [51, 112], "ValidacionCoaseguroCalculado"
        )


##* INITIALIZE THE VARIABLE TO INSTANCE THE MAIN CLASS
coaseguro: Optional[Coaseguro] = None


##* CALL THE MAIN FUNCTION WITH THE MAIN PARAMS
def main(params: dict) -> bool:
    try:
        global coaseguro

        ## Get the variables
        file_path: str = params.get("file_path")
        sheet_name: str = params.get("sheet_name")
        inconsistencies_file: str = params.get("inconsistencies_file")
        exception_file: str = params.get("exception_file")

        ## Pass the values to the constructor in the main class
        coaseguro = Coaseguro(
            file_path, sheet_name, inconsistencies_file, exception_file
        )
        return True
    except Exception as e:
        return f"ERROR: {e}"


def validate_coaseguro_percentage() -> str:
    try:
        ## Set local variables
        validation: str = coaseguro.is_coaseguro()
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_data_from_coaseguro() -> str:
    try:
        ## Set local variables
        validation: str = coaseguro.data_from_coaseguro_sheet()
        return "SUCCESS" in validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_positiva_calculado() -> str:
    try:
        ## Set local variables
        validation: str = coaseguro.positiva_calculados()
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_coasegura_calculado() -> str:
    try:
        ## Set local variables
        validation: str = coaseguro.coasegura_calculado()
        return validation
    except Exception as e:
        return f"ERROR: {e}"


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE PAGOS.xlsx",
        "sheet_name": "PAGOS",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBasePagos.xlsx",
        "exception_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\EXCEPCIONES BASE PAGOS.xlsx",
    }
    main(params)
    params = {
        "col_idx": "100",
        "option": "REACTIVADO",
        "new_sheet": "ValidacionReactivado",
    }
    print(validate_coasegura_calculado())
