import pandas as pd  # type: ignore
import numpy as np  # type: ignore
from typing import Optional
import os
import re


class Consecutivo:
    """Clase para manejar la información de coaseguros"""

    def __init__(
        self,
        file_path: str,
        sheet_name: str,
        inconsistencies_file: str,
        exception_file: str,
        consecutivo_sap_file: str,
    ):
        self.path_file = file_path
        self.sheet_name = sheet_name
        self.inconsistencies_file = inconsistencies_file
        self.exception_file = exception_file
        self.consecutivo_sap_file = consecutivo_sap_file

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

    def consecutivo(self) -> str:
        ## Initial data frames
        pagos_df: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        consecutivo_df: pd.DataFrame = self.read_excel(
            self.consecutivo_sap_file, "NUEMRO DE PAGO"
        )
        list_df: pd.DataFrame = self.read_excel(self.exception_file, "CONSECUTIVO SAP")

        ## Local variables
        initial_consecutivo: int = int(list_df.iloc[0, 1])
        final_consecutivo: int = int(list_df.iloc[0, 2])
        pending_list: list[int] = list_df.iloc[:, 0].dropna().astype(int).to_list()

        consecutivo_pagos: pd.Series = pagos_df.iloc[:, 73]
        length_consecutivo_pagos: int = int(len(consecutivo_pagos))

        return length_consecutivo_pagos


##* INITIALIZE THE VARIABLE TO INSTANCE THE MAIN CLASS
consecutivo: Optional[Consecutivo] = None


##* CALL THE MAIN FUNCTION WITH THE MAIN PARAMS
def main(params: dict) -> bool:
    try:
        global consecutivo

        ## Get the variables
        file_path: str = params.get("file_path")
        sheet_name: str = params.get("sheet_name")
        inconsistencies_file: str = params.get("inconsistencies_file")
        exception_file: str = params.get("exception_file")
        consecutivo_sap_file: str = params.get("consecutivo_sap_file")

        ## Pass the values to the constructor in the main class
        consecutivo = Consecutivo(
            file_path,
            sheet_name,
            inconsistencies_file,
            exception_file,
            consecutivo_sap_file,
        )
        return True
    except Exception as e:
        return f"ERROR: {e}"


def validate_consecutivo_sap(params: dict) -> str:
    try:
        ## Set local variables
        cut_off_date: str = params.get("cut_off_date")
        validation: str = consecutivo.consecutivo()
        return validation
    except Exception as e:
        return f"ERROR: {e}"


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE PAGOS.xlsx",
        "sheet_name": "PAGOS",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBasePagos.xlsx",
        "exception_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\EXCEPCIONES BASE PAGOS.xlsx",
        "consecutivo_sap_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\CONSECUTIVO SAP 2023.xlsx",
    }
    main(params)
    params = {"cut_off_date": "30/07/2024"}
    print(validate_consecutivo_sap(params))
