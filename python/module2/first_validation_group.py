import pandas as pd  # type:ignore
from typing import Optional
import os


class FirstValidationGroup:
    """Class to make the fist validation in the 'Base de Pagos' process"""

    def __init__(self, path_file: str, sheet_name: str, inconsistencies_file: str):
        self.path_file = path_file
        self.sheet_name = sheet_name
        self.inconsistencies_file = inconsistencies_file

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
        self, df: pd.DataFrame, col_idx: int, sheet_name: str
    ) -> str:
        """Method to validate the inconsistencies before append in a inconsistencies file"""
        if not df.empty:
            df = df.copy()
            df[f"COORDENADAS"] = df.apply(
                lambda row: f"{self.excel_col_name(col_idx + 1)}{row.name + 2}", axis=1
            )
            self.save_inconsistencies_file(df, sheet_name)
            return "SUCCESS: Inconsistencies guardadas correctamente"
        else:
            return "INFO: Validacion realizada, no se encontraron inconsistencias"

    def validate_empty_col(self, col_idx: int, mandatory: bool) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        if mandatory:
            data_frame["is_valid"] = data_frame.iloc[:, col_idx].isna()
            inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
            return self.validate_inconsistencies(
                inconsistencies, col_idx, "ValidacionColumnasVacias"
            )
        else:
            data_frame["is_valid"] = data_frame.iloc[:, col_idx].isna()
            inconsistencies: pd.DataFrame = data_frame[data_frame["is_valid"]]
            return self.validate_inconsistencies(
                inconsistencies, col_idx, "ValidacionColumnasVacias"
            )

    def number_type(self, col_idx: int) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = data_frame.iloc[:, col_idx].apply(
            lambda x: str(x).replace(".", "").isdigit() if pd.notna(x) else False
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, col_idx, "DatoTipoNumero")

    def date_type(self, col_idx: int) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = pd.to_datetime(
            data_frame.iloc[:, col_idx], errors="coerce"
        ).notna()
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, col_idx, "DatosTipoFecha"
        )

## Set global variables
validation_group: Optional[FirstValidationGroup] = None


def main(params: dict) -> bool:
    try:
        global validation_group

        ## Get the variables
        file_path: str = params.get("file_path")
        sheet_name: str = params.get("sheet_name")
        inconsistencies_file: str = params.get("inconsistencies_file")
        validation_group = FirstValidationGroup(
            file_path, sheet_name, inconsistencies_file
        )
        return True
    except Exception as e:
        return f"ERROR: {e}"


def validate_empty_cols(incomes: dict) -> str:
    try:
        col_idx = int(incomes.get("col_idx"))
        is_mandatory = incomes.get("is_mandatory")

        validate: str = validation_group.validate_empty_col(col_idx, is_mandatory)
        return validate
    except Exception as e:
        return f"ERROR: {e}"


def validate_number_type(params: dict) -> str:
    try:
        ## Set local variables
        index = int(params.get("col_idx"))

        validate: str = validation_group.number_type(index)
        return validate
    except Exception as e:
        return f"ERROR: {e}"
    

def validate_date_type(incomes: dict)-> str:
    try:
        ## Set local variables
        index = int(incomes.get("col_idx"))

        validate: str = validation_group.date_type(index)
        return validate
    except Exception as e:
        return f"ERROR: {e}"



if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE PAGOS.xlsx",
        "sheet_name": "PAGOS",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBasePagos.xlsx",
    }
    main(params)

    otro = {"col_idx": "1"}
    print(validate_date_type(otro))
