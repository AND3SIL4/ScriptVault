import pandas as pd  # type:ignore
from typing import Optional
import os
import re


class FirstValidationGroup:
    """Class to make the fist validation in the 'Base de Pagos' process"""

    def __init__(
        self,
        path_file: str,
        sheet_name: str,
        inconsistencies_file: str,
        exception_file: str,
    ):
        self.path_file = path_file
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
        return self.validate_inconsistencies(inconsistencies, col_idx, "DatosTipoFecha")

    def value_length(self, col_idx: int, length: int) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = data_frame.iloc[:, col_idx].apply(
            lambda value: len(str(value)) == length
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, col_idx, "LongitudValor")

    def validate_exception_list(
        self,
        col_idx: int,
        exception_col_name: int,
        exception_sheet: str,
        new_sheet: str,
    ) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        exception_data_frame: pd.DataFrame = self.read_excel(
            self.exception_file, exception_sheet
        )
        col_exception: pd.Series = exception_data_frame[exception_col_name].dropna()
        data_frame["is_valid"] = data_frame.iloc[:, col_idx].isin(col_exception)
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, col_idx, new_sheet)

    def no_special_characters(self, col_idx: int) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = (
            data_frame.iloc[:, col_idx]
            .astype(str)
            .apply(
                lambda value: not pd.isna(value)
                and bool(re.search(r"[^a-zA-Z0-9]", value))
            )
        )

        inconsistencies: pd.DataFrame = data_frame[data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, col_idx, "ValidacionCaracteresEspaciales"
        )

    def month_depends_on_date(self, date_idx: int, month_idx: int) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)

        ## Create s sub function to know the correct month depends on the number
        months: dict = {
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

        ## Create a sub function to validate the consistency of the date
        def validate_consistency(date: str, month: str) -> bool:
            date_parse = pd.to_datetime(date, format="%Y-%m-%d", errors="coerce")
            get_month = date_parse.month
            ## Call the dictionary
            standard_month = months.get(get_month)
            return month == standard_month

        data_frame["is_valid"] = data_frame.apply(
            lambda row: validate_consistency(
                row.iloc[date_idx],
                str(row.iloc[month_idx]),
            ),
            axis=1,
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, month_idx, "ValidacionMesCorte"
        )

    def radicado_format(self, col_idx) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = data_frame.iloc[:, col_idx].apply(
            lambda value: bool(re.search(r"^\d{4}\s\d{2}\s\d{3}\s\d{6}$", str(value)))
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, col_idx, "FormatoNumeroRadicado"
        )

    def acuerdo_range(self, col_idx: int) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = (
            data_frame.iloc[:, col_idx]
            .astype(int)
            .apply(lambda value: value >= 1 and value <= 30)
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, col_idx, "ValidacionAcuerdo"
        )

    def coaseguradora(
        self, file_idx: int, exception_sheet: str, exception_col: str
    ) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        exception_df: pd.DataFrame = self.read_excel(
            self.exception_file, exception_sheet
        )
        exception_col: pd.Series = exception_df[exception_col].dropna()
        file_col: pd.Series = data_frame.iloc[:, file_idx]
        data_frame["is_valid"] = (file_col.isin(exception_col)) | (pd.isna(file_col))

        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(
            inconsistencies, file_idx, "CompañiaCoaseguradora"
        )

    def only_two_options(self, col_idx: int, options: list[str], new_sheet: str) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["id_valid"] = data_frame.iloc[:, col_idx].apply(
            lambda value: (value in options) or (pd.isna(value))
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["id_valid"]]
        return self.validate_inconsistencies(inconsistencies, col_idx, new_sheet)

    def no_white_spaces(self, col_idx: int, new_sheet: str) -> str:
        data_frame: pd.DataFrame = self.read_excel(self.path_file, self.sheet_name)
        data_frame["is_valid"] = (
            data_frame.iloc[:, col_idx]
            .astype(str)
            .apply(
                lambda value: (pd.isna(value)) or not (bool(re.search(r"\s\s+", value)))
            )
        )
        inconsistencies: pd.DataFrame = data_frame[~data_frame["is_valid"]]
        return self.validate_inconsistencies(inconsistencies, col_idx, new_sheet)


## Set global variables
validation_group: Optional[FirstValidationGroup] = None


def main(params: dict) -> bool:
    try:
        global validation_group

        ## Get the variables
        file_path: str = params.get("file_path")
        sheet_name: str = params.get("sheet_name")
        inconsistencies_file: str = params.get("inconsistencies_file")
        exception_file: str = params.get("exception_file")

        ## Pass the values to the constructor in the main class
        validation_group = FirstValidationGroup(
            file_path, sheet_name, inconsistencies_file, exception_file
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


def validate_date_type(incomes: dict) -> str:
    try:
        ## Set local variables
        index = int(incomes.get("col_idx"))

        validate: str = validation_group.date_type(index)
        return validate
    except Exception as e:
        return f"ERROR: {e}"


def validate_length(incomes: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(incomes.get("col_idx"))
        length = int(incomes.get("length"))

        validation: str = validation_group.value_length(col_idx, length)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_exception_list(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        exception_col_name = params.get("exception_col_name")
        exception_sheet = params.get("exception_sheet")
        new_sheet = params.get("new_sheet")

        validate: str = validation_group.validate_exception_list(
            col_idx, exception_col_name, exception_sheet, new_sheet
        )
        return validate
    except Exception as e:
        return f"ERROR: {e}"


def validate_special_characters(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        validation: str = validation_group.no_special_characters(col_idx)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_month(params: dict) -> str:
    try:
        ## Set local variables
        date_idx = int(params.get("date_idx"))
        month_idx = int(params.get("month_idx"))

        validation: str = validation_group.month_depends_on_date(date_idx, month_idx)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_numero_radicado(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        validation: str = validation_group.radicado_format(col_idx)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_acuerdo_range(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        validation: str = validation_group.acuerdo_range(col_idx)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_compania_coaseguradora(params: dict) -> str:
    try:
        ## Set local variables
        file_idx = int(params.get("file_idx"))
        exception_sheet = params.get("exception_sheet")
        exception_col = params.get("exception_col")

        validation: str = validation_group.coaseguradora(
            file_idx, exception_sheet, exception_col
        )
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_only_two_options(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        options = params.get("options")
        new_sheet = params.get("new_sheet")
        validation: str = validation_group.only_two_options(col_idx, options, new_sheet)
        return validation
    except Exception as e:
        return f"ERROR: {e}"


def validate_no_white_spaces(params: dict) -> str:
    try:
        ## Set local variables
        col_idx = int(params.get("col_idx"))
        new_sheet = params.get("new_sheet")

        validation: str = validation_group.no_white_spaces(col_idx, new_sheet)
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
        "col_idx": "69",
        "new_sheet": "CodigoOficinaBanco",
    }
    print(validate_no_white_spaces(params))
