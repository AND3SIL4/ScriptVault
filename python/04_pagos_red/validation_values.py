import pandas as pd  # type: ignore
from typing import Optional
from datetime import datetime


class ValuesValidation:
    def __init__(
        self,
        file_path: str,
        inconsistencies_file: str,
        exception_file: str,
        sheet_name,
        file_name: str,
        previous_file: str,
    ):
        self.file_path = file_path
        self.inconsistencies_file = inconsistencies_file
        self.exception_file = exception_file
        self.sheet_name = sheet_name
        self.file_name = file_name
        self.previous_file = previous_file

    def read_excel(self, file_path: str, sheet_name: str) -> pd.DataFrame:
        """Method for returning a data frame"""
        return pd.read_excel(
            file_path, sheet_name=sheet_name, engine="openpyxl", dtype=str
        )

    def save_inconsistencies_file(self, df: pd.DataFrame, new_sheet: str) -> bool:
        """Method to save the inconsistencies data frame into the inconsistencies file"""
        try:
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
        except Exception as e:
            print(f"Error: {e}")
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
                        lambda row: f"{self.excel_col_name(i+1)}{row.name + 2}",
                        axis=1,
                    )
            self.save_inconsistencies_file(df, sheet_name)
            return "Success: Inconsistencies guardadas correctamente"
        else:
            return "Info: Validacion realizada, no se encontraron inconsistencias"

    def get_file_date(self) -> str:
        """Method to get the file date"""
        start: str = self.file_name.find("(") + 1
        end: str = self.file_name.find(")")

        # Extract the text
        if start > 0 and end > 0:
            file_date = self.file_name[start:end]
            return file_date
        else:
            return "error el obtener la fecha"


# Instance the main class
values_validation: Optional[ValuesValidation] = None


def extract_data_from_propuesta(data_frame: pd.DataFrame) -> pd.DataFrame:
    """Method to get the importan data from propuesta de pagos file and the return it into a data frame"""
    # Filter data frame by number of index only with the important columns
    # N° radicado casa matriz = 2
    # Valor movimiento 100% = 45
    data_frame = data_frame.iloc[:, [2, 45]].copy()
    # Drop rows with NaN values in the first column (N° radicado casa matriz)
    data_frame.dropna(subset=[data_frame.columns[0]], inplace=True)
    # Add additional columns
    data_frame["LLAVE RADICADO + MOVIMIENTO"] = (
        data_frame[data_frame.columns[0]] + " " + data_frame[data_frame.columns[1]]
    )
    data_frame["AÑO"] = datetime.now().year
    data_frame.iloc[:, 0] = data_frame.iloc[:, 0]
    return data_frame


def get_acm_report(acm_files: list[str]) -> pd.DataFrame:
    # Create list of data frames
    data_frames: list[pd.DataFrame] = []
    for file in acm_files:
        # Read the file
        df: pd.DataFrame = values_validation.read_excel(
            file, "FCT_RS_REPORTE_WS_AUDITORIA"
        )
        # Delete te no needed columns and rows
        df = df.iloc[3:, 1:]
        df.columns = df.iloc[0]
        df = df.iloc[1:].reset_index(drop=True)
        df = df[["id cuenta", "valor aprobado", "Valor Liquidado"]]
        # Append data frame into the main data frame
        data_frames.append(df)

    # Contact all the data frames into a new data frame and return it
    return pd.concat(data_frames, ignore_index=True)


def cross_file(propuesta_df: pd.DataFrame, acm_df: pd.DataFrame) -> pd.DataFrame:
    """Method to cross the propuesta and acm report and return the merged data frame"""
    # Merge the two data frames based on the radicado number
    merged_df = pd.merge(
        propuesta_df,
        acm_df,
        left_on=propuesta_df.columns[0],
        right_on=acm_df.columns[0],
        how="left",
        suffixes=("_PROPUESTA", "_ACM"),
    )
    # Filter file before return it
    merged_df = merged_df[
        [
            "AÑO",
            "No DE RADICADO CASA MATRIZ",
            "VR. MOVIMIENTO 100%",
            "LLAVE RADICADO + MOVIMIENTO",
            "Valor Liquidado",
            "valor aprobado",
        ]
    ]
    return merged_df


def apply_formulas(data_frame: pd.DataFrame, previous_df: pd.DataFrame) -> None:
    """Method to apply formulas to the merged data frame"""
    # Apply formulas

    return previous_df
    data_frame["Diferencia Aprobado-Liquidado"] = data_frame["Valor Liquidado"].astype(
        int
    ) - data_frame["valor aprobado"].astype(int)
    data_frame["Validacion Valores"] = data_frame.iloc[:, 2].astype(
        int
    ) == data_frame.iloc[:, 5].astype(int)
    # Validate duplicates
    data_frame["Duplicados casa matriz"] = (
        data_frame[data_frame.columns[1]]
        .duplicated(keep=False)
        .map({True: 2, False: 1})
    )
    data_frame["Duplicados llave"] = (
        data_frame[data_frame.columns[3]]
        .duplicated(keep=False)
        .map({True: 2, False: 1})
    )
    data_frame["Formato numero radicado"] = data_frame[data_frame.columns[1]].apply(
        lambda value: str(value).replace(",", "").replace(".", "").isdigit()
    )
    data_frame["Formato numero valor 100%"] = data_frame[data_frame.columns[2]].apply(
        lambda value: str(value).replace(",", "").replace(".", "").isdigit()
    )

    data_frame["Fecha propuesta"] = values_validation.get_file_date()


def validate_values(acm_files: list[str]) -> None:
    try:
        # Extract data from propuesta de pagos file
        propuesta_pago_df = values_validation.read_excel(
            values_validation.file_path, values_validation.sheet_name
        )
        # Get previous data frame
        previous_data_frame: pd.DataFrame = pd.read_excel(
            values_validation.previous_file,
            engine="pyxlsb",
            sheet_name="2024",
            dtype=str,
        )
        # Validate and extract the important data from propuesta and acm report
        propuesta_df: pd.DataFrame = extract_data_from_propuesta(propuesta_pago_df)
        acm_report: pd.DataFrame = get_acm_report(acm_files)
        merged_df: pd.DataFrame = cross_file(propuesta_df, acm_report)

        apply_formulas(merged_df, previous_data_frame)
        return merged_df
    except Exception as e:
        return False, f"Error: {e}"


# Call the main function
def main(params: dict) -> bool:
    try:
        global values_validation
        if values_validation is None:
            values_validation = ValuesValidation(
                file_path=params.get("file_path"),
                inconsistencies_file=params.get("inconsistencies_file"),
                exception_file=params.get("exception_file"),
                sheet_name=params.get("sheet_name"),
                file_name=params.get("file_name"),
                previous_file=params.get("previous_file"),
            )
            return True
    except Exception as e:
        print(f"Error: {e}")
        return False


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Input\PROPUESTA DE PAGO (23-10-2024).xlsx",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Temp\InconsistenciasBasePagosRedAsistencial.xlsx",
        "exception_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Input\EXCEPCIONES BASE PAGOS RED ASISTENCIAL.xlsx",
        "sheet_name": "Propuesta",
        "file_name": "PROPUESTA DE PAGO (23-10-2024).xlsx",
        "previous_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Input\VALIDACION PROPUESTA 23-10-2024 PLANTILLA.xlsb",
    }
    print(main(params))
    incomes = [
        r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Temp\FCT_RS_REPORTE_WS_AUDITORIA Desde el 01-01-2024 Hasta el 30-06-2024.xlsx",
        r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BasePagosRedAsistencial_SabanaPagosBasesSiniestralidad\Temp\FCT_RS_REPORTE_WS_AUDITORIA Desde el 01-07-2024 Hasta el 24-10-2024.xlsx",
    ]
    print(validate_values(incomes))
