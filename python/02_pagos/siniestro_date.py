import pandas as pd  # type: ignore
import traceback

def main(params: dict) -> str:
    try:
        # Set initial variables and values
        file_path = params.get("file_path")
        sheet_name = params.get("sheet_name")
        inconsistencies_file = params.get("inconsistencies_file")
        exception_file = params.get("exception_file")

        ## Validate if the required inputs are present
        if not all([file_path, sheet_name, inconsistencies_file, exception_file]):
            return "ERROR: Required inputs are missing"

        ## Read the file into a DataFrame
        data_frame: pd.DataFrame = pd.read_excel(
            file_path, sheet_name=sheet_name, engine="openpyxl"
        )
        exception_df: pd.DataFrame = pd.read_excel(
            exception_file, sheet_name="OTRAS EXCEPCIONES", engine="openpyxl"
        )

        ## List exception
        list_exception: list[str] = (
            exception_df.iloc[:, 6].dropna().astype(str).to_list()
        )

        ## Col indices
        aviso_siniestro_index = 21
        email_financiera_index = 72

        # Convert the columns to datetime
        data_frame.iloc[:, aviso_siniestro_index] = pd.to_datetime(
            data_frame.iloc[:, aviso_siniestro_index], errors="coerce"
        )
        data_frame.iloc[:, email_financiera_index] = pd.to_datetime(
            data_frame.iloc[:, email_financiera_index], errors="coerce"
        )

        # Calculate the difference in years
        data_frame["validate_dates"] = (
            data_frame.iloc[:, email_financiera_index]
            - data_frame.iloc[:, aviso_siniestro_index]
        ).dt.days / 365.25

        inconsistencies: pd.DataFrame = data_frame[
            (data_frame["validate_dates"] > 2)
        ].copy()

        def validate_exception(value: str) -> bool:
            return value in list_exception

        inconsistencies["validate_exception"] = inconsistencies.apply(
            lambda row: validate_exception(str(row.iloc[2])),
            axis=1,
        )

        inconsistencies = inconsistencies[~inconsistencies["validate_exception"]].copy()

        if not inconsistencies.empty:
            inconsistencies["COORDENADA_1"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(aviso_siniestro_index+1)}{row.name + 2}",
                axis=1,
            )

            inconsistencies["COORDENADA_ 2"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(email_financiera_index+1)}{row.name + 2}",
                axis=1,
            )
            new_sheet = "ValidacionAnioSiniestro"
            with pd.ExcelFile(inconsistencies_file, engine="openpyxl") as xls:
                if new_sheet in xls.sheet_names:
                    exiting_df: pd.DataFrame = pd.read_excel(
                        xls, engine="openpyxl", sheet_name=new_sheet
                    )
                    exiting_df = pd.concat(
                        [exiting_df, inconsistencies], ignore_index=True
                    )

            with pd.ExcelWriter(
                inconsistencies_file,
                engine="openpyxl",
                mode="a",
                if_sheet_exists="replace",
            ) as writer:
                inconsistencies.to_excel(writer, sheet_name=new_sheet, index=False)
                return "SUCCESS: inconsistencias registradas correctamente"
        else:
            return "INFO: validación realizada, no se encontraron inconsistencias"

        # Check if the DataFrame is empty
    except Exception as e:
        return f"ERROR: {traceback.format_exc()}"


def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name"""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE PAGOS.xlsx",
        "sheet_name": "PAGOS",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\INCONSISTENCIAS\InconBasePagos.xlsx",
        "exception_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\EXCEPCIONES BASE PAGOS.xlsx",
    }

    print(main(params))
