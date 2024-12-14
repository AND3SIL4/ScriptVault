import pandas as pd  # type: ignore


def main(params: dict) -> str:
    try:
        # Set initial variables and values
        file_path = params.get("file_path")
        sheet_name = params.get("sheet_name")
        inconsistencies_file = params.get("inconsistencies_file")
        exception_file = params.get("exception_file")

        # Validate if all the values required are present
        if not all([file_path, sheet_name, inconsistencies_file, exception_file]):
            raise Exception("Required inputs are missing")

        # Read the file into a DataFrame
        data_frame: pd.DataFrame = pd.read_excel(
            file_path, sheet_name=sheet_name, engine="openpyxl"
        )
        exception_df: pd.DataFrame = pd.read_excel(
            exception_file, sheet_name="PRESCRIPCION", engine="openpyxl"
        )

        # Set the exception list
        list_exception: list[str] = (
            exception_df.iloc[:, 0].dropna().astype(str).to_list()
        )

        # Get only "CONCEPTO" values that are equal to "PRESCRIPCIÓN"
        data_frame = data_frame[data_frame.iloc[:, 35] == "PRESCRIPCION"]

        # Convert date columns to datetime
        data_frame.iloc[:, 44] = pd.to_datetime(data_frame.iloc[:, 44], errors="coerce")
        data_frame.iloc[:, 27] = pd.to_datetime(data_frame.iloc[:, 27], errors="coerce")

        # Create rule for cases where the "PRESCRIPCION" date are less than 2 years
        # FECHA MOVIMIENTO: 44
        # FECHA SINIESTRO: 27
        data_frame["validate_dates"] = (
            data_frame.iloc[:, 44] - data_frame.iloc[:, 27]
        ).dt.days / 365.25

        # Filter data by result less than 2
        inconsistencies: pd.DataFrame = data_frame[
            (data_frame["validate_dates"] < 2)
        ].copy()

        # Validate exceptions by "N° RADICADO"
        inconsistencies["validate_exception"] = inconsistencies.apply(
            lambda row: row.iloc[2] in list_exception,
            axis=1,
        )

        # Overwrite the inconsistencies with the values that are not present in exception list
        inconsistencies = inconsistencies[~inconsistencies["validate_exception"]].copy()

        if not inconsistencies.empty:
            # Write the coordinates
            inconsistencies["COORDENADA_1"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(28)}{row.name + 2}",
                axis=1,
            )
            inconsistencies["COORDENADA_2"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(46)}{row.name + 2}", axis=1
            )
            # Append inconsistencies
            is_append_inconsistencies: bool = append_inconsistencias(
                inconsistencies_file, "FechaPrescripción", inconsistencies
            )
            if not is_append_inconsistencies:
                raise Exception("Fail to save the inconsistencies data frame")
            return "SUCCESS: inconsistencias registradas correctamente"
        else:
            return "INFO: validation realizada, no se encontraron inconsistencias"

    except Exception as e:
        return f"ERROR: {str(e)}"


def append_inconsistencias(file_path: str, new_sheet: str, data_frame) -> None:
    """This function get the inconsistencies data frame and append it into the inconsistencies file"""
    try:
        with pd.ExcelFile(file_path, engine="openpyxl") as xls:
            if new_sheet in xls.sheet_names:
                existing = pd.read_excel(xls, sheet_name=new_sheet, engine="openpyxl")
                data_frame = pd.concat([existing, data_frame], ignore_index=True)

        with pd.ExcelWriter(
            file_path, engine="openpyxl", if_sheet_exists="replace", mode="a"
        ) as writer:
            data_frame.to_excel(writer, index=False, sheet_name=new_sheet)
        return True
    except Exception as e:
        print(f"ERROR: {e}")
        return False


def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BaseObjetados_SabanaPagosBasesSiniestralidad\Temp\Objetados.xlsx",
        "sheet_name": "Objeciones 2022 - 2023 -2024",
        "exception_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BaseObjetados_SabanaPagosBasesSiniestralidad\Input\EXCEPCIONES BASE OBJETADOS.xlsx",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BaseObjetados_SabanaPagosBasesSiniestralidad\Output\Inconsistencias\InconsistenciasBaseObjetados.xlsx",
    }
    process = main(params)
    print(process)
