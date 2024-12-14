import pandas as pd  # type: ignore
import os


def main(params: dict):
    try:
        ##Set variables
        file_path = params.get("file_path")
        inconsistencies_file = params.get("inconsistencies_file")
        list_file = params.get("list_file")

        ##Validate if all required input are present
        if not all([file_path, inconsistencies_file]):
            return "Error: an input params is missing"

        ##Read the work book
        book: pd.DataFrame = pd.read_excel(
            file_path, sheet_name="CASOS NUEVOS", engine="openpyxl"
        )
        listas: pd.DataFrame = pd.read_excel(
            list_file, sheet_name="LISTAS", engine="openpyxl"
        )

        ## Concepto list values allowed
        concepto_list: list[str] = listas["CONCEPTO"].dropna().astype(str).to_list()
        ## Make the validation
        book["is_valid"] = book.apply(
            lambda row: is_valid(str(row.iloc[12]), str(row.iloc[35]), concepto_list),
            axis=1,
        )

        ##Filter the inconsistencies data frame
        inconsistencies: pd.DataFrame = book[~book["is_valid"]].copy()

        ## Store and return the inconsistencies
        if inconsistencies.empty:
            return "INFO: validación de columna Concepto realizada correctamente, no se encontraron inconsistencias"
        else:
            inconsistencies["COORDENADAS"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(35 + 1)}{row.name + 2}", axis=1
            )

            ## Create a sheet name to store the inconsistencies
            new_sheet = "ValidacionConceptoColumna"

            if os.path.exists(inconsistencies_file):
                with pd.ExcelFile(inconsistencies_file, engine="openpyxl") as xls:
                    if new_sheet in xls.sheet_names:
                        existing_file = pd.read_excel(
                            xls, sheet_name=new_sheet, engine="openpyxl"
                        )
                        inconsistencies = pd.concat(
                            [existing_file, inconsistencies], ignore_index=True
                        )

            with pd.ExcelWriter(
                inconsistencies_file,
                engine="openpyxl",
                mode="a",
                if_sheet_exists="replace",
            ) as writer:
                inconsistencies.to_excel(writer, index=False, sheet_name=new_sheet)
                return "Inconsistencies registered successfully"

    except Exception as e:
        print(f"ERROR: {str(e)}")


def is_valid(ramo: str, concepto: str, lista: list[str]) -> bool:
    """Method to validate if the ramo is OTROS GASTOS and depends on that validate into a list"""
    if ramo == "OTROS GASTOS":
        return concepto in lista
    return concepto == "nan"


def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


"""Apply with a use case"""
if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
        "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
        "list_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\EXCEPCIONES BASE REPARTO.xlsx",
    }

    print(main(params))
