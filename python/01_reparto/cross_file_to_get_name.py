import pandas as pd  # type: ignore
import os


def main(params: dict) -> None:
    try:
        ##Set initial variables
        file_path: str = params.get("file_path")
        inconsistencies_file: str = params.get("inconsistencias_file")
        sheet_name: str = params.get("sheet_name")
        col_idx: int = int(params.get("col_idx"))
        list_file: str = params.get("list_file")
        sheet_list: str = params.get("sheet_list")
        col_list: int = int(params.get("col_list"))
        except_sheet_name: str = params.get("except_sheet_name")
        except_col_idx: int = int(params.get("except_col_idx"))

        ##Validate if all the required input are present
        if not all(
            [
                file_path,
                inconsistencies_file,
                sheet_name,
                list_file,
                sheet_list,
                except_sheet_name,
            ]
        ):
            return "ERROR: an input required param is missing"

        ##Read work books
        file_df: pd.DataFrame = pd.read_excel(
            file_path, sheet_name=sheet_name, engine="openpyxl"
        )
        list_df: pd.DataFrame = pd.read_excel(
            list_file, sheet_name=sheet_list, engine="openpyxl"
        )
        except_df: pd.DataFrame = pd.read_excel(
            list_file, sheet_name=except_sheet_name, engine="openpyxl"
        )

        ##Validate and cross depends on the "Poliza" number
        col_file_1_name = file_df.columns[col_idx]
        col_file_2_name = list_df.columns[col_list]

        merged_df = pd.merge(
            file_df,
            list_df,
            how="left",
            left_on=col_file_1_name,
            right_on=col_file_2_name,
            suffixes=("_OLD", "_NEW"),
        )
        ##Validate and mark inconsistencies
        merged_df["is_valid"] = merged_df["TOMADOR_OLD"] == merged_df["TOMADOR_NEW"]
        merged_df["exists_in_list"] = merged_df[col_file_1_name].isin(
            except_df.iloc[:, except_col_idx]
        )
        inconsistencies = merged_df[
            ~merged_df["is_valid"] & ~merged_df["exists_in_list"]
        ].copy()

        if not inconsistencies.empty:
            inconsistencies["COORDINATE_1"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(col_idx + 1)}{row.name + 2}",
                axis=1,
            )
            inconsistencies["COORDINATE_2"] = inconsistencies.apply(
                lambda row: f"{get_excel_column_name(17)}{row.name + 2}", axis=1
            )
            return append_inconsistencias(
                inconsistencies_file, "NombreEstandarizados", inconsistencies
            )

    except Exception as e:
        return f"ERROR: {e}"


def append_inconsistencias(file_path: str, new_sheet: str, data_frame) -> None:
    """This function get the inconsistencies data frame and append it into the inconsistencies file"""
    if os.path.exists(file_path):
        with pd.ExcelFile(file_path, engine="openpyxl") as xls:
            if new_sheet in xls.sheet_names:
                existing = pd.read_excel(xls, sheet_name=new_sheet, engine="openpyxl")
                data_frame = pd.concat([existing, data_frame], ignore_index=True)

        with pd.ExcelWriter(
            file_path, engine="openpyxl", if_sheet_exists="replace", mode="a"
        ) as writer:
            data_frame.to_excel(writer, index=False, sheet_name=new_sheet)
            return "Inconsistencias registradas correctamente"


def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
        "inconsistencias_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "col_idx": "6",
        "list_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\Listas - BOT.xlsx",
        "sheet_list": "COASEGURO",
        "col_list": "0",
        "except_sheet_name": "EXCEPCIONES NOMBRES TOMADOR",
        "except_col_idx": "0",
    }

    print(main(params))
