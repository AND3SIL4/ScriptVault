import pandas as pd  # type: ignore


def main(params: dict):
    try:
        ## Set initial variables and values
        path_file: str = params.get("path_file")
        sheet_name: str = params.get("sheet_name")
        temp_file: str = params.get("temp_file")
        col_idx: int = int(params.get("col_idx"))
        begin_date: str = params.get("begin_date")
        cut_off_date: str = params.get("cut_off_date")

        ## Validate if all inputs required are present
        if not all([path_file, sheet_name, temp_file, begin_date, cut_off_date]):
            return "ERROR: an input required param is missing"

        ## Make a date type to filter the files
        begin_date = pd.to_datetime(begin_date, format="%d/%m/%Y")
        cut_off_date = pd.to_datetime(cut_off_date, format="%d/%m/%Y")

        ##Read the books and make a filter
        objetados_df: pd.DataFrame = pd.read_excel(
            path_file, sheet_name=sheet_name, engine="openpyxl"
        ).iloc[:, :111]

        ## Convert the column to date type
        objetados_df.iloc[:, col_idx] = pd.to_datetime(
            objetados_df.iloc[:, col_idx], format="%d/%m/%Y"
        )

        ## Make a filter
        objetados_df = objetados_df[
            (objetados_df.iloc[:, col_idx] >= begin_date)
            & (objetados_df.iloc[:, col_idx] <= cut_off_date)
        ]

        ## Save changes into a temp folder
        objetados_df.to_excel(temp_file, index=False, sheet_name=sheet_name)
        return True, "Temp file created successfully"

    except Exception as e:
        return False, f"Error: {e}"


if __name__ == "__main__":
    params = {
        "path_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BaseObjetados_SabanaPagosBasesSiniestralidad\Input\Objetados 2022 - 2023 - 2024.xlsx",
        "sheet_name": "Objeciones 2022 - 2023 -2024",
        "temp_file": r"C:\ProgramData\AutomationAnywhere\Bots\AD_GI_BaseObjetados_SabanaPagosBasesSiniestralidad\Temp\Objetados.xlsx",
        "col_idx": "44",
        "begin_date": "01/01/2024",
        "cut_off_date": "29/06/2024",
    }
    print(main(params))
