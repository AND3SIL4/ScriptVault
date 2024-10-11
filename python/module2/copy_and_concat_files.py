import pandas as pd  # type: ignore
import os


def main(params: dict):
    try:
        ##Set initial variables and values
        otros_ramos_file = params.get("otros_ramos_file")
        desempleo_file = params.get("desempleo_file")
        sheet_otros_ramos = params.get("sheet_otros_ramos")
        sheet_desempleo = params.get("sheet_desempleo")
        begin_date = params.get("begin_date")
        cut_off_date = params.get("cut_off_date")
        col_idx = int(params.get("col_idx"))
        destination_path = params.get("destination_path")

        ##Validate if all inputs required are present
        if not all(
            [
                otros_ramos_file,
                desempleo_file,
                sheet_otros_ramos,
                sheet_desempleo,
                begin_date,
                cut_off_date,
                col_idx,
            ]
        ):
            return "Error: an input required is missing"

        ##Make a date type to filter the files
        begin_date = pd.to_datetime(begin_date, format="%d/%m/%Y")
        cut_off_date = pd.to_datetime(cut_off_date, format="%d/%m/%Y")

        ##Read the books and make a filter
        otros_ramos_df = pd.read_excel(
            otros_ramos_file, sheet_name=sheet_otros_ramos, engine="openpyxl"
        )
        desempleo_df = pd.read_excel(
            desempleo_file, sheet_name=sheet_desempleo, engine="openpyxl"
        )

        ##Convert the column to date type
        otros_ramos_df.iloc[:, col_idx] = pd.to_datetime(
            otros_ramos_df.iloc[:, col_idx], format="%d/%m/%Y"
        )
        desempleo_df.iloc[:, col_idx] = pd.to_datetime(
            desempleo_df.iloc[:, col_idx], format="%d/%m/%Y"
        )

        ##Make a filter
        otros_ramos_filtered = otros_ramos_df[
            (otros_ramos_df.iloc[:, col_idx] >= begin_date)
            & (otros_ramos_df.iloc[:, col_idx] <= cut_off_date)
        ]

        desempleo_filtered = desempleo_df[
            (desempleo_df.iloc[:, col_idx] >= begin_date)
            & (desempleo_df.iloc[:, col_idx] <= cut_off_date)
        ]
        ##Link the files previously filtered
        base_pagos = pd.concat(
            [otros_ramos_filtered, desempleo_filtered], ignore_index=True
        )

        ##Save changes into a temp folder
        base_pagos.to_excel(destination_path, index=False, sheet_name="PAGOS")
        return "Temp file created successfully"

    except Exception as e:
        return f"Error: {e}"
