import pandas as pd
import os

def main(params):
    try:
        ##Set the variables
        file_path = params.get("file_path")
        sheet_name = params.get("sheet_name")
        firstColumn = params.get("col1")
        secondColumn = params.get("col2")
        inconsistencies_file = params.get("in_file")
        dic = {}
        inconsistencies = []

        ##Read the book and load it
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

        ##Make validation and return the inconsistencies
        for index, row in df.iterrows():
            col1 = row[firstColumn]
            col2 = row[secondColumn]

            ##If column 1 or column 2 are null or NaN only continue
            if pd.isna(col1) or pd.isna(col2):
                continue

            ##If column 1 is in dictionary make a validation
            if col1 in dic:
                ##Validate if column 2 is not into the dictionary, report to inconsistencies
                if dic[col1] != col2:
                    inconsistencies.append(
                        index + 2
                    )  ##Append the new inconsistency to the file
                    ##Report the inconsistency to the inconsistencies file
                    filtered = df[df[secondColumn] == col2]
                    new_sheet = "ValidacionUnSoloPerteneciente"
                    if os.path.exists(inconsistencies_file):
                        with pd.ExcelFile(
                            inconsistencies_file, engine="openpyxl"
                        ) as xls:
                            if new_sheet in xls.sheet_names:
                                existing = pd.read_excel(
                                    xls, sheet_name=new_sheet, engine="openpyxl"
                                )
                                filtered = pd.concat(
                                    [existing, filtered], ignore_index=True
                                )

                    with pd.ExcelWriter(
                        inconsistencies_file,
                        engine="openpyxl",
                        mode="a",
                        if_sheet_exists="replace",
                    ) as writer:
                        filtered.to_excel(writer, sheet_name=new_sheet, index=False)
                        print("Inconsitencias registradas correctamente")
            else:
                ##If there is no a value for column 2, put into the dictionary
                dic[col1] = col2

        ##Return the inconsistencies
        return inconsistencies

    except Exception as e:
        return f"Error en la ejecución: {e}"


if __name__ == "__main__":
    params = {
        "file_path": "C:/ProgramData/AutomationAnywhere/Bots/Logs/AD_RCSN_SabanaPagosYBasesParaSinestralidad/TempFolder/BASE DE REPARTO 2024.xlsx",
        "sheet_name": "CASOS NUEVOS",
        "col1": "No. SINIESTRO",
        "col2": """No POLIZA""",
        "in_file": "C:/ProgramData/AutomationAnywhere/Bots/Logs/AD_RCSN_SabanaPagosYBasesParaSinestralidad/OutputFolder/Inconsistencias/InconBaseReparto.xlsx",
    }

    print(main(params))
