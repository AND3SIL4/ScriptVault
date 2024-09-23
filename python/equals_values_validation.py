import pandas as pd
import os

def main(params: dict):
  try:
    ##Set variables
    file_path = params.get("file_path")
    inconsistencias_file = params.get("inconsistencias_file")
    sheet_name = params.get("sheet_name")
    col1 = int(params.get("col1"))
    col2 = int(params.get("col2"))


    ##Validate if all the inputs required are present
    if not all([file_path, inconsistencias_file, sheet_name, col1, col2]):
      return "Error: an input param required is missing"

    ##Read File
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

    ##Filter file
    filtered_file = df.iloc[:, col1] != df.iloc[:, col2]

    ##Make a validation and save the inconsistencies
    if filtered_file.empty:
      return "Validación realizada correctamente, no hay inconsistencias para registrar"
    else:
      new_sheet_name = "ValidacionIgualdadColumnas"
      if os.path.exists(inconsistencias_file):
        with pd.ExcelFile(inconsistencias_file, engine="openpyxl") as xls:
          if new_sheet_name in xls.sheet_names:
            existing_file = pd.read_excel(xls, sheet_name=new_sheet_name, engine="openpyxl")
            filtered_file = pd.concat([existing_file, filtered_file], ignore_index=True)
        
        with pd.ExcelWriter(inconsistencias_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
          filtered_file.to_excel(writer, sheet_name=new_sheet_name, index=False)
          return "Inconsistencias registradas correctamente"

  except Exception as e:
    return f"Error: {e}"

if __name__ == "__main__":
  params  = {
    "file_path": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE PAGOS.xlsx",
    "inconsistencias_file": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBasePagos.xlsx",
    "sheet_name": "PAGOS",
    "col1": "48",
    "col2": ""
  }

  print(main(params))
