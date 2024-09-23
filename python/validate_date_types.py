import pandas as pd
import os
from datetime import datetime

def main(params: dict):
  try:
    ##Set variables
    file_path = params.get("file_path")
    col_idx = int(params.get("col_idx"))
    inconsistencias_file = params.get("inconsistencias_file")
    sheet_name = params.get("sheet_name")
    can_be_null = bool(params.get("can_be_null"))

    if not all([file_path, col_idx, inconsistencias_file, sheet_name]):
      return "Error: input param is missing"

    ##Read data base
    df: pd.DataFrame = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

    ## Create a new column to validate dates, considering null
    df["valid_date"] = df.iloc[:, col_idx].astype(str).apply(lambda x: is_valid(x, can_be_null))

    ##Apply filter
    filtered_file = df[~df["valid_date"]].copy()

    print(filtered_file)

    if (filtered_file.empty):
      return "Validación correcta, no hay inconsistencias"
    else:
      filtered_file['COORDENADAS'] = filtered_file.apply(
      lambda row: f"{get_excel_column_name(col_idx + 1)}{row.name + 2}", axis=1
      )
      new_sheet_name = "ValidacionTipoFecha"
      if os.path.exists(inconsistencias_file):
        with pd.ExcelFile(inconsistencias_file, engine="openpyxl") as xls:
          if new_sheet_name in xls.sheet_names:
            existing_df = pd.read_excel(xls, sheet_name=new_sheet_name, engine="openpyxl")
            filtered_file = pd.concat([existing_df, filtered_file], ignore_index=True)
      
      with pd.ExcelWriter(inconsistencias_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        filtered_file.to_excel(writer, sheet_name=new_sheet_name, index=False)
        return "Inconsistencias registradas correctamente"

  except Exception as e:
    return f"Error: {e}"
  
def is_valid(value: str, can_be_null: bool):
    """Check if the value is a valid date. If nulls are allowed, treat NaN as valid."""
    if can_be_null:
       ##Can be null value
       if value.lower() == "nan" or value.strip() == "":
        return True
       else:
          try:
            datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
            return True
          except (ValueError, TypeError):
            return False
    else:
      try:
        ##Can be null value
        if value.lower() == "nan" or value.strip() == "":
          return False
        datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
        return True
      except (ValueError, TypeError):
        return False
      
def is_valid_date(date_string):
    """
    Intenta convertir una cadena a una fecha.
    Retorna True si es una fecha válida, False en caso contrario.
    """
    date_formats = [
        "%Y-%m-%d",  # 2023-05-15
        "%d/%m/%Y",  # 15/05/2023
        "%d-%m-%Y",  # 15-05-2023
        "%d.%m.%Y",  # 15.05.2023
        "%d/%m/%y",  # 15/05/23
        "%Y%m%d",    # 20230515
        # Agrega más formatos según sea necesario
    ]
    
    for date_format in date_formats:
        try:
            datetime.strptime(str(date_string), date_format)
            return True
        except ValueError:
            continue
    return False

  
def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ''
    while n > 0:
        n, remainder = divmod(n-1, 26)
        result = chr(65 + remainder) + result
    return result

if __name__ == "__main__":
  params = {
    "file_path": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
    "col_idx": "21",
    "inconsistencias_file": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
    "sheet_name": "CASOS NUEVOS",
    "can_be_null": False
  }

  print(main(params))