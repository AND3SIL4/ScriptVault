import pandas as pd
import os

def main(params):
  try:
    ##Set variables
    sheet_name = params.get("sheet_name")
    file_path = params.get("file_name") 
    col_idx1 =  int(params.get("col_idx1"))
    col_idx2 = int(params.get("col_idx2"))
    inconsistencias_file = params.get("inconsistencias_file")
    validation_type = int(params.get("validation_type"))

    ##Validate if all the required input are present
    if not all([sheet_name, file_path, col_idx1, col_idx2, inconsistencias_file, validation_type]):
      return "Error: an input param is missing"

    ##Load workbook
    df: pd.DataFrame = pd.read_excel(file_path, engine="openpyxl", sheet_name=sheet_name)

    ## Convert both columns to datetime to ensure comparison works
    df.iloc[:, col_idx1] = pd.to_datetime(df.iloc[:, col_idx1], errors='coerce')  # Convert to datetime
    df.iloc[:, col_idx2] = pd.to_datetime(df.iloc[:, col_idx2], errors='coerce')  # Convert to datetime

    ##Filter data frame
    dic_validation_type = {
      1: less,
      2: greater_or_equal,
      3: less_or_equal,
      4: greater
    }

    filtered_file = dic_validation_type.get(validation_type)(df, col_idx1, col_idx2)

    print(filtered_file)

    ##Validate if the file is empty and save depends on that
    if not filtered_file.empty:
      filtered_file['COORDENADAS_1'] = filtered_file.apply(
        lambda row: f"{get_excel_column_name(col_idx1 + 1)}{row.name + 2}", axis=1
      )
      filtered_file['COORDENADAS_2'] = filtered_file.apply(
        lambda row: f"{get_excel_column_name(col_idx2 + 1)}{row.name + 2}", axis=1
      )

      new_sheet_name = "Fechas"
      if os.path.exists(inconsistencias_file):
        with pd.ExcelFile(inconsistencias_file, engine="openpyxl") as xls:
          if new_sheet_name in xls.sheet_names:
            existing_df = pd.read_excel(xls, sheet_name=new_sheet_name, engine="openpyxl")
            filtered_file = pd.concat([existing_df, filtered_file], ignore_index=True)

      with pd.ExcelWriter(inconsistencias_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        filtered_file.to_excel(writer, sheet_name=new_sheet_name, index=False)
        return "Inconsistencias registradas correctamente"
    else:
      return f"No hay inconsistencias en las columnas {col_idx1} vs {col_idx2}"

  except Exception as e:
    return f"Error: {e}"

def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ''
    while n > 0:
        n, remainder = divmod(n-1, 26)
        result = chr(65 + remainder) + result
    return result

def less(df: pd.DataFrame, col_idx1: int, col_idx2: int):
  df["VALIDACION_FECHA"]= df.iloc[:, col_idx1] < df.iloc[:, col_idx2]
  return df[~df["VALIDACION_FECHA"]].copy()

def greater_or_equal(df: pd.DataFrame, col_idx1: int, col_idx2: int):
  df["VALIDACION_FECHA"]= df.iloc[:, col_idx1] >= df.iloc[:, col_idx2]
  return df[~df["VALIDACION_FECHA"]].copy()

def less_or_equal(df: pd.DataFrame, col_idx1: int, col_idx2: int):
  df["VALIDACION_FECHA"]= df.iloc[:, col_idx1] <= df.iloc[:, col_idx2]
  return df[~df["VALIDACION_FECHA"]].copy()

def greater(df: pd.DataFrame, col_idx1: int, col_idx2: int):
  df["VALIDACION_FECHA"]= df.iloc[:, col_idx1] <= df.iloc[:, col_idx2]
  return df[~df["VALIDACION_FECHA"]].copy()

if __name__== "__main__":
  params = {
    "file_name": "C:/ProgramData/AutomationAnywhere/Bots/Logs/AD_RCSN_SabanaPagosYBasesParaSinestralidad/TempFolder/BASE DE REPARTO 2024.xlsx",
    "sheet_name": "CASOS NUEVOS",
    "col_idx1": "21",
    "col_idx2": "24",
    "inconsistencias_file": "C:/ProgramData/AutomationAnywhere/Bots/Logs/AD_RCSN_SabanaPagosYBasesParaSinestralidad/OutputFolder/Inconsistencias/InconBaseReparto.xlsx",
    "validation_type": "3",
  }  
  print(main(params))