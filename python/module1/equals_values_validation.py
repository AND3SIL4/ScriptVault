import pandas as pd #type: ignore
import os

def main(params: dict):
  try:
    ##Set variables
    file_path = params.get("file_path")
    inconsistencias_file = params.get("inconsistencias_file")
    sheet_name = params.get("sheet_name")
    col1 = int(params.get("col1"))
    col2 = int(params.get("col2"))
    is_radicado = bool(params.get("is_radicado"))
    list_file = params.get("list_file")
    exception_col = int(params.get("exception_col"))

    ##Validate if all the inputs required are present
    if not all([file_path, inconsistencias_file, sheet_name, col1, col2, list_file]):
      return "Error: an input param required is missing"

    ##Read File
    df: pd.DataFrame = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
    list_df: pd.DataFrame = pd.read_excel(list_file, sheet_name="EXCEPCIONES SARLAF", engine="openpyxl")

    ##Apply validation
    df["VALIDATION_SARLAF"] = df.apply(
      lambda row: is_valid(
        str(row.iloc[col1]),
        str(row.iloc[col2]),
        list_df.iloc[:, exception_col].dropna().astype(str).values.tolist(),
        is_radicado
      ),
      axis=1
    )

    ##Filter data frame
    filtered_file = df[~df["VALIDATION_SARLAF"]].copy()

    print(filtered_file)

    ##Make a validation and save the inconsistencies
    if filtered_file.empty:
      return "Validación realizada correctamente, no hay inconsistencias para registrar"
    else:
      filtered_file["COORDENADA_1"] = filtered_file.apply(
          lambda row: f"{get_excel_column_name(col1 + 1)}{row.name + 2}", axis=1
      )
      filtered_file["COORDENADA_2"] = filtered_file.apply(
          lambda row: f"{get_excel_column_name(col2 + 1)}{row.name + 2}", axis=1
      )
      
      new_sheet_name = "SarlafValidacion"
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

def is_valid(value1: str, value2: str, list_exception: list[str], is_radicado: bool) -> bool:
  """Method to know if the value is equals and validate into the exception list"""
  if is_radicado:
    validation = value1 == value2
    return validation
  else:
    validation = (value1 == value2) or (value1 in list_exception)
    return validation

def get_excel_column_name(n):
  """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
  result = ''
  while n > 0:
      n, remainder = divmod(n-1, 26)
      result = chr(65 + remainder) + result
  return result

if __name__ == "__main__":
  params  = {
    "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
    "inconsistencias_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
    "sheet_name": "CASOS NUEVOS",
    "col1": "2",
    "col2": "23",
    "is_radicado": True,
    "list_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\Listas - BOT.xlsx",
    "exception_col": "0"
  }

  print(main(params))
