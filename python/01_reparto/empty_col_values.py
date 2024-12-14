import pandas as pd #type: ignore

def main(params: dict):
  """This function validate if a column cell should be empty or not 
  an report an inconsistency into a file. This validation only works with mandatory empty columns"""
  try: 
    ##Set the initial variables to reuse the source code
    file_path = params.get ("file_path")
    sheet_name = params.get("sheet_name")
    col_idx = int(params.get("col_idx"))
    inconsistencies_file = params.get("inconsistencies_file")

    ##Validate input params required are present 
    if not all([file_path, sheet_name, col_idx]):
      return "Error: an input required file is missing"

    ##Read the work book
    book: pd.DataFrame = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

    ##Apply validation
    book["EMPTY_VALUES"] = book.iloc[:, col_idx].astype(str).apply(lambda value: is_valid(value))

    ##Filter the inconsistencies data frame
    ##Create a copy to avoid modify the original data frame
    inconsistencies = book[~book["EMPTY_VALUES"]].copy()
    
    return (inconsistencies.iloc[:, col_idx])
    

    ##Write into the inconsistencies file
    if not inconsistencies.empty:
      inconsistencies['COORDENADAS'] = inconsistencies.apply(
        lambda row: f"{get_excel_column_name(col_idx + 1)}{row.name + 2}", axis=1
      )

      in_sheet_name = "ValidacionColumnasSinEspacios"
      with pd.ExcelFile(inconsistencies_file, engine="openpyxl") as xls:
        if in_sheet_name in xls.sheet_names:
          exiting = pd.read_excel(xls, sheet_name=in_sheet_name, engine="openpyxl")
          inconsistencies = pd.concat([exiting, inconsistencies], ignore_index=True)

      with pd.ExcelWriter(
          inconsistencies_file, engine="openpyxl", mode="a", 
          if_sheet_exists="replace"
        ) as writer:
        inconsistencies.to_excel(writer, sheet_name=in_sheet_name, index=False)
        return "Inconsistencias registradas correctamente"
      
    else:
      return "No se encontraron inconsistencias"

  except Exception as e:
    return f"Error: {e}"
  
def is_valid(input: str) -> bool:
  if input.lower() == "nan":
    return True
  else:
    return False
  
def get_excel_column_name(n):
  """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
  result = ''
  while n > 0:
      n, remainder = divmod(n-1, 26)
      result = chr(65 + remainder) + result
  return result

"""Apply with a use case"""
if __name__ == "__main__":
  params = {
    "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
    "sheet_name": "CASOS NUEVOS",
    "col_idx": "39",
    "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx"
  }

  print(main(params))