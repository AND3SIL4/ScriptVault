import pandas as pd #type: ignore
import os

def main(params: dict) -> str:
  try:
    ##Set the initial variables
    file_path: str = params.get("file_path")
    sheet_name: str = params.get("sheet_name")
    inconsistencies_file: str = params.get("inconsistencies_file")

    ##Validate if all the required inputs are present
    if not all(
      [
        file_path,
        sheet_name,
        inconsistencies_file
      ]
    ):
      return "ERROR: an input required param is missing"

    ##Read the work book
    base: pd.DataFrame = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

    ##Replace the NaN values with 0 in column "Credito"
    base.iloc[:, 98] =  base.iloc[:, 98].fillna(0)

    ##Create imaginary key to make the validation
    base["KEY_1"] = base.iloc[:, 0].astype(str) + "-" + base.iloc[:, 2].astype(str)
    base["KEY_2"] = base["KEY_1"] + "-" + base.iloc[:, 32].astype(str)
    base["KEY_3"] = base["KEY_2"] + "-" + base.iloc[:, 34].astype(str)
    base["KEY_4"] = base["KEY_2"] + "-" + base.iloc[:, 27].astype(str)
    base["KEY_5"] = base["KEY_2"] + "-" + base.iloc[:, 98].astype(str)
    base["KEY_6"] = base.iloc[:, 18].astype(str) + "-" + base.iloc[:, 32].astype(str) + "-" + base.iloc[:, 34].astype(str)
    base["KEY_7"] =base.iloc[:, 18].astype(str) + "-" + base.iloc[:, 32].astype(str) + "-" + base.iloc[:, 98].astype(str)

    first_validation = base[
      (base["KEY_1"].duplicated(keep=False)) &
      (base["KEY_2"].duplicated(keep=False)) & 
      (base["KEY_3"].duplicated(keep=False)) & 
      (base["KEY_4"].duplicated(keep=False)) &
      (base["KEY_5"].duplicated(keep=False))
    ]
    
    second_validation = base[
      (base["KEY_1"].duplicated(keep=False)) &
      (base["KEY_2"].duplicated(keep=False)) &
      (base["KEY_4"].duplicated(keep=False)) &
      (base["KEY_5"].duplicated(keep=False))
    ]

    third_validation = base[
      (base["KEY_1"].duplicated(keep=False)) &
      (base["KEY_6"].duplicated(keep=False)) &
      (base["KEY_7"].duplicated(keep=False)) &
      (base["KEY_4"].duplicated(keep=False))
    ]

    fourth_validation = base[
      (base["KEY_6"].duplicated(keep=False)) & 
      (base["KEY_7"].duplicated(keep=False))
    ]

    print(fourth_validation) ##!Uncomment or delete

    fourth_validation.to_excel(inconsistencies_file, index=False)

    ##Keys validation
    validate_empty_df(inconsistencies_file, "Llave1-2-3-4-5", first_validation)
    validate_empty_df(inconsistencies_file, "Llave1-2-4-5", second_validation)
 
  except Exception as e:
    return f"ERROR: {e}"
  

def append_inconsistencies_file(path_file: str, new_sheet: str, data_frame: pd.DataFrame)-> str:
  """
  This function append a data frame into the inconsistencies file when is necessary
 
  Params:
  str: path_file (the path of the inconsistencies file)
  str: new_sheet (the name that you wanna put into the sheet in inconsistencies file)
  pd.DataFrame: data_frame (the data frame filtered previously)

  Returns:
  str: The confirm message
  """
  if os.path.exists(path_file):
    with pd.ExcelFile (path_file, engine="openpyxl") as xls:
      if new_sheet in xls.sheet_names:
        existing = pd.read_excel(xls, sheet_name=new_sheet, engine="openpyxl")
        data_frame = pd.concat([existing, data_frame], ignore_index=True)

    with pd.ExcelWriter (path_file, engine="openpyxl", if_sheet_exists="replace", mode="a") as writer:
      data_frame.to_excel(writer, sheet_name=new_sheet, index=False)
      return "Inconsistencias registradas correctamente"

def validate_empty_df(path_file: str, new_sheet: str, data_frame: pd.DataFrame) -> str:
  """
  This function return a string message depends on the validation made
  
  Params:
  str: path_file (the path of the inconsistencies file)
  str: new_sheet (the name that you wanna put into the sheet in inconsistencies file)
  pd.DataFrame: data_frame (the data frame filtered previously)

  Return:
  str: confirmation message
  """
  if not data_frame.empty:
    return append_inconsistencies_file(path_file, new_sheet, data_frame)
  else:
    return "Validation realizada, no se encontraron inconsistencias"

if __name__ == "__main__":
  params = {
    "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx", 
    "sheet_name": "CASOS NUEVOS",
    "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
  }

  print(main(params))