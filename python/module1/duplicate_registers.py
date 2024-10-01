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

    ##Create imaginary key to make the validation
    base["KEY_1"] = base.iloc[:, 0].astype(str) + "-" + base.iloc[:, 2].astype(str)
    base["KEY_2"] = base["KEY_1"] + "-" + base.iloc[:, 32].astype(str)
    base["KEY_3"] = base["KEY_2"] + "-" + base.iloc[:, 34].astype(str)
    base["KEY_4"] = base["KEY_2"] + "-" + base.iloc[:, 27].astype(str)
    base["KEY_5"] = base["KEY_2"] + "-" + base.iloc[:, 98].astype(str)
    base["KEY_6"] = base.iloc[:, 18].astype(str) + "-" + base.iloc[:, 32].astype(str) + "-" + base.iloc[:, 34].astype(str)
    base["KEY_7"] =base.iloc[:, 18].astype(str) + "-" + base.iloc[:, 32].astype(str) + "-" + base.iloc[:, 98].astype(str)

    file = base[(base["KEY_1"].duplicated(keep=False)) &
                (base["KEY_2"].duplicated(keep=False)) & 
                (base["KEY_3"].duplicated(keep=False)) & 
                (base["KEY_4"].duplicated(keep=False)) &
                (base["KEY_5"].duplicated(keep=False))]

    print(file)


  except Exception as e:
    return f"ERROR: {e}"
  
if __name__ == "__main__":
  params = {
    "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx", 
    "sheet_name": "CASOS NUEVOS",
    "inconsistencies_file": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
  }

  print(main(params))