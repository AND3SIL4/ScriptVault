import pandas as pd
import os

def main(params: dict):
  try:
    ##Set variables
    file_path = params.get("file_path")
    sheet_name = params.get("sheet_name")
    inconsistencies_file = params.get("inconsistencies_file")

    ##Validate if the input params are present
    if not all([file_path, sheet_name, inconsistencies_file]):
      return "Error: an input param required is missing"

    ##Read the file
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

    ##Select cols to compare if the ramo code match with the siniestro
    df["col1"] = df["No. SINIESTRO"].astype(str)
    df["col2"] = df["CODIGO RAMO SAP"].astype(str).str[-2:]
    df["col3"] = df.iloc[:, 18].astype(str)

    ##Validate if col "Siniestro" :2 is equals to col "Codigo ramo" -2: 
    df["match"] = df["col1"].str[:2] == df["col2"]
    no_match = df[df["match"] == False].copy()

    ##Validate if the col "Siniestro" is equals to "Documento Riesgo"
    no_match["id"] = no_match["col1"] == no_match["col3"]
    doc_equals = no_match.loc[no_match["id"] == False]

    ##Make validation
    if (no_match.empty or doc_equals.empty):
      return "Todos los códigos coinciden correctamente"
    else:
      new_sheet_name = "ValidacionCodigos"
      if os.path.exists(inconsistencies_file):
        with pd.ExcelFile(inconsistencies_file, engine="openpyxl") as xls:
          if new_sheet_name in xls.sheet_names:
            existing_df = pd.read_excel(xls, sheet_name=new_sheet_name, engine="openpyxl")
            doc_equals = pd.concat([existing_df, doc_equals], ignore_index=True)

      with pd.ExcelWriter(inconsistencies_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        doc_equals.to_excel(writer, sheet_name="ValidacionCodigos", index=False)
        return "Inconsistencias registradas correctamente"

  except Exception as e:  
    return f"Error: {e}"

if __name__ == "__main__":
  dic = {
    "inconsistencies_file": "C:/ProgramData/AutomationAnywhere/Bots/Logs/AD_RCSN_SabanaPagosYBasesParaSinestralidad/OutputFolder/Inconsistencias/InconBasePagos.xlsx",
    "sheet_name": "PAGOS",
    "file_path": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE PAGOS.xlsx",
  }
  print(main(dic)) 