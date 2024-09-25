import pandas as pd
import os

def main(params: dict):
  try:
    ##Set variables
    file_path = params.get("file_path")
    sheet_name = params.get("sheet_name")
    inconsistencies_file = params.get("inconsistencies_file")
    siniestro_col = int(params.get("siniestro_col"))
    ramo_col = int(params.get("ramo_col"))
    document_col = int(params.get("document_col"))

    ##Validate if the input params are present
    if not all([file_path, sheet_name, inconsistencies_file]):
      return "Error: an input param required is missing"

    ##Read the file
    df: pd.DataFrame = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

    ##Select cols to compare if the ramo code match with the siniestro
    siniestro = df.iloc[:, siniestro_col].astype(str)
    ramo = df.iloc[:, ramo_col].astype(str).str[-2:]
    document = df.iloc[:, 18].astype(str)

    ##Validate if col "Siniestro" :2 is equals to col "Codigo ramo" -2: 
    df["is_valid"] = (siniestro.str[: 2] == ramo) | (siniestro == document)

    ##Validate if the col "Siniestro" is equals to "Documento Riesgo"
    inconsistencies = df[~df["is_valid"]].copy()

    print(inconsistencies) #!Uncomment to deploy

    ##Make validation
    if (inconsistencies.empty):
      return "Todos los códigos coinciden correctamente"
    else:
      # Add a column with Excel coordinates (e.g., A2, B3) of the inconsistent cells
      inconsistencies['COORDENADA_1'] = inconsistencies.apply(
        lambda row: f"{get_excel_column_name(siniestro_col + 1)}{row.name + 2}", axis=1
      )
      # Add a column with Excel coordinates (e.g., A2, B3) of the inconsistent cells
      inconsistencies['COORDENADA_2'] = inconsistencies.apply(
        lambda row: f"{get_excel_column_name(ramo_col + 1)}{row.name + 2}", axis=1
      )
      # Add a column with Excel coordinates (e.g., A2, B3) of the inconsistent cells
      inconsistencies['COORDENADA_3'] = inconsistencies.apply(
        lambda row: f"{get_excel_column_name(document_col + 1)}{row.name + 2}", axis=1
      )

      new_sheet_name = "ValidationCodigosRamo"
      if os.path.exists(inconsistencies_file):
        with pd.ExcelFile(inconsistencies_file, engine="openpyxl") as xls:
          if new_sheet_name in xls.sheet_names:
            existing_df = pd.read_excel(xls, sheet_name=new_sheet_name, engine="openpyxl")
            inconsistencies = pd.concat([existing_df, inconsistencies], ignore_index=True)

      with pd.ExcelWriter(inconsistencies_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        inconsistencies.to_excel(writer, sheet_name="ValidacionCodigos", index=False)
        return "Inconsistencias registradas correctamente"

  except Exception as e:  
    return f"Error: {e}"

def get_excel_column_name(n):
    """Convert a column number (1-based) to Excel column name (e.g., 1 -> A, 28 -> AB)."""
    result = ''
    while n > 0:
        n, remainder = divmod(n-1, 26)
        result = chr(65 + remainder) + result
    return result

if __name__ == "__main__":
  dic = {
    "inconsistencies_file": "C:/ProgramData/AutomationAnywhere/Bots/Logs/AD_RCSN_SabanaPagosYBasesParaSinestralidad/OutputFolder/Inconsistencias/InconBaseReparto.xlsx",
    "sheet_name": "CASOS NUEVOS",
    "file_path": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
    "siniestro_col": "0",
    "ramo_col": "11",
    "document_col": "18"
  }
  print(main(dic)) 