import pandas as pd
import os
import re

def main(params: dict):
  try:
    ##Set initial variables
    file_path = params.get("file_path")
    in_file = params.get("in_file")
    sheet_name = params.get("sheet_name")
    col_name = params.get("col_name")
    lst = params.get("lst")
    in_sheet = params.get("in_sheet")

    ##Validate if all the input params are present
    if not all([file_path, in_file, sheet_name, col_name, lst, in_sheet]):
      return "Error: an input param is missing"

    ##Read the work book
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")

    ##Validate the length in a specific column
    is_in =  df[~df[col_name].isin(lst)].copy()
  
    print(is_in)

    if is_in.empty:
      return "Validación realizada, no se encontraron inconsistencias"
    else:
      col_index = df.columns.get_loc(col_name) + 1  # Get column number (1-based)
      is_in['COORDENADAS'] = is_in.apply(
        lambda row: f"{get_excel_column_name(col_index)}{row.name + 2}", axis=1
      )

      ##Register into the inconsistencies file
      new_sheet = in_sheet
      ##Validate is the inconsistencies file exist
      if os.path.exists(in_file):
        with pd.ExcelFile(in_file, engine="openpyxl") as xls:
          if new_sheet in xls.sheet_names:
            existing_file = pd.read_excel(xls, engine="openpyxl", sheet_name=new_sheet)
            existing_file = existing_file.dropna(how="all", axis=1) ##Does not include the empty columns
            is_in = pd.concat([existing_file, is_in], ignore_index=True)

      ##Store data
      with pd.ExcelWriter(in_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        is_in.to_excel(writer, sheet_name=new_sheet, index=False)
        return "Inconsistencias registradas correctamente"
      
  except Exception as e:
    return f"Error: {e}"
  
def clean(string):
  """Validate the white spaces in strings."""
  string = re.sub(r"\s+", " ", string).strip()
  return string

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
    "in_file": "C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\OutputFolder\Inconsistencias\InconBaseReparto.xlsx",
    "sheet_name": "CASOS NUEVOS",
    "col_name": "RAMO",
    "lst": [
      "ACCIDENTES PERSONALES INDIVIDUAL",
      "ACCIDENTES PERSONALES GENERACION POSITIVA",
      "EXEQUIAS",
      "SALUD",
      "VIDA GRUPO",
      "VIDA GRUPO DEUDORES",
      "VIDA INDIVIDUAL",
      "DESEMPLEO",
    ],
    "in_sheet": "ValidacionRamos"

  }

  print(main(params))

"""
lst = [
  AUXILIO EDUCATIVO POR MUERTE DE LOS PADRES
AUXILIO FUNERARIO
AUXILIO POR ACCIDENTE
AUXILIO POR DESEMPLEO
AUXILIO POR DIAGNOSTICO DE CUALQUIER TIPO DE CANCER
AUXILIO POR MATRICULA POR ACCIDENTE
AUXILIO POR NO UTILIZACION DE LA POLIZA
AUXILIO POR PATERNIDAD O MATERNIDAD
AUXILIO POR TRATAMIENTO AMBULATORIO FUERA DE LA SEDE DE TRABAJO
BENEFICIO ADICIONAL POR MUERTE O DESMEMBRACION ACCIDENTAL
BENEFICIO DIARIO POR INCAPACIDAD TEMPORAL
DESEMPLEO
ENFERMEDADES GRAVES
ENFERMEDADES TROPICALES
FRACTURA
GASTOS DE TRASLADO
GASTOS MEDICOS
GASTOS MEDICOS POR COMPLICACION EN CIRUGIA
HERIDA ABIERTA
HOMICIDIO
INCAPACIDAD TOTAL Y PERMANENTE
INVALIDEZ Y/O DESMEMBRACION POR ACCIDENTE
MUERTE
MUERTE ACCIDENTAL
PROTECCION GARANTIZADA
REEMBOLSO PARA TRASLADO DE RESTOS MORTALES POR ACCIDENTE DEL ASEGURADO
REEMBOLSOS DE GASTOS FUNERARIOS
REHABILITACION INTEGRAL POR INVALIDEZ
RENTA DIARIA POR HOSPITALIZACION
RENTA DIARIA POR INCAPACIDAD DOMICILIARIA
RENTA MENSUAL POR INCAPACIDAD TOTAL Y PERMANENTE
RENTA MENSUAL POR INCAPACIDAD TOTAL Y PERMANENTE
RENTA MENSUAL POR MUERTE
RIESGO BIOLOGICO
SERVICIO DE AMBULANCIA AEREA
SUBSIDIO EDUCATIVO POR DESEMPLEO INVOLUNTARIO
TRASLADO RESTOS MORTALES POR ACCIDENTE
URGENCIA MEDICA

]"""