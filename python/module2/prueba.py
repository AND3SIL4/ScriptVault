import pandas as pd

file_path = r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\InputFolder\BDD PAGOS RECONOCIMIENTO POLIZAS DE VIDA PAGOS 2024 - OTROS RAMOS.xlsx"


data = pd.read_excel(file_path, engine="openpyxl")

print(data.iloc[:, :111])