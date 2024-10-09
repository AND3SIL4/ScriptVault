import pandas as pd  # type: ignore
import re


def main(params: dict):
    try:
        ##Set initial variables
        file_path: str = params.get("file_path")
        sheet_name: str = params.get("sheet_name")

        ##Validate if all the required inputs are present
        if not all([file_path, sheet_name]):
            return "ERROR: an required input is missing"

        data_frame: pd.DataFrame = pd.read_excel(
            file_path, sheet_name=sheet_name, engine="openpyxl"
        )

        ##Fix white spaces
        data_frame["MES DE ASIGNACION"] = (
            data_frame["MES DE ASIGNACION"]
            .astype(str)
            .apply(lambda x: clean_white_spaces(x))
        )

        months_by_number = {
            "ENERO": 1,
            "FEBRERO": 2,
            "MARZO": 3,
            "ABRIL": 4,
            "MAYO": 5,
            "JUNIO": 6,
            "JULIO": 7,
            "AGOSTO": 8,
            "SEPTIEMBRE": 9,
            "OCTUBRE": 10,
            "NOVIEMBRE": 11,
            "DICIEMBRE": 12,
        }

        data_frame["MES_NUMERO"] = data_frame["MES DE ASIGNACION"].map(months_by_number)

        dynamic_table = pd.pivot_table(
            data_frame,
            values="VALOR RESERVA",
            index="RAMO",
            columns="MES DE ASIGNACION",
            aggfunc="sum",
            fill_value=0,
        )

        dynamic_table = dynamic_table.reindex(
            columns=[
                k for k, v in sorted(months_by_number.items(), key=lambda item: item[1])
            ]
        )

        new_sheet = "REPORTE"
        with pd.ExcelWriter(
            file_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            dynamic_table.to_excel(writer, sheet_name=new_sheet, index=True)
            return "Tabla guardada correctamente"
    except Exception as e:
        return f"ERROR: {e}"


def clean_white_spaces(string: str):
    value = re.sub(r"[\s]", "", string)
    return value


if __name__ == "__main__":
    params = {
        "file_path": r"C:\ProgramData\AutomationAnywhere\Bots\Logs\AD_RCSN_SabanaPagosYBasesParaSinestralidad\TempFolder\BASE DE REPARTO 2024.xlsx",
        "sheet_name": "CASOS NUEVOS",
    }

    print(main(params))
