import os
from tools.boarding import board_by_excel, create_table

def table7(path_subarea) -> None:
    path_parts = path_subarea.split("/") #<--- LINUX
    subarea_id = path_parts[-1]
    proyect_folder = '/'.join(path_parts[:-2]) #<--- LINUX

    field_data = os.path.join(
        proyect_folder,
        "7. Informacion de Campo",
        subarea_id,
        "Embarque y Desembarque"
    )

    excels_by_tipicidad = {}
    for tipicidad in ["Tipico", "Atipico"]:
        tipicidad_folder = os.path.join(field_data, tipicidad)
        list_excels = os.listdir(tipicidad_folder)
        list_excels = [os.path.join(tipicidad_folder, file) for file in list_excels
                       if file.endswith(".xlsx") and not file.startswith("~")]
        excels_by_tipicidad[tipicidad] = list_excels

    tables_by_tipicidad = {}
    for tipicidad, excel_list in excels_by_tipicidad.items():
        tables_by_code = {}
        for excel in excel_list:
            codigo, tableList = board_by_excel(excel)
            tables_by_code[codigo] = tableList
        tables_by_tipicidad[tipicidad] = tables_by_code

    for key, value in tables_by_tipicidad.items():
        for code, tableData in value.items():
            create_table(tableData, key, code)