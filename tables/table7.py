import os
#from tables.tools.boarding import board_by_excel, create_table
from tools.boarding import board_by_excel, create_table
from docx import Document
from pathlib import Path
from dataclasses import asdict
import pandas as pd

def append_document_content(source_doc, target_doc) -> None:
    for element in source_doc.element.body:
        target_doc.element.body.append(element)

def create_table7(path_subarea) -> None:
    path_subarea = Path(path_subarea)
    subarea_id = path_subarea.name
    proyect_folder = path_subarea.parents[1]

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
    listDfs = []
    for tipicidad, excel_list in excels_by_tipicidad.items():
        tables_by_code = {}
        for excel in excel_list:
            codigo, tableList, name = board_by_excel(excel)
            if tipicidad == "Tipico":
                dataDict = tableList.copy()
                dataDict = [asdict(obj) for obj in dataDict]
                df = pd.DataFrame(dataDict) #<--- TODO: Estoy guardando lo que va a ser usado, ahora debes usar solo el promedio.
            listDfs.append((df, codigo, name))
            tables_by_code[codigo] = tableList
        tables_by_tipicidad[tipicidad] = tables_by_code

    return None

    count = 1
    list_REF = []
    for tipicidad, value in tables_by_tipicidad.items():
        if tipicidad == "Tipico":
            for code, tableData in value.items():
                table7_path_ref = create_table(tableData, code, count, path_subarea)
                list_REF.append(table7_path_ref)
                count += 1

    doc_target = Document(list_REF[0])
    table7_path = Path(path_subarea) / "Tablas" / "table7.docx"
    for i in range(len(list_REF)):
        if i == 0:
            continue
        doc_source = Document(list_REF[i])
        append_document_content(doc_source, doc_target)
        
        doc_target.save(table7_path)
        doc_target = Document(table7_path)

    doc_target.save(table7_path)
    return table7_path

if __name__ == '__main__':
    subareaPath = r"C:\Users\dacan\OneDrive\Desktop\PRUEBAS\Maxima Entropia\02 Proyecto SJL-El Agustino (57 Int. - 18 SA)\6. Sub Area Vissim\Sub Area 041"
    create_table7(subareaPath)