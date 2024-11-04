import os
from tables.tools.boarding import board_by_excel, create_table
#from tools.boarding import board_by_excel, create_table
from docx import Document
from pathlib import Path
from dataclasses import asdict
import pandas as pd
from docxtpl import DocxTemplate

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
    listCodes = []
    for tipicidad, excel_list in excels_by_tipicidad.items():
        tables_by_code = {}
        for excel in excel_list:
            codigo, tableList, name = board_by_excel(excel)
            if tipicidad == "Tipico":
                listCodes.append(codigo)
                dataDict = tableList.copy()
                dataDict = [asdict(obj) for obj in dataDict]
                df = pd.DataFrame(dataDict) #<--- TODO: Estoy guardando lo que va a ser usado, ahora debes usar solo el promedio.
                df['maximo'] = pd.to_numeric(df['maximo'], errors='coerce')
                df['promedio'] = pd.to_numeric(df['promedio'], errors='coerce')
                df['desviacion'] = pd.to_numeric(df['desviacion'], errors='coerce')
                listDfs.append((df, codigo, name))
            tables_by_code[codigo] = tableList
        tables_by_tipicidad[tipicidad] = tables_by_code

    ####################
    # Creating table 7 #
    ####################

    count = 1
    listTableEmbarking = {}
    for tipicidad, value in tables_by_tipicidad.items():
        if tipicidad == "Tipico":
            for code, tableData in value.items():
                table7_path_ref = create_table(tableData, code, count, path_subarea)
                listTableEmbarking[code] = table7_path_ref
                count += 1

    ###########################
    # Creating embarking list #
    ###########################

    embarkingFolder = path_subarea / "Tablas" / "Embarking"
    embarkingFolder.mkdir(parents=True, exist_ok=True)

    # Creating embarking list by code only Tipico typicality 
    listWordsEmbarking = {}
    for df, code, name in listDfs:
        dfMeanByTurn = df.groupby('turno')['promedio'].mean().round(2)
        doc = DocxTemplate("templates/template_embarquelist.docx")
        tableEmbarking = doc.new_subdoc(listTableEmbarking[code])
        variables = {
            "codinterseccion": code,
            "nominterseccion": name,
            "temprom_morning": dfMeanByTurn['Mañana'],
            "temprom_afternoon": dfMeanByTurn['Medio dia'],
            "temprom_night": dfMeanByTurn['Noche'],
            "table": tableEmbarking,
        }
        doc.render(variables)
        savePath = path_subarea / "Tablas" / "Embarking" / f"table7_{code}.docx"
        doc.save(savePath)
        listWordsEmbarking[code] = savePath

    # This will to be at the end with the mixed list
    table7_path = Path(path_subarea) / "Tablas" / "table7.docx"
    for i, code in enumerate(listCodes):
        if i == 0:
            fullPathRef = listWordsEmbarking[code]
            docTarget = Document(fullPathRef)
            continue

        docSource = Document(listWordsEmbarking[code])
        append_document_content(docSource, docTarget)

    docTarget.save(table7_path)

    return table7_path