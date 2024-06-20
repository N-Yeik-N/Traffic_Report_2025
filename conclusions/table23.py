import os
import json
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

tipicoList = ["HPM", "HPT", "HPN"]
atipicoList = ["HPM", "HPT", "HPN"]

def _align_content(table) -> None:
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                for run in paragraph.runs:
                    run.font.size = Pt(7)
                    run.font.name = 'Arial'

    for i in range(len(table.columns)):
        cell = table.cell(0,i)
        cell_xml_element = cell._tc
        table_cell_properties = cell_xml_element.get_or_add_tcPr()
        shade_obj = OxmlElement('w:shd')
        shade_obj.set(qn('w:fill'),'B4C6E7')
        table_cell_properties.append(shade_obj)

    for cell in table.rows[0].cells:
        for paragraph in cell.paragraphs:
            run = paragraph.runs[0]
            run.font.bold = True

    for row in table.rows:
        for col_index in range(3):  # Columnas 0, 1 y 2
            cell = row.cells[col_index]
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True

def create_table23(subareaPath):
    actualPath = os.path.join(subareaPath, "Actual")
    basePath = os.path.join(subareaPath, "Output_Base")
    listJSONPathsActual = []
    listJSONPathsBase = []
    listNames = []
    for tipicidad in ["Tipico", "Atipico"]:
        #Actual
        tipicidadPathActual = os.path.join(actualPath, tipicidad)
        scenariosListActual = os.listdir(tipicidadPathActual)
        scenariosListActual = [file for file in scenariosListActual if not file.endswith(".ini") and file in ["HPM", "HPT", "HPN"]]

        #Output Base
        tipicidadPathBase = os.path.join(basePath, tipicidad)
        scenariosListBase = os.listdir(tipicidadPathBase)
        scenariosListBase = [file for file in scenariosListBase if not file.endswith(".ini") and file in ["HPM", "HPT", "HPN"]]

        if tipicidad == "Tipico":
            for tipicoUnit in tipicoList:
                for scenarioNameActual, scenarioNameBase in zip(scenariosListActual, scenariosListBase):
                    if tipicoUnit == scenarioNameActual:

                        scenarioPathActual = os.path.join(tipicidadPathActual, scenarioNameActual)
                        scenarioContentActual = os.listdir(scenarioPathActual)
                        if "table.json" in scenarioContentActual:
                            jsonFileActual = os.path.join(scenarioPathActual, "table.json")
                            listJSONPathsActual.append(jsonFileActual)    
                            listNames.append(scenarioNameActual)

                        scenarioPathBase = os.path.join(tipicidadPathBase, scenarioNameBase)
                        scenarioContentBase = os.listdir(scenarioPathBase)
                        if "table.json" in scenarioContentBase:
                            jsonFileBase = os.path.join(scenarioPathBase, "table.json")
                            listJSONPathsBase.append(jsonFileBase)

        elif tipicidad == "Atipico":
            for tipicoUnit in tipicoList:
                for scenarioNameActual, scenarioNameBase in zip(scenariosListActual, scenariosListBase):
                    if tipicoUnit == scenarioNameActual:
                        
                        scenarioPathActual = os.path.join(tipicidadPathActual, scenarioNameActual)
                        scenarioContentActual = os.listdir(scenarioPathActual)
                        if "table.json" in scenarioContentActual:
                            jsonFileActual = os.path.join(scenarioPathActual, "table.json")
                            listJSONPathsActual.append(jsonFileActual)    
                            listNames.append(scenarioNameActual)

                        scenarioPathBase = os.path.join(tipicidadPathBase, scenarioNameBase)
                        scenarioContentBase = os.listdir(scenarioPathBase)
                        if "table.json" in scenarioContentBase:
                            jsonFileBase = os.path.join(scenarioPathBase, "table.json")
                            listJSONPathsBase.append(jsonFileBase)

    doc = Document()
    table = doc.add_table(rows = 1, cols = 9)

    table.cell(0,0).text = "Tipicidad"
    table.cell(0,1).text = "Escenario"
    table.cell(0,1).merge(table.cell(0,2))

    #Creating Headers
    for i, texto in enumerate(["Nodo", "Número de\nVehículos\n(veh)", "Cola Máx.\nPromedio\n(m)", "Demora por parada\nPromedio\n(s/veh)", "Demora\nPromedio\n(s/veh)", "LOS"]):
        table.cell(0,i+3).text = texto

    nroRow = 1

    for i, jsonPathActual in enumerate(listJSONPathsActual):
        with open(jsonPathActual, 'r') as file:
            data = json.load(file)

        with open(listJSONPathsBase[i], 'r') as file2:
            data2 = json.load(file2)
        
        for j in range(len(data['nodes_names'])): #NOTE: Estoy considerando que son del mismo tamaño, puede que no sea así siempre.
            #Actual
            new_row = table.add_row()
            new_row.cells[2].text = "Actual"
            #new_row.cells[2].text = listNames[idScenario]
            new_row.cells[3].text = data['nodes_names'][j]                  #Nodo
            new_row.cells[4].text = str(int(data['nodes_totres'][j][4]))    #Número de Vehículos
            new_row.cells[5].text = str(int(data['nodes_totres'][j][3]))    #Cola Máx. Promedio
            new_row.cells[6].text = str(int(data['nodes_totres'][j][1]))    #Pare Promedio
            new_row.cells[7].text = str(int(data['nodes_totres'][j][0]))    #Demora Promedio
            new_row.cells[8].text = data['nodes_los'][j]
            nroRow += 1

            #Base
            new_row = table.add_row()
            table.cell(nroRow-1,3).merge(table.cell(nroRow,3))
            #table.cell(nroRow-1,2).merge(table.cell(nroRow,2))
            new_row.cells[2].text = "Propuesto"
            #new_row.cells[3].text = data2['nodes_names'][j]                 #Nodo
            new_row.cells[4].text = str(int(data2['nodes_totres'][j][4]))    #Número de Vehículos
            new_row.cells[5].text = str(int(data2['nodes_totres'][j][3]))    #Cola Máx. Promedio
            new_row.cells[6].text = str(int(data2['nodes_totres'][j][1]))    #Pare Promedio
            new_row.cells[7].text = str(int(data2['nodes_totres'][j][0]))    #Demora Promedio
            new_row.cells[8].text = data2['nodes_los'][j]
            nroRow += 1

        if i == 0:
            numberNodes = len(data['nodes_names'])

    idScenario = 0
    for j in range(1, len(table.rows), len(data['nodes_names'])*2):
        if idScenario >= 3: idScenario = 0
        table.cell(j,1).text = listNames[idScenario]
        if j == 1:
            beforeRow = j
            idScenario += 1
            continue
        table.cell(beforeRow,1).merge(table.cell(beforeRow+len(data['nodes_names'])*2-1,1))
        beforeRow = j
        idScenario += 1

    table.cell(beforeRow,1).merge(table.cell(beforeRow+len(data['nodes_names'])*2-1,1))
    
    #Tipicidad
    table.cell(1,0).text = "TÍPICO"
    table.cell(1,0).merge(table.cell(numberNodes*3*2,0)) #Solo 3 horas punta por propuesto y actual
    table.cell(numberNodes*3*2+1,0).text = "ATÍPICO"
    table.cell(numberNodes*3*2+1,0).merge(table.cell(numberNodes*3*2*2,0)) #Solo 3 horas puntas

    _align_content(table)

    finalPath = os.path.join(subareaPath, "Tablas", "resumenTable.docx")
    table.style = 'Table Grid'
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.save(finalPath)
    return finalPath