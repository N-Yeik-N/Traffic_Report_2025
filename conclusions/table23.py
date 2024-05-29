import os
import json
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

tipicoList = ["HVMAD","HPMAD","HVM","HPM","HVT","HPT","HVN","HPN"]
atipicoList = ["HVMAD", "HPM", "HPT", "HVN", "HPN"]

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
    listJSONPaths = []
    listNames = []
    for tipicidad in ["Tipico", "Atipico"]:
        tipicidadPath = os.path.join(actualPath, tipicidad)
        scenariosList = os.listdir(tipicidadPath)
        scenariosList = [file for file in scenariosList if not file.endswith(".ini")]

        if tipicidad == "Tipico":
            for tipicoUnit in tipicoList:
                for scenarioName in scenariosList:
                    if tipicoUnit == scenarioName:
                        scenarioPath = os.path.join(tipicidadPath, scenarioName)
                        scenarioContent = os.listdir(scenarioPath)
                        if "table.json" in scenarioContent:
                            jsonFile = os.path.join(scenarioPath, "table.json")
                            listJSONPaths.append(jsonFile)    
                            listNames.append(scenarioName)

        elif tipicidad == "Atipico":
            for atipicoUnit in atipicoList:
                for scenarioName in scenariosList:
                    if atipicoUnit == scenarioName:
                        scenarioPath = os.path.join(tipicidadPath, scenarioName)
                        scenarioContent = os.listdir(scenarioPath)
                        if "table.json" in scenarioContent:
                            jsonFile = os.path.join(scenarioPath, "table.json")
                            listJSONPaths.append(jsonFile)
                            listNames.append(scenarioName)      

    doc = Document()
    table = doc.add_table(rows = 1, cols = 9)

    table.cell(0,0).text = "Tipicidad"
    table.cell(0,1).text = "Escenario"
    table.cell(0,1).merge(table.cell(0,2))

    for i, texto in enumerate(["Nodo", "Demora\nPromedio", "Pare\nPromedio", "Cola Máx.\nPromedio", "Número de\nVehículos", "LOS"]):
        table.cell(0,i+3).text = texto

    for i, jsonPath in enumerate(listJSONPaths):
        with open(jsonPath, 'r') as file:
            data = json.load(file)

        for j in range(len(data['nodes_names'])):
            new_row = table.add_row()
            new_row.cells[3].text = data['nodes_names'][j]
            new_row.cells[4].text = str(int(data['nodes_totres'][j][0]))
            new_row.cells[5].text = str(int(data['nodes_totres'][j][1]))
            new_row.cells[6].text = str(int(data['nodes_totres'][j][3]))
            new_row.cells[7].text = str(int(data['nodes_totres'][j][4]))
            new_row.cells[8].text = data['nodes_los'][j]
        
        if i == 0:
            numberNodes = len(data['nodes_names'])
    
    table.cell(1,0).text = "TÍPICO"
    table.cell(1,0).merge(table.cell(numberNodes*8,0))
    table.cell(numberNodes*8+1,0).text = "ATÍPICO"
    table.cell(numberNodes*8+1,0).merge(table.cell(numberNodes*13,0))

    indexNames = 0
    for i in range(1, numberNodes*13+1, 2):
        table.cell(i,1).text = 'Actual'
        table.cell(i,1).merge(table.cell(i+1,1))
        table.cell(i,2).text = listNames[indexNames]
        table.cell(i,2).merge(table.cell(i+1,2))
        indexNames += 1

    _align_content(table)

    finalPath = os.path.join(subareaPath, "Tablas", "resumenTable.docx")
    table.style = 'Table Grid'
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.save(finalPath)
    return finalPath