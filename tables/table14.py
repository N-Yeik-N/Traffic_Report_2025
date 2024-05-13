import os
from pathlib import Path
import re
from tools.matrix import read_matrix
#docx
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell

def create_table14(path_subarea):
    subareaName = os.path.split(path_subarea)[-1]
    subareaID = subareaName[-3:]
    #subarea_id = path_parts[-1]
    actualFolder = Path(path_subarea) / "Actual"


    scenarios_by_tipicidad = {}
    for tipicidad in ["Tipico", "Atipico"]:
        tipFolder = actualFolder / tipicidad
        listScenarios = os.listdir(tipFolder)
        
        listScenarios = [scenario for scenario in listScenarios if os.path.isdir(tipFolder / scenario)]
        scenarios_by_tipicidad[tipicidad] = listScenarios

    matrixes_by_tipicidad = {}
    for tipicidad, listScenarios in scenarios_by_tipicidad.items():
        listMatrixes = []
        for scenario in listScenarios:
            scenarioFolder = actualFolder / tipicidad / scenario
            listExcels = os.listdir(scenarioFolder)
            excelMatrix = [excel for excel in listExcels
                          if excel.endswith(".xlsx") and not excel.startswith("~") and 'Matriz' in excel]
            if len(excelMatrix) > 1: print("Error: Hay más de una matriz en: ", scenario)
            listMatrixes.append(scenarioFolder / excelMatrix[0])

        matrixes_by_tipicidad[tipicidad] = listMatrixes
    
    pattern = r'Matriz-OD_([A-Z]+).xlsx'
    for tipicidad, listExcels in matrixes_by_tipicidad.items():
        for excel in listExcels:
            nameExcel = os.path.split(excel)[-1]
            coincidence = re.search(pattern, nameExcel)
            if coincidence:
                nameScenario = coincidence.group(1)
            else:
                print("Error: No se encontro un nombre correcto de escenario en: ", excel)

            try:
                ORIGINS, DESTINYS, MATRIX, nameScenario, tipicidad = read_matrix(excel)
                table_creation(ORIGINS, DESTINYS, MATRIX)
            except Exception as inst:
                raise inst
            
            #doc_template = DocxTemplate("./templates/template_tablas.docx")
            texto = f"Orígenes - Destinos de la subárea {subareaID} {nameScenario} día {tipicidad.lower()}"
            
def table_creation(ORIGINS, DESTINYS, MATRIX, nameScenario, tipicidad):
    doc = Document()
    table = doc.add_table(rows = 2+len(ORIGINS), cols = 2+len(DESTINYS))
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table.cell(0,2).text = "Destino"
    table.cell(0,2).merge(table.cell(0,1+len(DESTINYS)))
    for elem in DESTINYS:
        table.cell(1,2+DESTINYS.index(elem)).text = elem

    table.cell(2,0).text = "Origen"
    table.cell(2,0).merge(table.cell(1+len(ORIGINS),0))

    for elem in ORIGINS:
        table.cell(2+ORIGINS.index(elem),1).text = elem

    for i, row in enumerate(MATRIX):
        for j, elem in enumerate(row):
            table.cell(i+2,j+2).text = elem

    #Format
    table.cell(2,0).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    _set_vertical_cell_direction(
        table.cell(2,0),
        "btLr",
    )

    #Filling blank spaces:
    table.cell(0,0).text = ""
    table.cell(0,1).text = ""
    table.cell(1,0).text = ""
    table.cell(1,1).text = ""

    for selected_row in [0]:
        for cell in table.rows[selected_row].cells:
            for paragraph in cell.paragraphs:
                run = paragraph.runs[0]
                run.font.bold = True

    for i in range(len(table.rows)):
        cell = table.cell(i,0)
        for paragraph in cell.paragraphs:
            run = paragraph.runs[0]
            run.font.bold = True
    
    for i in range(len(table.columns)):
        if i < 2: continue
        cell_xml_element = table.rows[0].cells[i]._tc
        table_cell_properties = cell_xml_element.get_or_add_tcPr()
        shade_obj = OxmlElement('w:shd')
        shade_obj.set(qn('w:fill'),'B4C6E7')
        table_cell_properties.append(shade_obj)

    for i in range(len(table.rows)):
        if i < 2: continue
        cell = table.cell(i,0)
        cell_xml_element = cell._tc
        table_cell_properties = cell_xml_element.get_or_add_tcPr()
        shade_obj = OxmlElement('w:shd')
        shade_obj.set(qn('w:fill'),'B4C6E7')
        table_cell_properties.append(shade_obj)

    for id in range(table.columns.__len__()):
        for cell in table.columns[id].cells:
            cell.width = Inches(0.25)

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                run = paragraph.runs[0]
                run.font.name = 'Arial Narrow'
                run.font.size = Pt(11)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    cell_exceptions = [(0,0),(0,1),(1,0),(1,1)]
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            if (i,j) in cell_exceptions: continue
            print(i,j)
            cell.style = 'Table Grid'

    if tipicidad == 'tipico': tip = 'T'
    elif tipicidad == 'atipico': tip = 'A'

    final_name = f"table_{nameScenario}_{nameScenario}_{tip}.docx"

    doc.save(final_name)

    return final_name

def _set_vertical_cell_direction(cell: _Cell, direction: str) -> None:
    #direction: tbRl -- Top to bottom, Right to left
    #direction: btLr -- Bottom to top, Left to right
    assert direction in ("tbRl", "btLr")
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    textDirection = OxmlElement('w:textDirection')
    textDirection.set(qn('w:val'), direction)
    tcPr.append(textDirection)