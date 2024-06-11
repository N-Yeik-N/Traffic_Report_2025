import json
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
import os
from docxcompose.composer import Composer
from docxtpl import DocxTemplate
from unidecode import unidecode
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

dictNames = {
    "HVMAD": "Hora Valle Madrugada",
    "HPMAD": "Hora Punta Madrugada",
    "HPM": "Hora Punta Mañana",
    "HVM": "Hora Valle Mañana",
    "HPT": "Hora Punta Tarde",
    "HVT": "Hora Valle Tarde",
    "HPN": "Hora Punta Noche",
    "HVN": "Hora Valle Noche",
}

def _combine_all_docx(filePathMaster, filePathsList, finalPath) -> None:
    number_of_sections = len(filePathsList)
    master = Document(filePathMaster)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document(filePathsList[i])
        composer.append(doc_temp)
    composer.save(finalPath)

def _generate_table_ref(objectResultPath, new_text):
    doc_template = DocxTemplate("./templates/template_tablas4.docx")
    new_table = doc_template.new_subdoc(objectResultPath)
    doc_template.render({
        "texto": new_text,
        "tabla": new_table,
    })
    objectResultPathRef = objectResultPath[:-5] + '_REF.docx'
    doc_template.save(objectResultPathRef)
    return objectResultPathRef

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

NODES_TOTRES = [
    "Nodo",
    "VehDelay\n(Avg,Avg,All)",
    "StopDelay\n(Avg,Avg,All)",
    "QLenMax\n(Avg,Avg)",
    "QLenMax\n(Avg,Total)",
    "Vehs\n(Avg,Total,All)",
    "LOS Value\n(Avg,Total,All)",
]

def read_json(jsonPathActual, subareaPath, name, tipicidad) -> None:

    if jsonPathActual:
        with open(jsonPathActual, 'r') as file:
            data = json.load(file)

    #Tabla: resultados vehiculares RED
    doc = Document()
    table = doc.add_table(rows = 1, cols = len(data['vehicle_performance']['Avg'])+1+1)

    table.cell(0,1).text = "SimRun"
    for i, elem in enumerate(data['vehicle_performance']['Avg']):
        table.cell(0,i+2).text = f'{elem}'

    if jsonPathActual:
        for i,num_runs in enumerate(data['vehicle_performance']): #Start row = 1
            new_row  = table.add_row()
            new_row.cells[1].text = num_runs
            for j,attribute in enumerate(data['vehicle_performance'][num_runs]):
                try:
                    new_row.cells[j+2].text = str(int(data['vehicle_performance'][num_runs][attribute]))
                except TypeError:
                    new_row.cells[j+2].text = str(0)
            last_row_actual = i

    table.cell(0,0).text = "Escenarios"
    table.cell(1,0).text = "Actual"
    table.cell(1,0).merge(table.cell(last_row_actual+1,0))

    _align_content(table)

    table.style = 'Table Grid'

    vehicularResultPath = os.path.join(subareaPath, "Tablas", f"vehicularResults_{name}_{unidecode(tipicidad)}.docx")
    doc.save(vehicularResultPath)
    new_text = f"Rendimiento de vehículos de la red en la {dictNames[name]} día {tipicidad.lower()}"
    vehicularResultPathRef = _generate_table_ref(vehicularResultPath, new_text)

    #Tabla: resultados peatonales RED
    doc = Document()
    table = doc.add_table(rows=1, cols=len(data['pedestrian_performance']['Avg'])+1+1)

    table.cell(0,1).text = "SimRun"
    for i, elem in enumerate(data['pedestrian_performance']['Avg']):
        table.cell(0,i+2).text = f'{elem}'
    
    if jsonPathActual:
        for i, num_runs in enumerate(data['pedestrian_performance']):
            new_row = table.add_row()
            new_row.cells[1].text = num_runs
            for j, attribute in enumerate(data['pedestrian_performance'][num_runs]):
                try:
                    new_row.cells[j+2].text = str(round(float(data['pedestrian_performance'][num_runs][attribute]),4))
                except TypeError:
                    new_row.cells[j+2].text = str(0)
            last_row_actual = i

    table.cell(0,0).text = "Escenarios"
    table.cell(1,0).text = "Actual"
    table.cell(1,0).merge(table.cell(last_row_actual+1, 0))

    _align_content(table)

    table.style = 'Table Grid'
    pedestrianResultPath = os.path.join(subareaPath, "Tablas", f"pedestrianResults_{name}_{unidecode(tipicidad)}.docx")
    doc.save(pedestrianResultPath)
    new_text = f"Rendimiento de peatones de la red en la {dictNames[name]} día {tipicidad.lower()}"
    pedestrianResultPathRef = _generate_table_ref(pedestrianResultPath, new_text)

    table.style = 'Table Grid'

    #Tabla: Resultados de rendimiento de los nodos
    if jsonPathActual:
        with open(jsonPathActual, 'r') as file:
            data = json.load(file)

    #Computing number of columns
    jumpRows = []
    nodeNames = []
    for nodeName in data["node_results"]:
        jumpRows.append(len(data["node_results"][nodeName]))
        nodeNames.append(nodeName)
    
    doc = Document()
    table = doc.add_table(rows = 1, cols = 8)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table.cell(0,0).text = "Indicadores de Evaluación"
    table.cell(0,0).merge(table.cell(0,1))

    for i, indicator in enumerate(["Volumen\n(veh)", "Long. de Cola Prom.\n(m)", "Long. de Cola Máx.\n(m)", "Demora por Paradas\n(s/veh)", "Demora\n(s/veh)", "LOS\n(A-F)"]):
        table.cell(0, i+2).text = indicator

    for nodeName in data["node_results"]:
        for _, indicatorList in data["node_results"][nodeName].items():
            newRow = table.add_row()
            sentido = indicatorList["Sentido"].split("-")[0]
            nameTable = indicatorList["Nombre"]+f"\n({sentido})"
            newRow.cells[1].text = nameTable
            newRow.cells[2].text = str(int(float(indicatorList['Numero de Vehiculos'])))
            newRow.cells[3].text = indicatorList['Longitud de Cola Prom.']
            newRow.cells[4].text = indicatorList['Longitud de Cola Max.']
            newRow.cells[5].text = indicatorList['Demora en Paradas Prom.']
            newRow.cells[6].text = indicatorList['Demora Promedio']
            newRow.cells[7].text = indicatorList['LOS']

    rowNum = 1
    for nodeName, jump in zip(nodeNames, jumpRows):
        if rowNum == 1:
            table.cell(rowNum, 0).text = nodeName
            table.cell(rowNum, 0).merge(table.cell(rowNum+jump-1, 0))
            rowNum += jump
        else:
            table.cell(rowNum, 0).text = nodeName
            table.cell(rowNum, 0).merge(table.cell(rowNum+jump-1, 0))
            rowNum += jump

    #Bond
    for i in range(len(table.columns)):
        cell = table.cell(0,i)
        for paragraph in cell.paragraphs:
            run = paragraph.runs[0]
            run.font.bold = True

    _align_content(table)

    #Width
    for idColumn, widthColumn in zip([0, 1, 2, 6, 7], [1.2, 4.0, 1.5, 1.5, 1.5]):
        for cell in table.columns[idColumn].cells:
            cell.width = Cm(widthColumn)

    table.style = 'Table Grid'

    nodeResultPath = os.path.join(subareaPath, "Tablas", f"nodeResults_{name}_{unidecode(tipicidad)}.docx")
    doc.save(nodeResultPath)
    new_text = f"Resultados de los nodos en la {dictNames[name]} día {tipicidad.lower()}"
    nodeResultPathRef = _generate_table_ref(nodeResultPath, new_text)

    return nodeResultPathRef, pedestrianResultPathRef, vehicularResultPathRef

def generate_results(subareaPath) -> None:
    actualPath = os.path.join(subareaPath, "Actual")
    listWords = []
    for tipicidad in ["Tipico", "Atipico"]:
        tipicidadFolder = os.path.join(actualPath, tipicidad)
        tipicidadContent = os.listdir(tipicidadFolder)
        tipicidadContent = [file for file in tipicidadContent if not file.endswith(".ini")]

        for scenario in tipicidadContent:
            if not scenario in ["HPM", "HPT", "HPN"]: continue
            scenarioPath = os.path.join(tipicidadFolder, scenario)
            scenarioContent = os.listdir(scenarioPath)
            if "table.json" in scenarioContent:
                jsonPathActual = os.path.join(scenarioPath, "table.json")
                if tipicidad == "Tipico": textTipicidad = "típico"
                elif tipicidad == "Atipico": textTipicidad = "atípico"
                nodeResultPathRef, pedestrianResultPathRef, vehicularResultPathRef = read_json(jsonPathActual, subareaPath, scenario, textTipicidad)
                listWords.extend([nodeResultPathRef, pedestrianResultPathRef, vehicularResultPathRef])

    resultTablesPath = os.path.join(subareaPath, "Tablas", "0_resultTables.docx")
    filePathMaster = listWords[0]
    filePathList = listWords[1:]
    _combine_all_docx(filePathMaster, filePathList, resultTablesPath)
        
    return resultTablesPath
