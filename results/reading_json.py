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

def _create_final_tables(subareaPath, finalName, listContent):
    resultTablesPath = os.path.join(subareaPath, "Tablas", finalName)
    if len(listContent) > 1:
        filePathMaster = listContent[0]
        filePathList = listContent[1:]
        _combine_all_docx(filePathMaster, filePathList, resultTablesPath)
    else:
        resultTablesPath = listContent[0]

    return resultTablesPath

NODES_TOTRES = [
    "Nodo",
    "VehDelay\n(Avg,Avg,All)",
    "StopDelay\n(Avg,Avg,All)",
    "QLenMax\n(Avg,Avg)",
    "QLenMax\n(Avg,Total)",
    "Vehs\n(Avg,Total,All)",
    "LOS Value\n(Avg,Total,All)",
]

def read_json(jsonPathActual, jsonPathOutputBase, jsonPathOutputProyectado,subareaPath, name, tipicidad) -> None:

    if jsonPathActual:
        with open(jsonPathActual, 'r') as file:
            data = json.load(file)

    if jsonPathOutputBase:
        with open(jsonPathOutputBase, 'r') as file2:
            data2 = json.load(file2)

    if jsonPathOutputProyectado:
        with open(jsonPathOutputProyectado, 'r') as file3:
            data3 = json.load(file3)

    #################################
    # Resultados vehiculares global #
    #################################

    doc = Document()
    table = doc.add_table(rows = 1, cols = len(data['vehicle_performance']['Avg'])+1+1)

    #Writing headers
    table.cell(0,1).text = "SimRun"
    for i, elem in enumerate(data['vehicle_performance']['Avg']):
        table.cell(0,i+2).text = f'{elem}'

    #Actual
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

    #Output Base
    if jsonPathOutputBase:
        for i,num_runs in enumerate(data2['vehicle_performance']): #Start row = 1
            new_row  = table.add_row()
            new_row.cells[1].text = num_runs
            for j,attribute in enumerate(data2['vehicle_performance'][num_runs]):
                try:
                    new_row.cells[j+2].text = str(int(data2['vehicle_performance'][num_runs][attribute]))
                except TypeError:
                    new_row.cells[j+2].text = str(0)
            last_row_base = i

    if jsonPathOutputProyectado:
        for i,num_runs in enumerate(data3['vehicle_performance']): #Start row = 1
            new_row  = table.add_row()
            new_row.cells[1].text = num_runs
            for j,attribute in enumerate(data3['vehicle_performance'][num_runs]):
                try:
                    new_row.cells[j+2].text = str(int(data3['vehicle_performance'][num_runs][attribute]))
                except TypeError:
                    new_row.cells[j+2].text = str(0)
            last_row_proyectado = i

    last_row_base += last_row_actual
    last_row_proyectado += last_row_base

    table.cell(0,0).text = "Escenarios"
    #Actual
    table.cell(1,0).text = "Actual"
    table.cell(1,0).merge(table.cell(last_row_actual+1,0))
    #Output Base
    table.cell(last_row_actual+2,0).text = "Propuesto Base"
    table.cell(last_row_actual+2,0).merge(table.cell(last_row_base+2,0))
    #Output Proyectado
    table.cell(last_row_base+3,0).text = "Propuesto Proyectado"
    table.cell(last_row_base+3,0).merge(table.cell(last_row_proyectado+3,0))

    _align_content(table)

    table.style = 'Table Grid'

    vehicularResultPath = os.path.join(subareaPath, "Tablas", f"vehicularResults_{name}_{unidecode(tipicidad)}.docx")
    doc.save(vehicularResultPath)
    new_text = f"Rendimiento de vehículos de la red en la {dictNames[name]} día {tipicidad.lower()}"
    vehicularResultPathRef = _generate_table_ref(vehicularResultPath, new_text)

    ################################
    # Resultados peatonales global #
    ################################

    doc = Document()
    table = doc.add_table(rows=1, cols=len(data['pedestrian_performance']['Avg'])+1+1)

    #Writing headers
    table.cell(0,1).text = "SimRun"
    for i, elem in enumerate(data['pedestrian_performance']['Avg']):
        table.cell(0,i+2).text = f'{elem}'
    
    #Actual
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

    #Output Base
    if jsonPathOutputBase:
        for i, num_runs in enumerate(data2['pedestrian_performance']):
            new_row = table.add_row()
            new_row.cells[1].text = num_runs
            for j, attribute in enumerate(data2['pedestrian_performance'][num_runs]):
                try:
                    new_row.cells[j+2].text = str(round(float(data2['pedestrian_performance'][num_runs][attribute]),4))
                except TypeError:
                    new_row.cells[j+2].text = str(0)
            last_row_base = i

    #Output Proyectado
    if jsonPathOutputProyectado:
        for i, num_runs in enumerate(data3['pedestrian_performance']):
            new_row = table.add_row()
            new_row.cells[1].text = num_runs
            for j, attribute in enumerate(data3['pedestrian_performance'][num_runs]):
                try:
                    new_row.cells[j+2].text = str(round(float(data3['pedestrian_performance'][num_runs][attribute]),4))
                except TypeError:
                    new_row.cells[j+2].text = str(0)
            last_row_proyectado = i   

    last_row_base += last_row_actual
    last_row_proyectado += last_row_base

    table.cell(0,0).text = "Escenarios"
    #Actual
    table.cell(1,0).text = "Actual"
    table.cell(1,0).merge(table.cell(last_row_actual+1, 0))
    #Output Base
    table.cell(last_row_actual+2,0).text = "Propuesto Base"
    table.cell(last_row_actual+2,0).merge(table.cell(last_row_base+2,0))
    #Output Proyectado
    table.cell(last_row_base+3,0).text = "Propuesto Proyectado"
    table.cell(last_row_base+3,0).merge(table.cell(last_row_proyectado+3,0))

    _align_content(table)

    table.style = 'Table Grid'
    pedestrianResultPath = os.path.join(subareaPath, "Tablas", f"pedestrianResults_{name}_{unidecode(tipicidad)}.docx")
    doc.save(pedestrianResultPath)
    new_text = f"Rendimiento de peatones de la red en la {dictNames[name]} día {tipicidad.lower()}"
    pedestrianResultPathRef = _generate_table_ref(pedestrianResultPath, new_text)

    table.style = 'Table Grid'

    ##########################################
    # Resultados de rendimiento de los nodos #
    ##########################################

    #********** ACTUAL VALUES **********

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

    nodeResultActualPath = os.path.join(subareaPath, "Tablas", f"nodeResults_{name}_{unidecode(tipicidad)}_actual.docx")
    doc.save(nodeResultActualPath)
    new_text = f"Resultados actuales de los nodos en la {dictNames[name]} del día {tipicidad.lower()}"
    nodeResultActualPathRef = _generate_table_ref(nodeResultActualPath, new_text)

    #********** OUTPUT BASE VALUES **********

    #Computing number of columns
    jumpRows = []
    nodeNames = []
    for nodeName in data2["node_results"]:
        jumpRows.append(len(data2["node_results"][nodeName]))
        nodeNames.append(nodeName)
    
    doc = Document()
    table = doc.add_table(rows = 1, cols = 8)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table.cell(0,0).text = "Indicadores de Evaluación"
    table.cell(0,0).merge(table.cell(0,1))

    for i, indicator in enumerate(["Volumen\n(veh)", "Long. de Cola Prom.\n(m)", "Long. de Cola Máx.\n(m)", "Demora por Paradas\n(s/veh)", "Demora\n(s/veh)", "LOS\n(A-F)"]):
        table.cell(0, i+2).text = indicator

    for nodeName in data2["node_results"]:
        for _, indicatorList in data2["node_results"][nodeName].items():
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

    nodeResultBasePath = os.path.join(subareaPath, "Tablas", f"nodeResults_{name}_{unidecode(tipicidad)}_base.docx")
    doc.save(nodeResultBasePath)
    new_text = f"Resultados propuestos base de los nodos en la {dictNames[name]} del día {tipicidad.lower()}"
    nodeResultBasePathRef = _generate_table_ref(nodeResultBasePath, new_text)

    #********** OUTPUT PROYECTADO VALUES **********

    #Computing number of columns
    jumpRows = []
    nodeNames = []
    for nodeName in data3["node_results"]:
        jumpRows.append(len(data3["node_results"][nodeName]))
        nodeNames.append(nodeName)
    
    doc = Document()
    table = doc.add_table(rows = 1, cols = 8)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table.cell(0,0).text = "Indicadores de Evaluación"
    table.cell(0,0).merge(table.cell(0,1))

    for i, indicator in enumerate(["Volumen\n(veh)", "Long. de Cola Prom.\n(m)", "Long. de Cola Máx.\n(m)", "Demora por Paradas\n(s/veh)", "Demora\n(s/veh)", "LOS\n(A-F)"]):
        table.cell(0, i+2).text = indicator

    for nodeName in data3["node_results"]:
        for _, indicatorList in data3["node_results"][nodeName].items():
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

    nodeResultProyectadoPath = os.path.join(subareaPath, "Tablas", f"nodeResults_{name}_{unidecode(tipicidad)}_proyectado.docx")
    doc.save(nodeResultProyectadoPath)
    new_text = f"Resultados propuestos proyectados de los nodos en la {dictNames[name]} del día {tipicidad.lower()}"
    nodeResultProyectadoPathRef = _generate_table_ref(nodeResultProyectadoPath, new_text)

    ######################
    # RESULTADOS FINALES #
    ######################

    return nodeResultActualPathRef, nodeResultBasePathRef, nodeResultProyectadoPathRef, pedestrianResultPathRef, vehicularResultPathRef

def generate_results(subareaPath) -> None:
    actualPath = os.path.join(subareaPath, "Actual")
    outputBasePath = os.path.join(subareaPath, "Output_Base")
    outputProyectadoPath = os.path.join(subareaPath, "Output_Proyectado")
    listWords = {
        "Tipico": {
            "Vehicular": {
                "Nodo": [],
                "Red": [],
            },
            "Peatonal": {
                "Red": [],
            }
        },
        "Atipico": {
            "Vehicular": {
                "Nodo": [],
                "Red": [],
            },
            "Peatonal": {
                "Red": [],
            }
        }
    }
    for tipicidad in ["Tipico", "Atipico"]:
        #Actual
        tipicidadFolderActual = os.path.join(actualPath, tipicidad)
        tipicidadContentActual = os.listdir(tipicidadFolderActual)
        tipicidadContentActual = [file for file in tipicidadContentActual if not file.endswith(".ini") and file in ["HPM", "HPT", "HPN"]]

        #Output Base
        tipicidadFolderOutputBase = os.path.join(outputBasePath, tipicidad)
        if os.path.exists(tipicidadFolderOutputBase):
            tipicidadContentOutputBase = os.listdir(tipicidadFolderOutputBase)
            tipicidadContentOutputBase = [file for file in tipicidadContentOutputBase if not file.endswith(".ini") and file in ["HPM", "HPT", "HPN"]]
        else: tipicidadContentOutputBase = None

        #Output Projected
        tipicidadFolderOutputProyectado = os.path.join(outputProyectadoPath, tipicidad)
        if os.path.exists(tipicidadFolderOutputProyectado):
            tipicidadContentOutputProyectado = os.listdir(tipicidadFolderOutputProyectado)
            tipicidadContentOutputProyectado = [file for file in tipicidadContentOutputProyectado if not file.endswith(".ini") and file in ["HPM", "HPT", "HPN"]]
        else: tipicidadContentOutputProyectado = None

        if tipicidadContentOutputBase != None and tipicidadContentOutputProyectado != None:
            for scenarioActual, scenarioOutputBase, scenarioOutputProyectado in zip(tipicidadContentActual, tipicidadContentOutputBase, tipicidadContentOutputProyectado):
                checkActual = False
                checkOutputBase = False
                checkOutputProyectado = False

                #Actual 
                scenarioPathActual = os.path.join(tipicidadFolderActual, scenarioActual)
                scenarioContentActual = os.listdir(scenarioPathActual)
                if "table.json" in scenarioContentActual:
                    jsonPathActual = os.path.join(scenarioPathActual, "table.json")
                    checkActual = True

                #Output Base
                scenarioPathOutputBase = os.path.join(tipicidadFolderOutputBase, scenarioOutputBase)
                scenarioContentOutputBase = os.listdir(scenarioPathOutputBase)
                if "table.json" in scenarioContentOutputBase:
                    jsonPathOutputBase = os.path.join(scenarioPathOutputBase, "table.json")
                    checkOutputBase = True

                #Output Proyectado
                scenarioPathOutputProyectado = os.path.join(tipicidadFolderOutputProyectado, scenarioOutputProyectado)
                scenarioContentOutputProyectado = os.listdir(scenarioPathOutputProyectado)
                if "table.json" in scenarioContentOutputProyectado:
                    jsonPathOutputProyectado = os.path.join(scenarioPathOutputProyectado, "table.json")
                    checkOutputProyectado = True

                if checkActual and checkOutputBase and checkOutputProyectado:
                    if tipicidad == "Tipico": textTipicidad = "típico"
                    elif tipicidad == "Atipico": textTipicidad = "atípico"
                    nodeResultActualPathRef, nodeResultBasePathRef, nodeResultProyectadoPathRef, pedestrianResultPathRef, vehicularResultPathRef = read_json(
                        jsonPathActual,
                        jsonPathOutputBase,
                        jsonPathOutputProyectado,
                        subareaPath,
                        scenarioActual,
                        textTipicidad)
                    #Output Base
                    listWords[tipicidad]["Vehicular"]["Nodo"].extend([nodeResultActualPathRef, nodeResultBasePathRef, nodeResultProyectadoPathRef])
                    listWords[tipicidad]["Peatonal"]["Red"].extend([pedestrianResultPathRef])
                    listWords[tipicidad]["Vehicular"]["Red"].extend([vehicularResultPathRef])
                    #listWords.extend([nodeResultActualPathRef, nodeResultBasePathRef, nodeResultProyectadoPathRef, pedestrianResultPathRef, vehicularResultPathRef])


        elif tipicidadContentOutputBase != None:
            for scenarioActual, scenarioOutputBase in zip(tipicidadContentActual, tipicidadContentOutputBase):
                checkActual = False
                checkOutputBase = False
                #Actual 
                scenarioPathActual = os.path.join(tipicidadFolderActual, scenarioActual)
                scenarioContentActual = os.listdir(scenarioPathActual)
                if "table.json" in scenarioContentActual:
                    jsonPathActual = os.path.join(scenarioPathActual, "table.json")
                    checkActual = True

                #Output Base
                scenarioPathOutputBase = os.path.join(tipicidadFolderOutputBase, scenarioOutputBase)
                scenarioContentOutputBase = os.listdir(scenarioPathOutputBase)
                if "table.json" in scenarioContentOutputBase:
                    jsonPathOutputBase = os.path.join(scenarioPathOutputBase, "table.json")
                    checkOutputBase = True

                if checkActual and checkOutputBase:
                    if tipicidad == "Tipico": textTipicidad = "típico"
                    elif tipicidad == "Atipico": textTipicidad = "atípico"
                    nodeResultActualPathRef, nodeResultBasePathRef, nodeResultProyectadoPathRef, pedestrianResultPathRef, vehicularResultPathRef = read_json(
                        jsonPathActual,
                        jsonPathOutputBase,
                        None,
                        subareaPath,
                        scenarioActual,
                        textTipicidad)
                    #Output Base
                    listWords[tipicidad]["Vehicular"]["Nodo"].extend([nodeResultActualPathRef, nodeResultBasePathRef, nodeResultProyectadoPathRef])
                    listWords[tipicidad]["Peatonal"]["Red"].extend([pedestrianResultPathRef])
                    listWords[tipicidad]["Vehicular"]["Red"].extend([vehicularResultPathRef])
                    #listWords.extend([nodeResultActualPathRef, nodeResultBasePathRef, pedestrianResultPathRef, vehicularResultPathRef])


        else:
            for scenarioActual in tipicidadContentActual:
                checkActual = False
                checkOutputBase = False
                #Actual 
                scenarioPathActual = os.path.join(tipicidadFolderActual, scenarioActual)
                scenarioContentActual = os.listdir(scenarioPathActual)
                if "table.json" in scenarioContentActual:
                    jsonPathActual = os.path.join(scenarioPathActual, "table.json")
                    checkActual = True

                if checkActual:
                    if tipicidad == "Tipico": textTipicidad = "típico"
                    elif tipicidad == "Atipico": textTipicidad = "atípico"
                    nodeResultActualPathRef, nodeResultBasePathRef, nodeResultProyectadoPathRef, pedestrianResultPathRef, vehicularResultPathRef = read_json(
                        jsonPathActual,
                        None,
                        None,
                        subareaPath,
                        scenarioActual,
                        textTipicidad)
                    #Output Base
                    listWords[tipicidad]["Vehicular"]["Nodo"].extend([nodeResultActualPathRef, nodeResultBasePathRef, nodeResultProyectadoPathRef])
                    listWords[tipicidad]["Peatonal"]["Red"].extend([pedestrianResultPathRef])
                    listWords[tipicidad]["Vehicular"]["Red"].extend([vehicularResultPathRef])
                    #listWords.extend([nodeResultActualPathRef, pedestrianResultPathRef, vehicularResultPathRef])

    resultsPaths = {
        "Tipico": {
            "Vehicular": {
                "Nodo": None,
                "Red": None,
            },
            "Peatonal": {
                "Red": None,
            }
        },
        "Atipico": {
            "Vehicular": {
                "Nodo": None,
                "Red": None,
            },
            "Peatonal": {
                "Red": None,
            }
        }
    }
    #print(listWords)
    for tipicidad, typicalContent in listWords.items():
        for vehicleType, vehicleContent in typicalContent.items():
            for contentType, content in vehicleContent.items():
                resultTablePath = _create_final_tables(subareaPath, f"results_{contentType}_{vehicleType}_{tipicidad}.docx", content)
                resultsPaths[tipicidad][vehicleType][contentType] = resultTablePath
        
    return resultsPaths

# if __name__ == "__main__":
#     subareaPath = r"C:\Users\dacan\OneDrive\Desktop\PRUEBAS\Maxima Entropia\04 Proyecto Universitaria (37 Int. - 19 SA)\6. Sub Area Vissim\Sub Area 001"
#     resultsPaths = generate_results(subareaPath)
#     print(resultsPaths)