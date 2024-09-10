import json
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
import os
from docxcompose.composer import Composer
from docxtpl import DocxTemplate
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def get_color_by_los(los: str) -> str:
    colores = {
        "A": "00B050", # Verde
        "B": "B5E6A2", # Verde amarillento
        "C": "FFFF99", # Amarillo
        "D": "FFD961", # Naranja
        "E": "EB844B", # Naranja rojizo
        "F": "FF3B3B", # Rojo
    }
    return colores.get(los, "FFFFFF") #Blanco por defecto

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

def _filling_df(df: pd.DataFrame, data: json, rowDf: int, tipicidad: str, state: str, scenario: str):
    for _, nodeName in enumerate(data["nodes_names"]):
        name = nodeName[0]
        for _, listAcess in data["node_results"][name].items():
            sentidoOrigin = listAcess["Sentido"].split("-")
            nombre = listAcess["Nombre"]+"\n"+sentidoOrigin[0]
            volumen = int(float(listAcess["Numero de Vehiculos"]))
            cola_prom = listAcess["Longitud de Cola Prom."]
            cola_max = listAcess["Longitud de Cola Max."]
            paradas = listAcess["Demora en Paradas Prom."]
            demora = listAcess["Demora Promedio"]
            los = listAcess["LOS"]
            df.loc[rowDf] = [name, nombre, volumen, cola_prom, cola_max, paradas, demora, los, tipicidad.capitalize(), state, scenario]
            rowDf += 1

    return rowDf, df

def _filling_df2(df: pd.DataFrame, data: json, rowDf: int, tipicidad: str, state: str, scenario: str):
    for number, dictDatos in data["vehicle_performance"].items():
        simrun = number
        delayavg = dictDatos["DelayAvg"]
        delaystopavg = dictDatos["DelayStopAvg"]
        speedavg = dictDatos["SpeedAvg"]
        stopsavg = dictDatos["StopsAvg"]
        vehact = dictDatos["VehAct"]
        veharr = dictDatos["VehArr"]
        demandlatent = dictDatos["DemandLatent"]
        df.loc[rowDf] = [state, simrun, delayavg, delaystopavg, speedavg, stopsavg, vehact, veharr, demandlatent, tipicidad.capitalize(), scenario]
        rowDf += 1

    return rowDf, df

def _filling_df3(df: pd.DataFrame, data: json, rowDf: int, tipicidad: str, state: str, scenario: str):
    for number, dictDatos in data["pedestrian_performance"].items():
        simrun = number
        densavg = dictDatos["DensAvg"]
        flowavg = dictDatos["FlowAvg"]
        normspeedavg = dictDatos["NormSpeedAvg"]
        speedavg = dictDatos["SpeedAvg"]
        stoptmavg = dictDatos["StopTmAvg"]
        travtmavg = dictDatos["TravTmAvg"]
        df.loc[rowDf] = [state, simrun, densavg, flowavg, normspeedavg, speedavg, stoptmavg, travtmavg, tipicidad.capitalize(), scenario]
        rowDf += 1
        
    return rowDf, df

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
                    run.font.size = Pt(11)
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

def _read_folders(subareaPath, folderName):
    contentByFolder = {
            "HPM": {
                "Tipico": None,
                "Atipico": None,
            },
            "HPT": {
                "Tipico": None,
                "Atipico": None,
            },
            "HPN": {
                "Tipico": None,
                "Atipico": None,
            },
        }

    folderPath = os.path.join(subareaPath, folderName)

    if os.path.exists(folderPath):
        for tipicidad in ["Tipico", "Atipico"]:
            tipicidadFolder = os.path.join(folderPath, tipicidad)
            tipicidadContent = os.listdir(tipicidadFolder)
            tipicidadContent = [file for file in tipicidadContent if not file.endswith(".ini") and file in ["HPM", "HPT", "HPN"]]

            for scenario in ["HPM", "HPT", "HPN"]:
                jsonPath = os.path.join(tipicidadFolder, scenario, "table.json")
                contentByFolder[scenario][tipicidad] = jsonPath

    return contentByFolder

def result_nodos(jsonPathActual, jsonPathOutputBase, jsonPathOutputProyectado, scenario, tipicidad, df, rowDf) -> None:
    if jsonPathActual:
        with open(jsonPathActual, 'r') as file:
            data = json.load(file)

    if jsonPathOutputBase:
        with open(jsonPathOutputBase, 'r') as file2:
            data2 = json.load(file2)

    if jsonPathOutputProyectado:
        with open(jsonPathOutputProyectado, 'r') as file3:
            data3 = json.load(file3)

    #Creating dataframe
    if jsonPathActual:
        rowDf, df = _filling_df(df, data, rowDf, tipicidad, "Actual", scenario)
    if jsonPathOutputBase:
        rowDf, df = _filling_df(df, data2, rowDf, tipicidad, "Propuesta Base", scenario)
    if jsonPathOutputProyectado:
        rowDf, df = _filling_df(df, data3, rowDf, tipicidad, "Propuesta Proyectada", scenario)
            
    return rowDf, df

def result_vehicular(jsonPathActual, jsonPathOutputBase, jsonPathOutputProyectado, scenario, tipicidad, df, rowDf) -> None:
    if jsonPathActual:
        with open(jsonPathActual, 'r') as file:
            data = json.load(file)

    if jsonPathOutputBase:
        with open(jsonPathOutputBase, 'r') as file2:
            data2 = json.load(file2)

    if jsonPathOutputProyectado:
        with open(jsonPathOutputProyectado, 'r') as file3:
            data3 = json.load(file3)

    #Creating dataframe
    if jsonPathActual:
        rowDf, df = _filling_df2(df, data, rowDf, tipicidad, "Actual", scenario)
    if jsonPathOutputBase:
        rowDf, df = _filling_df2(df, data2, rowDf, tipicidad, "Propuesta Base", scenario)
    if jsonPathOutputProyectado:
        rowDf, df = _filling_df2(df, data3, rowDf, tipicidad, "Propuesta Proyectada", scenario)

    return rowDf, df

def result_peatonal(jsonPathActual, jsonPathOutputBase, jsonPathOutputProyectado, scenario, tipicidad, df, rowDf) -> None:
    if jsonPathActual:
        with open(jsonPathActual, 'r') as file:
            data = json.load(file)

    if jsonPathOutputBase:
        with open(jsonPathOutputBase, 'r') as file2:
            data2 = json.load(file2)

    if jsonPathOutputProyectado:
        with open(jsonPathOutputProyectado, 'r') as file3:
            data3 = json.load(file3)

    #Creating dataframe
    if jsonPathActual:
        rowDf, df = _filling_df3(df, data, rowDf, tipicidad, "Actual", scenario)
    if jsonPathOutputBase:
        rowDf, df = _filling_df3(df, data2, rowDf, tipicidad, "Propuesta Base", scenario)
    if jsonPathOutputProyectado:
        rowDf, df = _filling_df3(df, data3, rowDf, tipicidad, "Propuesta Proyectada", scenario)

    return rowDf, df

def create_tables_nodos(df: pd.DataFrame, tipicidad: str, scenario: str, subareaPath: str):
    doc = Document()
    table = doc.add_table(rows=1, cols=8)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    #Writing headers
    headers = [
        "Inters.\nEscenario", "Nombre de Acceso",
        "Volumen\n(veh/h)", "Long. de Cola Prom.\n(m)", "Long. de Cola Max.\n(m)",
        "Demora por Paradas\n(s/veh)", "Demora\n(s/veh)", "LOS\n(A-F)"
        ]
    
    for i, header in enumerate(headers):
        table.cell(0, i).text = header

    codigo = df["Intersección"].unique().tolist()[0]

    dfActual = df[df["State"] == "Actual"].reset_index(drop=True)
    dfBase = df[df["State"] == "Propuesta Base"].reset_index(drop=True)
    dfProyectado = df[df["State"] == "Propuesta Proyectada"].reset_index(drop=True)

    start = 1
    for j in range(dfActual.shape[0]):
        newRow = table.add_row()
        for i, column in enumerate(["Intersección", "Nombre", "Volumen", "QueueAvg", "QueueMax", "StopDelay", "Delay", "LOS"]):
            if i == 0: continue
            if column in ["QueueAvg", "QueueMax"]:
                newRow.cells[i].text = str(round(float(dfActual.loc[j, column])))
            elif column in ["StopDelay", "Delay"]:
                newRow.cells[i].text = str(round(float(dfActual.loc[j, column]), 1))
            else:
                newRow.cells[i].text = str(dfActual.loc[j, column])
            
            if column == "LOS":
                color_hex = get_color_by_los(str(dfActual.loc[j, column]))
                shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color_hex))
                newRow.cells[i]._element.get_or_add_tcPr().append(shading_elm)

    end = dfActual.shape[0] + start -1

    table.cell(start, 0).text = dfActual.iloc[0, 0] + "\nActual"
    table.cell(start, 0).merge(table.cell(end, 0))

    start = end+1
    for j in range(dfBase.shape[0]):
        newRow = table.add_row()
        for i, column in enumerate(["Intersección", "Nombre", "Volumen", "QueueAvg", "QueueMax", "StopDelay", "Delay", "LOS"]):
            if i == 0: continue
            if column in ["QueueAvg", "QueueMax"]:
                newRow.cells[i].text = str(round(float(dfBase.loc[j, column])))
            elif column in ["StopDelay", "Delay"]:
                newRow.cells[i].text = str(round(float(dfBase.loc[j, column]), 1))
            else:
                newRow.cells[i].text = str(dfBase.loc[j, column])
            if column == "LOS":
                color_hex = get_color_by_los(str(dfBase.loc[j, column]))
                shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color_hex))
                newRow.cells[i]._element.get_or_add_tcPr().append(shading_elm)
    end = dfBase.shape[0] + start - 1

    table.cell(start, 0).text = dfBase.iloc[0,0] + "\nPropuesta\nBase"
    table.cell(start, 0).merge(table.cell(end, 0))

    start = end+1  
    for j in range(dfProyectado.shape[0]):
        newRow = table.add_row()
        for i, column in enumerate(["Intersección", "Nombre", "Volumen", "QueueAvg", "QueueMax", "StopDelay", "Delay", "LOS"]):
            if i == 0: continue
            if column in ["QueueAvg", "QueueMax"]:
                newRow.cells[i].text = str(round(float(dfProyectado.loc[j, column])))
            elif column in ["StopDelay", "Delay"]:
                newRow.cells[i].text = str(round(float(dfProyectado.loc[j, column]), 1))
            else:
                newRow.cells[i].text = str(dfProyectado.loc[j, column])
            if column == "LOS":
                color_hex = get_color_by_los(str(dfProyectado.loc[j, column]))
                shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color_hex))
                newRow.cells[i]._element.get_or_add_tcPr().append(shading_elm)
    end = dfProyectado.shape[0] + start - 1

    table.cell(start, 0).text = dfProyectado.iloc[0,0] + "\nPropuesta\nProyectada"
    table.cell(start, 0).merge(table.cell(end, 0))

    #Aesthetics
    _align_content(table)

    for selected_row in [0]:
        for cell in table.rows[selected_row].cells:
            for paragraph in cell.paragraphs:
                run = paragraph.runs[0]
                run.font.bold = True

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                run = paragraph.runs[0]
                run.font.name = 'Arial Narrow'
                run.font.size = Pt(11)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for id, x in zip([1],[1]):
        for cell in table.columns[id].cells:
            cell.width = Inches(x)

    table.style = 'Table Grid'

    #Saving table
    result_nodo_path = os.path.join(subareaPath, "Tablas", f"nodos_{tipicidad}_{scenario}_{codigo}.docx")
    doc.save(result_nodo_path)
    new_text = f"Resultados de la intersección {codigo} en la {dictNames[scenario].lower()} del día {tipicidad.lower()}"
    resultNode = _generate_table_ref(result_nodo_path, new_text)

    return resultNode

def create_tables_vehicular(df: pd.DataFrame, tipicidad: str, scenario: str, subareaPath: str):
    df = df.reset_index(drop=True)

    doc = Document()
    table = doc.add_table(rows = 1, cols = 9)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    headers = [
        "Escenarios", "Num. Sim.", "Demora\nPromedio", "Demora\nParadas\nPromedio", "Velocidad\nPromedio",
        "Paradas\nPromedio", "Veh.\nAct.", "Veh.\nArr.", "Demora\nLatente"
    ]
    
    for i, header in enumerate(headers):
        table.cell(0, i).text = header

    for j in range(df.shape[0]):                                                                                                                
        newRow = table.add_row()
        for i, column in enumerate([
            "Escenario", "SimRun", "DelayAvg", "DelayStopAvg", "SpeedAvg",
            "StopsAvg", "VehAct", "VehArr", "DemandLatent"
        ]):
            if i == 0: continue
            if column == "SimRun":
                newRow.cells[i].text = str(df.loc[j, column])
            elif column in ["DelayAvg", "DelayStopAvg", "SpeedAvg", "StopsAvg"]:
                newRow.cells[i].text = str(round(float(df.loc[j, column]),2))
            elif column in ["VehAct", "VehArr", "DemandLatent"]:
                newRow.cells[i].text = str(int(float(df.loc[j, column])))
    
    counts = df["Escenario"].value_counts()

    countActual = counts["Actual"]
    countBase = counts["Propuesta Base"]
    countProyectada = counts["Propuesta Proyectada"]

    if  countActual > 0:
        table.cell(1,0).text = "Actual"
        table.cell(1,0).merge(table.cell(12,0))

    if  countBase > 0:
        table.cell(13,0).text = "Propuesta Base"
        table.cell(13,0).merge(table.cell(24,0))

    if countProyectada > 0:
        table.cell(25,0).text = "Propuesta Proyectada"
        table.cell(25,0).merge(table.cell(36,0))

    _align_content(table)

    for selected_row in [0]:
        for cell in table.rows[selected_row].cells:
            for paragraph in cell.paragraphs:
                run = paragraph.runs[0]
                run.font.bold = True

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                try:
                    run = paragraph.runs[0]
                    run.font.name = 'Arial Narrow'
                    run.font.size = Pt(11)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                except:
                    pass

    table.style = 'Table Grid'

    #Saving tableq
    result_vehicular_path = os.path.join(subareaPath, "Tablas", f"vehicular_{tipicidad}_{scenario}.docx")
    doc.save(result_vehicular_path)
    new_text = f"Rendimiento de vehículos de la red en la {dictNames[scenario].lower()} día {tipicidad.lower()}"
    resultVehicular = _generate_table_ref(result_vehicular_path, new_text)

    return resultVehicular

def create_tables_peatonal(df: pd.DataFrame, tipicidad: str, scenario: str, subareaPath: str) -> str:
    df = df.reset_index(drop=True)

    doc = Document()
    table = doc.add_table(rows = 1, cols = 8)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    headers = [
        "Escenarios", "Num. Sim", "Dens. Prom.", "Flujo Prom.", "Vel. Norm. Prom.", "Vel. Prom.", "Tiempo Parada Prom.", "Tiempo Viaje Prom."
    ]

    for i, header in enumerate(headers):
        table.cell(0,i).text = header

    for j in range(df.shape[0]):
        newRow = table.add_row()
        for i, column in enumerate([
            "Escenario", "SimRun", "DensAvg", "FlowAvg", "NormSpeedAvg", "SpeedAvg", "StopTmAvg", "TravTmAvg"
        ]):
            if i == 0: continue
            if column == "SimRun":
                newRow.cells[i].text = str(df.loc[j, column])
            elif column in ["DensAvg", "FlowAvg", "NormSpeedAvg", "SpeedAvg", "StopTmAvg", "TravTmAvg"]:
                try:
                    newRow.cells[i].text = f"{float(df.loc[j, column]):.4f}"
                except:
                    newRow.cells[i].text = str(df.loc[j, column])

    counts = df["Escenario"].value_counts()

    countActual = counts["Actual"]
    countBase = counts["Propuesta Base"]
    countProyectada = counts["Propuesta Proyectada"]

    if  countActual > 0:
        table.cell(1,0).text = "Actual"
        table.cell(1,0).merge(table.cell(12,0))

    if  countBase > 0:
        table.cell(13,0).text = "Propuesta Base"
        table.cell(13,0).merge(table.cell(24,0))

    if countProyectada > 0:
        table.cell(25,0).text = "Propuesta Proyectada"
        table.cell(25,0).merge(table.cell(36,0))

    _align_content(table)

    for selected_row in [0]:
        for cell in table.rows[selected_row].cells:
            for paragraph in cell.paragraphs:
                run = paragraph.runs[0]
                run.font.bold = True

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                try:
                    run = paragraph.runs[0]
                    run.font.name = 'Arial Narrow'
                    run.font.size = Pt(11)
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                except:
                    pass

    table.style = 'Table Grid'

    #Saving tableq
    result_peatonal_path = os.path.join(subareaPath, "Tablas", f"peatonal_{tipicidad}_{scenario}.docx")
    doc.save(result_peatonal_path)
    new_text = f"Rendimiento de peatones de la red en la {dictNames[scenario].lower()} día {tipicidad.lower()}"
    resultPeatonal = _generate_table_ref(result_peatonal_path, new_text)

    return resultPeatonal

def generate_results(subareaPath: str) -> list[str]:
    df = pd.DataFrame(
        columns= [
            "Intersección", "Nombre", "Volumen", "QueueAvg", "QueueMax",
            "StopDelay", "Delay", "LOS", "Tipicidad", "State", "Scenario"
        ]
    )
    
    df2 = pd.DataFrame(
        columns=[
            "Escenario", "SimRun", "DelayAvg", "DelayStopAvg", "SpeedAvg", "StopsAvg", "VehAct", "VehArr", "DemandLatent", "Tipicidad", "Scenario"
        ]
    )

    df3 = pd.DataFrame(
        columns=[
            "Escenario", "SimRun", "DensAvg", "FlowAvg", "NormSpeedAvg", "SpeedAvg", "StopTmAvg", "TravTmAvg", "Tipicidad", "Scenario"
        ]
    )

    #Reading Folders
    actualContent = _read_folders(subareaPath, "Actual")
    outputbaseContent = _read_folders(subareaPath, "Output_Base")
    outputproyectadoContent = _read_folders(subareaPath, "Output_Proyectado")

    rowDf = 0
    rowDf2 = 0
    rowDf3 = 0
    for scenario in ["HPM", "HPT", "HPN"]:
        for tipicidad in ["Tipico", "Atipico"]:
            jsonPathActual = actualContent[scenario][tipicidad]
            jsonPathOutputBase = outputbaseContent[scenario][tipicidad]
            jsonPathOutputProyectado = outputproyectadoContent[scenario][tipicidad]

            rowDf, df = result_nodos(
                jsonPathActual,
                jsonPathOutputBase,
                jsonPathOutputProyectado,
                scenario,
                tipicidad.lower(),
                df,
                rowDf
            )
            
            rowDf2, df2 = result_vehicular(
                jsonPathActual,
                jsonPathOutputBase,
                jsonPathOutputProyectado,
                scenario,
                tipicidad.lower(),
                df2,
                rowDf2
            )

            rowDf3, df3 = result_peatonal(
                jsonPathActual,
                jsonPathOutputBase,
                jsonPathOutputProyectado,
                scenario,
                tipicidad.lower(),
                df3,
                rowDf3
            )

    ########################
    # Resultados por nodos #
    ########################

    results_nodes = {
        "Tipico": {
            "HPM": None,
            "HPT": None,
            "HPN": None,
        },
        "Atipico": {
            "HPM": None,
            "HPT": None,
            "HPN": None,
        },
    }

    intersecciones = df['Intersección'].unique().tolist()
    for tipicidad in ["Tipico", "Atipico"]:
        for turno in ["HPM", "HPT", "HPN"]:
            listPaths = []    
            for ints in intersecciones:
                filtered_df = df[
                    (df["Intersección"] == ints) & 
                    (df["Scenario"] == turno) &
                    (df["Tipicidad"] == tipicidad)
                    ]
                tablaPath = create_tables_nodos(filtered_df, tipicidad, turno, subareaPath)
                listPaths.append(tablaPath)
            
            nodoPath = os.path.join(subareaPath, "Tablas", f"nodeResults_{tipicidad}_{turno}_REF.docx")
            if len(listPaths) > 1:
                filePathMaster = listPaths[0]
                filePathList = listPaths[1:]
                _combine_all_docx(filePathMaster, filePathList, nodoPath)
            else:
                nodoPath = listPaths[0]

            results_nodes[tipicidad][turno] = nodoPath

    ########################################
    # Resultados por rendimiento vehicular #
    ########################################

    results_vehicular = {
        "Tipico": {
            "HPM": None,
            "HPT": None,
            "HPN": None,
        },
        "Atipico": {
            "HPM": None,
            "HPT": None,
            "HPN": None,
        },
    }

    for tipicidad in ["Tipico", "Atipico"]:
        for turno in ["HPM", "HPT", "HPN"]:
            filtered_df = df2[
                (df2["Scenario"] == turno) &
                (df2["Tipicidad"] == tipicidad)
                ]
    
            vehicularResultPathRef = create_tables_vehicular(filtered_df, tipicidad, turno, subareaPath)
            results_vehicular[tipicidad][turno] = vehicularResultPathRef
            
    #######################################
    # Resultados por rendimiento peatonal #
    #######################################

    results_peatonal = {
        "Tipico": {
            "HPM": None,
            "HPT": None,
            "HPN": None,
        },
        "Atipico": {
            "HPM": None,
            "HPT": None,
            "HPN": None,
        },
    }

    for tipicidad in ["Tipico", "Atipico"]:
        for turno in ["HPM", "HPT", "HPN"]:
            filtered_df = df3[
                (df3["Scenario"] == turno) &
                (df3["Tipicidad"] == tipicidad)
                ]
    
            peatonalResultPathRef = create_tables_peatonal(filtered_df, tipicidad, turno, subareaPath)
            results_peatonal[tipicidad][turno] = peatonalResultPathRef

    return results_nodes, results_vehicular, results_peatonal