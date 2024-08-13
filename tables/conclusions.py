import os
import json

from docxcompose.composer import Composer
from docx import Document
from docxtpl import DocxTemplate

tipicoList = ["HPM", "HPT", "HPN"]
atipicoList = ["HPM", "HPT", "HPN"]

def _combine_all_docx(filePathMaster, filePathsList, finalPath) -> None:
    number_of_sections = len(filePathsList)
    master = Document(filePathMaster)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document(filePathsList[i])
        composer.append(doc_temp)
    composer.save(finalPath)

def _read_json(subareaPath):
    actualPath = os.path.join(subareaPath, "Actual")
    basePath = os.path.join(subareaPath, "Output_Base")
    proyectadoPath = os.path.join(subareaPath, "Output_Proyectado")

    pathsByTipicidad = {
            "Actual": {
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
            },
            "Propuesto": {
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
            },
            "Proyectado": {
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
        }

    for tipicidad in ["Tipico", "Atipico"]:
        #Actual
        tipicidadPathActual = os.path.join(actualPath, tipicidad)
        scenariosListActual = os.listdir(tipicidadPathActual)
        scenariosListActual = [file for file in scenariosListActual if not file.endswith(".ini") and file in ["HPM", "HPT", "HPN"]]

        #Output Base
        tipicidadPathBase = os.path.join(basePath, tipicidad)
        scenariosListBase = os.listdir(tipicidadPathBase)
        scenariosListBase = [file for file in scenariosListBase if not file.endswith(".ini") and file in ["HPM", "HPT", "HPN"]]

        #Output Proyectado
        tipicidadPathProyectado = os.path.join(proyectadoPath, tipicidad)
        scenariosListProyectado = os.listdir(tipicidadPathProyectado)
        scenariosListProyectado = [file for file in scenariosListProyectado if not file.endswith(".ini") and file in ["HPM", "HPT", "HPN"]]

        if tipicidad == "Tipico":
            for tipicoUnit in tipicoList:
                for i in range(len(scenariosListActual)):
                    if tipicoUnit == scenariosListActual[i]:

                        scenarioPathActual = os.path.join(tipicidadPathActual, scenariosListActual[i])
                        scenarioContentActual = os.listdir(scenarioPathActual)
                        if "table.json" in scenarioContentActual:
                            jsonFileActual = os.path.join(scenarioPathActual, "table.json")
                            pathsByTipicidad["Actual"][tipicidad][tipicoUnit] = jsonFileActual                            #listJSONPathsActual.append(jsonFileActual)    
                            #listNames.append(scenariosListActual[i])

                        scenarioPathBase = os.path.join(tipicidadPathBase, scenariosListBase[i])
                        scenarioContentBase = os.listdir(scenarioPathBase)
                        if "table.json" in scenarioContentBase:
                            jsonFileBase = os.path.join(scenarioPathBase, "table.json")
                            #listJSONPathsBase.append(jsonFileBase)
                            pathsByTipicidad["Propuesto"][tipicidad][tipicoUnit] = jsonFileBase

                        scenarioPathProyectado = os.path.join(tipicidadPathProyectado, scenariosListProyectado[i])
                        scenarioContentProyectado = os.listdir(scenarioPathProyectado)
                        if "table.json" in scenarioContentProyectado:
                            jsonFileProyectado = os.path.join(scenarioPathProyectado, "table.json")
                            #listJSONPathsProyectado.append(jsonFileProyectado)
                            pathsByTipicidad["Proyectado"][tipicidad][tipicoUnit] = jsonFileProyectado

        elif tipicidad == "Atipico":
            for tipicoUnit in tipicoList:
                for i in range(len(scenariosListActual)):
                    if tipicoUnit == scenariosListActual[i]:
                        
                        scenarioPathActual = os.path.join(tipicidadPathActual, scenariosListActual[i])
                        scenarioContentActual = os.listdir(scenarioPathActual)
                        if "table.json" in scenarioContentActual:
                            jsonFileActual = os.path.join(scenarioPathActual, "table.json")
                            #listJSONPathsActual.append(jsonFileActual)    
                            #listNames.append(scenariosListActual[i])
                            pathsByTipicidad["Actual"][tipicidad][tipicoUnit] = jsonFileActual

                        scenarioPathBase = os.path.join(tipicidadPathBase, scenariosListBase[i])
                        scenarioContentBase = os.listdir(scenarioPathBase)
                        if "table.json" in scenarioContentBase:
                            jsonFileBase = os.path.join(scenarioPathBase, "table.json")
                            #listJSONPathsBase.append(jsonFileBase)
                            pathsByTipicidad["Propuesto"][tipicidad][tipicoUnit] = jsonFileBase 

                        scenarioPathProyectado = os.path.join(tipicidadPathProyectado, scenariosListProyectado[i])
                        scenarioContentProyectado = os.listdir(scenarioPathProyectado)
                        if "table.json" in scenarioContentProyectado:
                            jsonFileProyectado = os.path.join(scenarioPathProyectado, "table.json")
                            #listJSONPathsProyectado.append(jsonFileProyectado)
                            pathsByTipicidad["Proyectado"][tipicidad][tipicoUnit] = jsonFileProyectado

    return pathsByTipicidad

def get_conclusions(subareaPath: str):
    pathsByTipicidad = _read_json(subareaPath)
    
    count = 0
    listConclusionsLOS = []
    listConclusionsQUEUE = []
    for tipicidad in ["Tipico", "Atipico"]:
        for scenario in ["HPM", "HPT", "HPN"]:
            with open(pathsByTipicidad["Actual"][tipicidad][scenario], 'r') as file:
                dataActual = json.load(file)

            with open(pathsByTipicidad["Propuesto"][tipicidad][scenario], 'r') as file:
                dataBase = json.load(file)

            with open(pathsByTipicidad["Proyectado"][tipicidad][scenario], 'r') as file:
                dataProyectado = json.load(file)

            for i, [code] in enumerate(dataActual["nodes_names"]):

                #Párrafos de nivel de servicio
                losActual = dataActual["nodes_los"][i]
                losPropuesto = dataBase["nodes_los"][i]
                losProyectado = dataProyectado["nodes_los"][i]
                delaysActual = f'{dataActual["nodes_totres"][i][0]:.1f}'
                delaysPropuesto = f'{dataBase["nodes_totres"][i][0]:.1f}'
                delaysProyectado = f'{dataProyectado["nodes_totres"][i][0]:.1f}'

                if scenario == "HPM": peakhour = "hora punta mañana"
                elif scenario == "HPT": peakhour = "hora punta tarde"
                elif scenario == "HPN": peakhour = "hora punta noche"

                checkSignalBefore = dataActual["nodes_signalized"][i][0]
                checkSignalAfter = dataBase["nodes_signalized"][i][0] 

                if checkSignalBefore == "NONSIGNALIZED" and checkSignalAfter == "SIGNALIZED":
                    signalCheckTxt = "En el escenario actual no existían semáforos, pero en el escenario propuesto sí existen."
                elif checkSignalBefore == "SIGNALIZED" and checkSignalAfter == "SIGNALIZED":
                    signalCheckTxt = ""

                docTemplate = DocxTemplate("./templates/template_lista6.docx")
                docTemplate.render({
                    "tipicidad": tipicidad,
                    "peakhour": peakhour,
                    "code": code,
                    "los_actual": losActual,
                    "los_propuesto": losPropuesto,
                    "los_proyectado": losProyectado,
                    "delay_actual": delaysActual,
                    "delay_propuesto": delaysPropuesto,
                    "delay_proyectado": delaysProyectado,
                    "signal_check": signalCheckTxt,
                })
                finalPath = os.path.join(subareaPath, "Tablas", f"conclusions_los_{count}.docx")
                docTemplate.save(finalPath)
                listConclusionsLOS.append(finalPath)

                #Párrafos de cola
                actualQueue = f'{dataActual["nodes_totres"][i][3]:.0f}'
                propQueue = f'{dataBase["nodes_totres"][i][3]:.0f}'

                docTemplate = DocxTemplate("./templates/template_lista5.docx")
                docTemplate.render({
                    "codinterseccion": code,
                    "tipicidad": tipicidad,
                    "turno": peakhour,
                    "actual_queue_max": actualQueue,
                    "propuesto_queue_max": propQueue,
                })
                finalPath = os.path.join(subareaPath, "Tablas", f"conclusion_queue_{count}.docx")
                docTemplate.save(finalPath)
                listConclusionsQUEUE.append(finalPath)

                count += 1

    conclusionLosPath = os.path.join(
        subareaPath, "Tablas", "conclusion_los.docx"
    )
    filePathMaster = listConclusionsLOS[0]
    filePathList = listConclusionsLOS[1:]
    _combine_all_docx(filePathMaster, filePathList, conclusionLosPath)

    conclusionQueuePath = os.path.join(
        subareaPath, "Tablas", "conclusion_queue.docx"
    )
    filePathMaster = listConclusionsQUEUE[0]
    filePathList = listConclusionsQUEUE[1:]
    _combine_all_docx(filePathMaster, filePathList, conclusionQueuePath)

    return conclusionLosPath, conclusionQueuePath