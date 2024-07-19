import os
import xml.etree.ElementTree as ET
import re
import pandas as pd

#Doc
from docxtpl import DocxTemplate, InlineImage
from docxcompose.composer import Composer
from docx import Document
from docx.shared import Inches
from unidecode import unidecode

def _combine_all_docx(filePathMaster, filePathsList, finalPath) -> None:
    number_of_sections = len(filePathsList)
    master = Document(filePathMaster)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document(filePathsList[i])
        composer.append(doc_temp)

    composer.save(finalPath)

def get_sigs_actual(
        subareaPath,
        typesig: str #Actual / Propuesto
        ) -> None:

    #Getting list of node codes
    listFiles = os.listdir(subareaPath)
    skeletonFile = [file for file in listFiles if file.endswith(".inpx")][0]
    skeletonPath = os.path.join(subareaPath, skeletonFile)

    tree = ET.parse(skeletonPath)
    network_tag = tree.getroot()

    listCodeNodes = []
    for node_tag in network_tag.findall("./nodes/node"):
        uda_tag = node_tag.find("./uda")
        listCodeNodes.append(uda_tag.attrib['value'])

    #Getting list of .png files
    pngList_by_Code = {}
    for nodeCode in listCodeNodes:
        pngList_by_Code[nodeCode] = []

    pattern = r"([A-Z]+-[0-9]+)"
    actualPath = os.path.join(subareaPath,typesig)
    for tipicidad in ["Tipico"]:
        tipicidadPath = os.path.join(actualPath, tipicidad)
        scenariosList = os.listdir(tipicidadPath)
        scenariosList = [file for file in scenariosList if not file.endswith(".ini")]
        for scenario in scenariosList:
            scenarioPath = os.path.join(tipicidadPath, scenario)
            scenarioContent = os.listdir(scenarioPath)
            if not scenario in ['HPM','HPT','HPN']: continue
            scenarioContent = [file for file in scenarioContent if file.endswith(".png")]
            for pngFile in scenarioContent:
                if re.search(pattern, pngFile):
                    pngList_by_Code[pngFile[:-4]].append(os.path.join(scenarioPath, pngFile))

    #Obtaining name of intersections:
    dataExcel = "./data/Datos Generales.xlsx"
    df_datos = pd.read_excel(dataExcel, sheet_name="DATOS", header=0, usecols="A:B")

    dictCode = {}
    for code, listPathsPNGs in pngList_by_Code.items():
        dictTurns = {}
        for pathPNG in listPathsPNGs:
            pathPNG_parts = pathPNG.split("\\")
            scenarioName = pathPNG_parts[-2]
            if scenarioName == 'HPM': turno = 'Mañana'
            elif scenarioName == 'HPT': turno = 'Tarde'
            elif scenarioName == 'HPN': turno = 'Noche'
            else: print(f"Error: No se encontró ningún escenario de HPM, HPT o HPN: {pathPNG}")
            texto = f"Tiempo de ciclo y fases semafóricas en el Turno {turno} de la intersección {code}"
            pathImage = pathPNG
            dictTurns[turno] = (texto, pathImage)
        dictCode[code] = dictTurns

    #Creating individual images with references
    imagesDirectory = os.path.join(subareaPath, "Imagenes")
    if not os.path.exists(imagesDirectory): os.mkdir(imagesDirectory)

    listWordPaths = []
    for code, dictTurns in dictCode.items():
        for turno in ["Mañana", "Tarde", "Noche"]:
            try:
                text, pathImg = dictTurns[turno]
            except KeyError as e:
                #No existen datos para un turno en específico de una intersección específica.
                continue

            try:
                parrafo_sig = df_datos[df_datos["CODE"] == code]["SIG"].values[0]
            except IndexError as e:
                parrafo_sig = "NO SE ENCONTRÓ NOMBRE RELACIÓN CON ESE CÓDIGO: data/Datos Generales.xlsx"

            doc_template = DocxTemplate("./templates/template_imagenes_parrafo.docx")
            newImage = InlineImage(doc_template, pathImg, width=Inches(6))
            doc_template.render({"parrafo_sig": parrafo_sig, "texto": text, "tabla": newImage}) #NOTE: {{sigactual}}
            turno_text = unidecode(turno)
            finalPath = os.path.join(imagesDirectory, f"{code}_{turno_text}_sig{typesig}.docx")
            doc_template.save(finalPath) #TODO: <---- Este estaba comentando, parace que este código no esta completo.
            listWordPaths.append(finalPath)
    
    sigactual_path = os.path.join(subareaPath, "Tablas", f"sig{typesig}.docx")
    
    filePathMaster = listWordPaths[0]
    filePathList = listWordPaths[1:]

    _combine_all_docx(filePathMaster, filePathList, sigactual_path)

    return sigactual_path