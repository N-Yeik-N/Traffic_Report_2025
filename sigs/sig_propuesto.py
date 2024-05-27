import os
import xml.etree.ElementTree as ET
import re
#Doc
from docxtpl import DocxTemplate, InlineImage
from docxcompose.composer import Composer
from docx import Document
from docx.shared import Inches
from unidecode import unidecode

scenarios_by_tipicidad = {
    'Típico': ['HVMAD', 'HPMAD', 'HVM', 'HPM', 'HVT', 'HPT', 'HVN', 'HPN'],
    'Atípico': ['HVMAD', 'HPM', 'HPT', 'HPN', 'HVN'],
}

def _combine_all_docx(filePathMaster, filePathsList, finalPath) -> None:
    number_of_sections = len(filePathsList)
    master = Document(filePathMaster)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document(filePathsList[i])
        composer.append(doc_temp)

    composer.save(finalPath)

def get_sigs_propuesto(subareaPath) -> None:

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
    actualPath = os.path.join(subareaPath,"Propuesto - base")
    for tipicidad in ["Tipico", "Atipico"]:
        tipicidadPath = os.path.join(actualPath, tipicidad)
        scenariosList = os.listdir(tipicidadPath)
        scenariosList = [file for file in scenariosList if not file.endswith(".ini")]
        for scenario in scenariosList:
            scenarioPath = os.path.join(tipicidadPath, scenario)
            scenarioContent = os.listdir(scenarioPath)
            scenarioContent = [file for file in scenarioContent if file.endswith(".png")]
            for pngFile in scenarioContent:
                if re.search(pattern, pngFile):
                    pngList_by_Code[pngFile[:-4]].append(os.path.join(scenarioPath, pngFile))

    listData = []
    for codeNode in listCodeNodes:
        for tipicidad in ["Típico", "Atípico"]:
            for code, listPathsPNGs in pngList_by_Code.items():
                if code == codeNode:
                    listCompare = scenarios_by_tipicidad[tipicidad]
                    for compareName in listCompare:
                        for pathPNG in listPathsPNGs:
                            pathPNG_parts = pathPNG.split("\\")
                            scenarioName = pathPNG_parts[-2]    
                            if scenarioName == compareName:
                                texto = f"Intersección {code} en la {scenarioName} del día {tipicidad}"
                                pathImage = pathPNG
                                listData.append((texto, pathImage, tipicidad, scenarioName, code))
                                break

    #Creating individual images with references
    imagesDirectory = os.path.join(subareaPath, "Imagenes")
    if not os.path.exists(imagesDirectory): os.mkdir(imagesDirectory)

    listWordPaths = []
    for text, pathImg, tipicidad, scenarioName, code in listData:
        doc_template = DocxTemplate("./templates/template_tablas4.docx")
        newImage = InlineImage(doc_template, pathImg, width=Inches(6))
        doc_template.render({"texto": text, "tabla": newImage})
        finalPath = os.path.join(imagesDirectory, f"SP_{code}_{unidecode(tipicidad).upper()}_{scenarioName}.docx")
        doc_template.save(finalPath)
        listWordPaths.append(finalPath)
    
    sigsProposed = os.path.join(subareaPath, "Tablas", "SigProposed.docx")
    filePathMaster = listWordPaths[0]
    filePathList = listWordPaths[1:]

    _combine_all_docx(filePathMaster, filePathList, sigsProposed)

    return sigsProposed

# if __name__ == '__main__':
#     subareaPath = r"C:\Users\dacan\OneDrive\Desktop\PRUEBAS\Maxima Entropia\04 Proyecto Universitaria (37 Int. - 19 SA)\6. Sub Area Vissim\Sub Area 016"
#     sigsProposed = get_sigs_propuesto(subareaPath)
#     print(sigsProposed)