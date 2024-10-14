import os

#docx
from docxcompose.composer import Composer
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Inches, Cm
from images.tools.diana import create_dianas
from images.tools.r2 import create_r2s

def _combine_all_docx(filePathMaster, filePathsList, finalPath) -> None:
    number_of_sections = len(filePathsList)
    master = Document(filePathMaster)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document(filePathsList[i])
        composer.append(doc_temp)

    composer.save(finalPath)

def create_resultados_images(subareaPath) -> str:
    #Locating GEH-R2 file:
    balancedFolder = os.path.join(subareaPath, "Balanceado")
    for tipicidad in ["Tipico", "Atipico"]:
        tipicidadFolder = os.path.join(balancedFolder, tipicidad)
        scenariosContent = os.listdir(tipicidadFolder)
        scenariosContent = [file for file in scenariosContent if not file.endswith(".ini")]
        for scenario in scenariosContent:
            scenarioFolder = os.path.join(tipicidadFolder, scenario)
            #Looking for GEH-R2 excel
            folderContent = os.listdir(scenarioFolder)
            for file in folderContent:
                if "GEH-R2.xlsm" in file:
                    excelPath = os.path.join(scenarioFolder, file)
                    break

    #Creating images:
    try:
        create_r2s(excelPath)
    except Exception as e:
        print("Error creando gráfica de R2s")
        print(str(e))
        raise e
    try:
        create_dianas(excelPath)
    except TypeError as e:
        print("Tabla 16\tError\tFaltan más intersecciones o tipos vehiculares en el excel de GEH-R2")
    except Exception as e:
        raise e

    #Creating tables
    tablasPath = os.path.join(subareaPath, "Tablas")
    listContent = os.listdir(tablasPath)

    gehFiles = [file for file in listContent if file.endswith(".png") and 'GEH_' in file]
    r2Files = [file for file in listContent if file.endswith(".png") and 'R2_' in file]

    imagesDict = {} #'VEHTYPE': ('GEH_VEHTYPE.png', 'R2_VEHTYPE.png')

    for gehFile in gehFiles:
        for r2File in r2Files:
            if gehFile.split('_')[1][:-4] == r2File.split('_')[1][:-4]:
                imagesDict[gehFile.split('_')[1][:-4]] = (os.path.join(tablasPath, gehFile), os.path.join(tablasPath, r2File))
                break

    listWordPaths = []
    for vehicleType, (gehFilePath, r2FilePath) in imagesDict.items():
        doc_template = DocxTemplate(r"templates\template_tablas3.docx")
        gehImage = InlineImage(doc_template, gehFilePath, height = Inches(5))
        r2Image = InlineImage(doc_template, r2FilePath, height = Inches(5))
        text = f'Análisis de GEH y R2 de {vehicleType}'
        doc_template.render({
            'texto': text,
            'gehImage': gehImage,
            'r2Image': r2Image,
        })
        finalPath = os.path.join(tablasPath, f"resultados_{vehicleType}.docx")
        doc_template.save(finalPath)
        listWordPaths.append(finalPath)

    resultsPath = os.path.join(tablasPath, "Resultados.docx")

    filePathMaster = listWordPaths[0]
    filePathList = listWordPaths[1:]

    _combine_all_docx(filePathMaster, filePathList, resultsPath)
    
    return resultsPath