import os

#docx
from docxcompose.composer import Composer
from docxtpl import DocxTemplate, InlineImage
from docx import Document
from docx.shared import Inches

def _combine_all_docx(filePathMaster, filePathsList, finalPath) -> None:
    number_of_sections = len(filePathsList)
    master = Document(filePathMaster)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document(filePathsList[i])
        composer.append(doc_temp)

    composer.save(finalPath)

def create_resultados_images(subareaPath) -> str:
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
        gehImage = InlineImage(doc_template, gehFilePath, width = Inches(2.37))
        r2Image = InlineImage(doc_template, r2FilePath, width = Inches(3))
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