from pdf2image import convert_from_path
import os
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from docx import Document
from pathlib import Path
from docxcompose.composer import Composer

def _combine_all_docx(filePathMaster, filePathsList, finalPath) -> None:
    number_of_sections = len(filePathsList)
    master = Document(filePathMaster)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document(filePathsList[i])
        composer.append(doc_temp)

    composer.save(finalPath)

def get_codes(subarea_path) -> list:
    num_subarea = os.path.split(subarea_path)[1][-3:]
    df_general = pd.read_excel("./data/Datos Generales.xlsx", sheet_name="DATOS", header=0, usecols="A:E")
    listCodes = df_general[df_general['Sub_Area'] == int(num_subarea)]['Code'].unique().tolist()
    return listCodes

def convert_pdf_to_image(pdf_path, output_path, name) -> None:
    try:
        images = convert_from_path(pdf_path)
    except AttributeError as e:
        return print("ERROR: No se pudo convertir el PDF a imagen. Si ya fue creado antes debes borrarlo primero.")
    except Exception as inst:
        print(f"ERROR: No se puede el PDF a imagen:\n{pdf_path}")
        raise inst

    assert len(images) <= 1, "ERROR: PDF tiene mas de una hoja."
    
    destiny_path = os.path.join(output_path, name + '.png')
    for i in range(len(images)):
        final_path = os.path.join(destiny_path)
        images[i].save(final_path, 'PNG')

    #print(f"Imagenes guardadas en {destiny_path}")

    return destiny_path

def create_histogramas_subdocs(resultList: list, path_subarea: str | Path, agentType: str) -> str: 
    PATH_TEMPLATE = r".\templates\template_imagenes.docx"

    listWords = []
    for code, tipicidad, imagePath in resultList:
        doc = DocxTemplate(PATH_TEMPLATE)
        if tipicidad == "T": tipico_texto = "típico"
        elif tipicidad == "A": tipico_texto = "atípico"
        if agentType == "Vehicular":
            texto = f"Histograma vehicular de las HP por cada turno en la intersección {code} día {tipico_texto}"
        elif agentType == "Peatonal":
            texto = f"Histograma peatonal de las HP por cada turno en la intersección {code} día {tipico_texto}"
        image = InlineImage(doc, imagePath, width=Inches(6))

        variables = {
            'texto': texto,
            'tabla': image,
        }

        doc.render(variables)
        final_path = Path(path_subarea) / "Tablas" / f"histograma_{code}_{tipicidad}.docx"
        doc.save(final_path)
        listWords.append(final_path) #Already in order

    histograma_path = os.path.join(path_subarea, "Tablas", f"histogramas_{agentType}.docx")
    filePathMaster = listWords[0]
    filePathList = listWords[1:]
    _combine_all_docx(filePathMaster, filePathList, histograma_path)

    return histograma_path

def create_flujogramas_vehicular_subdocs(resultList: list, path_subarea: str | Path) -> str:
    PATH_TEMPLATE = r".\templates\template_imagenes.docx"

    listWords = []
    for code, imagePath in resultList:
        doc = DocxTemplate(PATH_TEMPLATE)
        texto = f"Flujograma vehicular de la intersección {code} HPM día típico"
        image = InlineImage(doc, imagePath, width=Inches(6))

        variables = {
            'texto': texto,
            'tabla': image,
        }

        doc.render(variables)
        final_path = Path(path_subarea) / "Tablas" / f"flujograma_vehicular_{code}.docx"
        doc.save(final_path)
        listWords.append(final_path)

    flujogramas_path = os.path.join(path_subarea, "Tablas", "flujogramas_vehiculares.docx")

    filePathMaster = listWords[0]
    filePathList = listWords[1:]
    _combine_all_docx(filePathMaster, filePathList, flujogramas_path)

    return flujogramas_path

def create_flujograma_peatonal_subdocs(resultList: list, path_subarea: str | Path) -> str:
    PATH_TEMPLATE = r".\templates\template_imagenes.docx"

    listWords = []
    for code, imagePath in resultList:
        doc = DocxTemplate(PATH_TEMPLATE)
        texto = f"Flujograma peatonal de la intersección {code} HPM día típico"
        image = InlineImage(doc, imagePath, width=Inches(6))

        variables = {
            'texto': texto,
            'tabla': image,
        }

        doc.render(variables)
        final_path = Path(path_subarea) / "Tablas" / f"flujograma_peatonal_{code}.docx"
        doc.save(final_path)
        listWords.append(final_path)

    flujograma_path = os.path.join(path_subarea, "Tablas", "flujogramas_peatonales.docx")
    filePathMaster = listWords[0]
    filePathList = listWords[1:]
    _combine_all_docx(filePathMaster, filePathList, flujograma_path)

    return flujograma_path