from pdf2image import convert_from_path
import os
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from docx import Document
from pathlib import Path

def _append_document_content(source_doc, target_doc) -> None:
    for element in source_doc.element.body:
        target_doc.element.body.append(element)

def get_codes(subarea_path) -> list:
    num_subarea = os.path.split(subarea_path)[1][-3:]
    df_general = pd.read_excel("./data/Datos Generales.xlsx", sheet_name="DATOS", header=0, usecols="A:E")
    listCodes = df_general[df_general['Sub_Area'] == int(num_subarea)]['Code'].unique().tolist()
    return listCodes

def convert_pdf_to_image(pdf_path, output_path, name) -> None:
    #print(pdf_path)
    try:
        images = convert_from_path(pdf_path)
    except Exception as e:
        print(e)
        return print("ERROR: No se pudo convertir el PDF a imagen. Si ya fue creado antes debes borrarlo primero.")

    if len(images) > 1:
        return print("ERROR: PDF tiene mas de una hoja.")
    
    destiny_path = os.path.join(output_path, name + '.png')
    for i in range(len(images)):
        final_path = os.path.join(destiny_path)
        images[i].save(final_path, 'PNG')

    #print(f"Imagenes guardadas en {destiny_path}")

    return destiny_path

def create_histogramas_subdocs(resultList: list, path_subarea: str | Path) -> str: 
    PATH_TEMPLATE = r".\templates\template_tablas.docx"

    listWords = []
    for code, tipicidad, imagePath in resultList:
        doc = DocxTemplate(PATH_TEMPLATE)
        if tipicidad == "T": tipico_texto = "típico"
        elif tipicidad == "A": tipico_texto = "atípico"

        texto = f"Histograma vehicular de las HP por cada turno en la intersección {code} día {tipico_texto}"
        image = InlineImage(doc, imagePath, width=Inches(6))

        variables = {
            'texto': texto,
            'tabla': image,
        }

        doc.render(variables)
        final_path = Path(path_subarea) / "Tablas" / f"histograma_{code}_{tipicidad}.docx"
        doc.save(final_path)
        listWords.append(final_path) #Already in order

    doc_target = Document(listWords[0])
    for i in range(len(listWords)):
        if i == 0: continue
        doc_source = Document(listWords[i])
        _append_document_content(doc_source, doc_target)
        histograma_path = Path(path_subarea) / "Tablas" / "histogramas.docx"
        doc_target.save(histograma_path)
        doc_target = Document(histograma_path)
        
    doc_target.save(histograma_path)

    return histograma_path

def create_flujogramas_vehicular_subdocs(resultList: list, path_subarea: str | Path) -> str:
    PATH_TEMPLATE = r".\templates\template_tablas.docx"

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

    doc_target = Document(listWords[0])
    for i in range(len(listWords)):
        if i == 0: continue
        doc_source = Document(listWords[i])
        _append_document_content(doc_source, doc_target)
        flujogramas_path = Path(path_subarea) / "Tablas" / "flujogramas_vehiculares.docx"
        doc_target.save(flujogramas_path)
        doc_target = Document(flujogramas_path)

    doc_target.save(flujogramas_path)

    return flujogramas_path

def create_flujograma_peatonal_subdocs(resultList: list, path_subarea: str | Path) -> str:
    PATH_TEMPLATE = r".\templates\template_tablas.docx"

    listWords = []
    for code, imagePath in resultList:
        doc = DocxTemplate(PATH_TEMPLATE)
        texto = f"Flujograma peatonal de laintersección {code} HPM día típico"
        image = InlineImage(doc, imagePath, width=Inches(6))

        variables = {
            'texto': texto,
            'tabla': image,
        }

        doc.render(variables)
        final_path = Path(path_subarea) / "Tablas" / f"flujograma_peatonal_{code}.docx"
        doc.save(final_path)
        listWords.append(final_path)
        listWords.append(final_path)

    doc_target = Document(listWords[0])
    for i in range(len(listWords)):
        if i == 0: continue
        doc_source = Document(listWords[i])
        _append_document_content(doc_source, doc_target)
        flujograma_path = Path(path_subarea) / "Tablas" / "flujogramas_peatonales.docx"
        doc_target.save(flujograma_path)
        doc_target = Document(flujograma_path)
    
    doc_target.save(flujograma_path)

    return flujograma_path