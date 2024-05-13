import os
import pandas as pd
from pathlib import Path
from tables.tools.pedestrian import *

#docx
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def create_table12(path_subarea):
    path_excel = r"data\Cronograma Vissim.xlsx"
    df = pd.read_excel(path_excel, "Cronograma", header=0, usecols="A:E", skiprows=1)
    df = df.drop(columns=["Codigo TDR", "Entregable"])

    numsubarea = os.path.split(path_subarea)[1][-3:]
    no = int(numsubarea)

    code_by_subarea = df[df['Sub Area'] == no]['Codigo'].tolist()

    path_parts = path_subarea.split("/")
    subarea_id = path_parts[-1]
    proyect_folder = '/'.join(path_parts[:-2])

    field_data = Path(proyect_folder) / "7. Informacion de Campo" / subarea_id / "Peatonal"
    excel_tipicidades = {}

    for tipicidad in ["Tipico", "Atipico"]:
        tip_data = field_data / tipicidad
        list_excels = os.listdir(tip_data)
        list_excels = [str(tip_data / file) for file in list_excels if file.endswith(".xlsm") and not file.startswith("~")]
        excel_tipicidades[tipicidad] = list_excels

    listDatesTipico = []
    listDatesAtipico = []
    for tipicidad, list_excels in excel_tipicidades.items():
        for excel in list_excels:
            if tipicidad == "Tipico":
                listDatesTipico.append(get_dates(excel))
            else:
                listDatesAtipico.append(get_dates(excel))

    listDatesTipico = list(set(listDatesTipico))

    if len(listDatesTipico) > 1:
        print("Para los contes peatonales TÍPICOS se tienen fechas distintas de toma de datos.")
        print("Se utilizará solo una fecha")
    typicalDay = listDatesTipico[0]

    listDatesAtipico = list(set(listDatesAtipico))

    if len(listDatesAtipico) > 1:
        print("Para los contes peatonales ATÍPICOS se tienen fechas distintas de toma de datos.")
        print("Se utilizará solo una fecha")
    atypicalDay = listDatesAtipico[0]

    doc = Document()
    table = doc.add_table(rows = 7, cols = 5)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    #Headers:
    for i, header in enumerate(["Intersección", "Día", "Tipicidad", "Turno", "Horario"]):
        table.cell(0,i).text = header

    for i, day in enumerate([typicalDay, atypicalDay]):
        if i == 0:
            for j in range(1,4,1):
                table.cell(j,1).text = day
        else:
            for j in range(4,7,1):
                table.cell(j,1).text = day

    for i, turno in enumerate(["Mañana", "Tarde", "Noche"]*2):
        table.cell(i+1,3).text = turno

    for i, tipicidad in enumerate(["Tipico", "Atipico"]*3):	
        table.cell(i+1,2).text = tipicidad

    for i, horario in enumerate([
        "06:30 - 09:30",
        "12:00 - 15:00",
        "17:30 - 20:30",
        "06:30 - 09:30",
        "12:00 - 15:00",
        "17:30 - 20:30",
    ]):
        table.cell(i+1, 4).text = horario

    texto = ""
    for i, code in enumerate(code_by_subarea):
        if i == len(code_by_subarea)-1:
            texto += code
        else:
            texto += code + '\n'

    table.cell(1,0).text = texto
    table.cell(1,0).merge(table.cell(6,0))

    for selected_row in [0]:
        for cell in table.rows[selected_row].cells:
            for paragraph in cell.paragraphs:
                run = paragraph.runs[0]
                run.font.bold = True

    for i in range(len(table.columns)):
        cell_xml_element = table.rows[0].cells[i]._tc
        table_cell_properties = cell_xml_element.get_or_add_tcPr()
        shade_obj = OxmlElement('w:shd')
        shade_obj.set(qn('w:fill'),'B4C6E7')
        table_cell_properties.append(shade_obj)

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                run = paragraph.runs[0]
                run.font.name = 'Arial Narrow'
                run.font.size = Pt(11)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for id, x in zip([0,1,2,3,4],
                     [0.9,0.8,0.8,0.5,1]):
        for cell in table.columns[id].cells:
            cell.width = Inches(x)

    table.style = 'Table Grid'
    table12_path = Path(path_subarea) / "Tablas" / "table12.docx"
    doc.save(table12_path)

    return table12_path