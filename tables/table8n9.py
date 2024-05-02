import os
from tools.traffic_lights import get_info
#docx
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def table8n9(path_subarea):
    path_parts = path_subarea.split("/") #<--- LINUX
    subarea_id = path_parts[-1]
    proyect_folder = '/'.join(path_parts[:-2]) #<--- LINUX

    field_data = os.path.join(
        proyect_folder,
        "7. Informacion de Campo",
        subarea_id,
        "Tiempo de Ciclo Semaforico"
    )

    list_excels = os.listdir(field_data)
    list_excels = [os.path.join(field_data, file) for file in list_excels
                if file.endswith(".xlsx") and not file.startswith("~")]

    phasesList = []
    for excel in list_excels:
        phasesData = get_info(excel, "Tipico")
        phasesList.append(phasesData)

    #Número de filas:
    number_rows = []
    for phaseData in phasesList:
        number_rows.append(len(phaseData.phases))

    doc = Document()
    table = doc.add_table(rows=sum(number_rows)+1, cols=7)
    for i, texto in enumerate([
            "Código", "Intersección", "Tiempo de Ciclo", "Fases", "Verde", "Ámbar", "Rojo"
        ]):
            table.cell(0, i).text = texto
    row_no = 1
    for phaseData in phasesList: #TODO: Solo se escribe una vez datos generales y luego por filas. Averiguar qué hacer.
        table.cell(row_no, 0).text = phasesData.codigo
        table.cell(row_no, 1).text = phasesData.nombre
        table.cell(row_no, 2).text = phasesData.cycletime

        table.cell(row_no, 3).text = phasesData.phases[0] #Verde
        table.cell(row_no, 4).text = phasesData.phases[1] #Rojo
        table.cell(row_no, 5).text = phasesData.phases[2] #Ámbar

        row_no += 1

    table.style = "Table Grid"
    doc.save(f"./db/table9.docx")
        
if __name__ == '__main__':
    PATH = r"/home/chiky/Projects/REPORT/data/1. Proyecto Surco (Sub. 16 -59)/6. Sub Area Vissim/Sub Area 016"
    table8n9(PATH)