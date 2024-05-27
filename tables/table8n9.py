import os
from tables.tools.traffic_lights import get_info
import pandas as pd
from tables.tools.cycles import get_dates_cycles
from pathlib import Path

#docx
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def create_table9(path_subarea):
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
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, texto in enumerate([
            "Código", "Intersección", "Tiempo de Ciclo", "Fases", "Verde", "Ámbar", "Rojo"
        ]):
            table.cell(0, i).text = texto
    row_no = 1

    codeList = []
    for phase in phasesList: #TODO: Solo se escribe una vez datos generales y luego por filas. Averiguar qué hacer.
        table.cell(row_no, 0).text = phase.codigo
        codeList.append(phase.codigo)
        table.cell(row_no, 0).merge(table.cell(row_no+len(phaseData.phases)-1, 0))

        table.cell(row_no, 1).text = phase.nombre
        table.cell(row_no, 1).merge(table.cell(row_no+len(phaseData.phases)-1, 1))

        table.cell(row_no, 2).text = str(phase.cycletime)+' segundos'
        table.cell(row_no, 2).merge(table.cell(row_no+len(phaseData.phases)-1, 2))

        no_phase = 0
        for j in range(row_no, row_no+len(phase.phases)):
            table.cell(j, 3).text = f"Fase {no_phase+1}"
            table.cell(j, 4).text = str(phase.phases[no_phase][0]) #Verde
            table.cell(j, 5).text = str(phase.phases[no_phase][1]) #Rojo
            table.cell(j, 6).text = str(phase.phases[no_phase][2]) #Ámbar
            no_phase += 1

        row_no += len(phase.phases)

    #ESTÉTICA

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

    for id, x in zip([0,1,2,3,4,5,6],[0.5,2,1,0.7,0.5,0.5,0.5]):
        for cell in table.columns[id].cells:
            cell.width = Inches(x)

    table.style = "Table Grid"

    #The next code deletes the empty paragraph at the end of the table
    """ if doc.paragraphs[-1].text == "":
        p = doc.paragraphs[-1]._element
        p.getparent().remove(p)
        p._p = p._element = None #=O """

    table9_path = Path(path_subarea) / "Tablas" / "table9_sinREF.docx"
    doc.save(table9_path)

    doc_template = DocxTemplate(r"templates\template_tablas.docx")
    if len(codeList) == 1:
        texto = f"Tiempos de ciclo de la intersección {codeList[0]}"
    elif len(codeList) > 1:
        texto_aux = ""
        for i, code in enumerate(codeList):
            if i == len(codeList) - 1:
                texto_aux += f"y {code}"
            elif i == len(codeList) - 2:
                texto_aux += f"{code}"
            else:
                texto_aux += f"{code}, "
            
        texto = f"Tiempos de ciclos y fases de las intersecciones {texto_aux}"

    table9 = doc_template.new_subdoc(table9_path)

    doc_template.render({
        "texto": texto,
        "tabla": table9
    })

    finalPath = os.path.join(path_subarea, "Tablas", "table9.docx")
    doc_template.save(finalPath)

    return finalPath

def create_table8(path_subarea):
    path_excel = r"data\Cronograma Vissim.xlsx"
    df = pd.read_excel(path_excel, "Cronograma", header=0, usecols="A:E", skiprows=1)
    df = df.drop(columns=["Codigo TDR", "Entregable"])

    numsubarea = os.path.split(path_subarea)[1][-3:]
    no = int(numsubarea)

    code_by_subarea = df[df['Sub Area'] == no]['Codigo'].tolist()
    
    try:
        df_dates = get_dates_cycles(code_by_subarea)
    except Exception as e:
        print("Posiblemente no existe un elemento en la base de datos: Datos de Ciclos.xlsx")
        print("Error: ", e)

    doc = Document()
    table = doc.add_table(rows = 7, cols = 5)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    #Headers
    for i, header in enumerate(["Intersección", "Día", "Tipicidad", "Turno", "Horario"]):
        table.cell(0,i).text = header

    typicalDay = df_dates['Dia Tipico'].unique().tolist()[0]
    atypicalDay = df_dates['Dia Atipico'].unique().tolist()[0]
    
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
    path_table8 = Path(path_subarea) / "Tablas" / "table8.docx"
    doc.save(path_table8)

    return path_table8