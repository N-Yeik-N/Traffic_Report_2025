import os
import pandas as pd
from pathlib import Path
from tables.tools.peakfinder import peakhour_finder, compute_ph_system

#docx
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def create_table2n3(path_subarea):
    path_parts = path_subarea.split("/") #<--- Linux...
    subarea_id = path_parts[-1]
    proyect_folder = '/'.join(path_parts[:-2]) #<--- Linux...
    field_data = Path(proyect_folder) / "7. Informacion de Campo" / subarea_id / "Vehicular"

    excel_tipicidades = {}

    for tipicidad in ["Tipico","Atipico"]:
        tip_data = field_data / tipicidad
        list_excels = os.listdir(tip_data)
        list_excels = [str(tip_data / file) for file in list_excels if file.endswith(".xlsm") and not file.startswith("~")]
        excel_tipicidades[tipicidad] = list_excels

    #################
    # Intersections #
    #################

    tipico_info = {}
    count_tip = 1
    atipico_info = {}
    count_ati = 1

    system_tip = {}
    system_ati = {}
    count_sys_t = 1
    count_sys_a = 1

    day_tip_list = []
    day_ati_list = []

    path_excel = r"data\Cronograma Vissim.xlsx"
    df = pd.read_excel(path_excel, sheet_name='Cronograma', header=0, usecols="A:E", skiprows=1)
    df = df.drop(columns=["Codigo TDR", "Entregable"])

    numsubarea = os.path.split(path_subarea)[1][-3:]
    no = int(numsubarea)

    code_by_subarea = df[df['Sub Area'] == no]['Codigo'].tolist()

    for key, data in excel_tipicidades.items():
        for excel in data:
            excel_dict = peakhour_finder(excel)
            #System level
            if key == "Tipico":
                system_tip[count_sys_t] = excel_dict
                count_sys_t += 1
                day_tip_list.append(excel_dict.fecha)
            elif key == "Atipico":
                system_ati[count_sys_a] = excel_dict
                count_sys_a += 1
                day_ati_list.append(excel_dict.fecha)
            #Intersection level
            hour1 = excel_dict.id_morning//4
            hour2 = excel_dict.id_evening//4
            hour3 = excel_dict.id_night//4
            minutes1 = excel_dict.id_morning%4*15
            minutes2 = excel_dict.id_evening%4*15
            minutes3 = excel_dict.id_night%4*15
            ph1 = "{:02d}:{:02d} - {:02d}:{:02d}".format(hour1, minutes1, hour1+1, minutes1)
            ph2 = "{:02d}:{:02d} - {:02d}:{:02d}".format(hour2, minutes2, hour2+1, minutes2)
            ph3 = "{:02d}:{:02d} - {:02d}:{:02d}".format(hour3, minutes3, hour3+1, minutes3)
            if key == "Tipico":
                node = {
                    'codinterseccion': excel_dict.codigo,
                    'nominterseccion': excel_dict.name,
                    'hpinterseccionmt': ph1,
                    'hpintersecciontt': ph2,
                    'hpinterseccionnt': ph3,
                }
                tipico_info[count_tip] = node
                count_tip += 1

            elif key == "Atipico":
                node = {
                    'codinterseccion': excel_dict.codigo,
                    'nominterseccion': excel_dict.name,
                    'hpinterseccionma': ph1,
                    'hpinterseccionta': ph2,
                    'hpinterseccionna': ph3,
                }

                atipico_info[count_ati] = node
                count_ati += 1

    #TIPICO

    MORNING = []
    EVENING = []
    NIGHT = []
    for key, datos in system_tip.items():
        MORNING.append((datos.id_morning, datos.vol_morning))
        EVENING.append((datos.id_evening, datos.vol_evening))
        NIGHT.append((datos.id_night, datos.vol_night))

    hoursystem1 = compute_ph_system(MORNING)
    hoursystem2 = compute_ph_system(EVENING)
    hoursystem3 = compute_ph_system(NIGHT)

    phsystem1 = "{:02d}:{:02d} - {:02d}:{:02d}".format(hoursystem1//4, hoursystem1%4*15, hoursystem1//4+1, hoursystem1%4*15)
    phsystem2 = "{:02d}:{:02d} - {:02d}:{:02d}".format(hoursystem2//4, hoursystem2%4*15, hoursystem2//4+1, hoursystem2%4*15)
    phsystem3 = "{:02d}:{:02d} - {:02d}:{:02d}".format(hoursystem3//4, hoursystem3%4*15, hoursystem3//4+1, hoursystem3%4*15)

    dict_system_t = {
        'hpsistemamt': phsystem1,
        'hpsistematt': phsystem2,
        'hpsistemant': phsystem3,
    }

    #ATIPICO

    MORNING = []
    EVENING = []
    NIGHT = []
    for key, datos in system_ati.items():
        MORNING.append((datos.id_morning, datos.vol_morning))
        EVENING.append((datos.id_evening, datos.vol_evening))
        NIGHT.append((datos.id_night, datos.vol_night))

    hoursystem1 = compute_ph_system(MORNING)
    hoursystem2 = compute_ph_system(EVENING)
    hoursystem3 = compute_ph_system(NIGHT)

    phsystem1 = "{:02d}:{:02d} - {:02d}:{:02d}".format(hoursystem1//4, hoursystem1%4*15, hoursystem1//4+1, hoursystem1%4*15)
    phsystem2 = "{:02d}:{:02d} - {:02d}:{:02d}".format(hoursystem2//4, hoursystem2%4*15, hoursystem2//4+1, hoursystem2%4*15)
    phsystem3 = "{:02d}:{:02d} - {:02d}:{:02d}".format(hoursystem3//4, hoursystem3%4*15, hoursystem3//4+1, hoursystem3%4*15)

    dict_system_a = {
        'hpsistemama': phsystem1,
        'hpsistemata': phsystem2,
        'hpsistemana': phsystem3,
    }

    day_tip = list(set(day_tip_list))[0]
    day_ati = list(set(day_ati_list))[0]
    dcontet = day_tip.strftime("%d de %B del %Y") #<---
    dcontea = day_ati.strftime("%d de %B del %Y") #<---

    dconteot = day_tip.strftime("%d/%m/%Y") #<---
    dconteoa = day_ati.strftime("%d/%m/%Y") #<---

    ###################
    # CREATING TABLE 2#
    ###################

    doc = Document()
    table = doc.add_table(rows = 5+len(code_by_subarea)*2, cols = 5)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    #Default texts
    table.cell(0,0).text = "Código"
    table.cell(0,1).text = 'Intersección'
    table.cell(0,2).text = 'Hora Punta Turno Mañana'
    table.cell(0,3).text = 'Hora Punta Turno Tarde'
    table.cell(0,4).text = 'Hora Punta Turno Noche'
    table.cell(1,0).text = "Día típico"
    table.cell(1,0).merge(table.cell(1,4))

    hp_row1 = 1+len(code_by_subarea)+1
    table.cell(hp_row1,0).text = "Hora Punta"
    table.cell(hp_row1,0).merge(table.cell(hp_row1,1))
    table.cell(hp_row1,2).text = dict_system_t['hpsistemamt']
    table.cell(hp_row1,3).text = dict_system_t['hpsistematt']
    table.cell(hp_row1,4).text = dict_system_t['hpsistemant']

    table.cell(1+len(code_by_subarea)+2,0).text = "Día Atípico"
    table.cell(1+len(code_by_subarea)+2,0).merge(table.cell(1+len(code_by_subarea)+2,4))

    hp_row2 = 1+len(code_by_subarea)*2+3
    table.cell(hp_row2,0).text = "Hora Punta"
    table.cell(hp_row2,0).merge(table.cell(hp_row2,1))
    table.cell(hp_row2,2).text = dict_system_a['hpsistemama']
    table.cell(hp_row2,3).text = dict_system_a['hpsistemata']
    table.cell(hp_row2,4).text = dict_system_a['hpsistemana']

    start_tipico = 2
    start_atipico = 2+len(code_by_subarea)+2

    for key, node in tipico_info.items():
        table.cell(start_tipico+key-1,0).text = node['codinterseccion']
        table.cell(start_tipico+key-1,1).text = node['nominterseccion']
        table.cell(start_tipico+key-1,2).text = node['hpinterseccionmt']
        table.cell(start_tipico+key-1,3).text = node['hpintersecciontt']
        table.cell(start_tipico+key-1,4).text = node['hpinterseccionnt']
        
    for key, node in atipico_info.items():
        table.cell(start_atipico+key-1,0).text = node['codinterseccion']
        table.cell(start_atipico+key-1,1).text = node['nominterseccion']
        table.cell(start_atipico+key-1,2).text = node['hpinterseccionma']
        table.cell(start_atipico+key-1,3).text = node['hpinterseccionta']
        table.cell(start_atipico+key-1,4).text = node['hpinterseccionna']

    for selected_row in [0, 1, hp_row1, 1+len(code_by_subarea)+2, hp_row2]:
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

    for id, x in zip([0,1],[0.5,3]):
        for cell in table.columns[id].cells:
            cell.width = Inches(x)

    table.style = 'Table Grid'

    final_path_2 = Path(path_subarea) / "Tablas" / "table2.docx"

    doc.save(final_path_2)

    ####################
    # CREATING TABLE 3 #
    ####################

    doc = Document()
    table = doc.add_table(rows=7 ,cols=5)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table.cell(0,0).text = "Intersección"
    table.cell(0,1).text = "Día"
    table.cell(0,2).text = "Tipicidad"
    table.cell(0,3).text = "Turno"
    table.cell(0,4).text = "Horario"

    #codinterseccion:
    set_codes = ""
    for i, (key, value) in enumerate(tipico_info.items()):
        if i == len(tipico_info)-1:
            set_codes += value['codinterseccion']
        else:
            set_codes += value['codinterseccion']+'\n'
    
    table.cell(1,0).text = set_codes
    table.cell(1,0).merge(table.cell(6,0))

    for i in range(3):
        table.cell(1+i,2).text = "Típico"
        table.cell(4+i,2).text = "Atípico"

    for i, valor in enumerate(["Mañana", "Tarde", "Noche"]):
        table.cell(1+i,3).text = valor
        table.cell(4+i,3).text = valor

    for i, valor in enumerate(["06:30 - 09:30",
                "12:00 - 15:00",
                "17:30 - 20:30",
                "06:30 - 09:30",
                "12:00 - 15:00",
                "17:30 - 20:30",]):
        table.cell(1+i,4).text = valor

    for i in range(3):
        table.cell(1+i,1).text = dconteot
        table.cell(4+i,1).text = dconteoa

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

    for id, x in zip([0,1,2,3,4],[0.5,1,1,0.8,1.2]):
        for cell in table.columns[id].cells:
            cell.width = Inches(x)

    table.style = 'Table Grid'

    final_path_3 = Path(path_subarea) / "Tablas" / "table3.docx"

    doc.save(final_path_3)

    return final_path_2, final_path_3, dconteot, dconteoa