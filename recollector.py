import os
import re
from pathlib import Path
from peakhour.peakfinder import peakhour_finder, compute_ph_system
from tale.tale_tools import tale_by_excel
from boarding.boarding_tools import board_by_excel
import locale
from clases_data import *
from dataclasses import asdict

try:
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
except locale.Error as e:
    pass
    #print(f"Error al establecer la configuración regional: {e}")

#############
# UBICACION #
#############

def location(path_subarea):
    # Folders #
    path_parts = path_subarea.split("/") #<--- Linux...
    numsubarea = os.path.split(path_subarea)[1][-3:] #<---
    nameproject = path_parts[-3] #<---
    pattern = r'Proyecto\s+([^()]+)\s+\((.*?)\)'
    coincidence = re.search(pattern, nameproject)
    if coincidence:
        nomdistrito = coincidence.group(1) #<---

    dict_INFO = {
        'numsubarea': numsubarea,
        'nameproject': nameproject,
        'nomdistrito': nomdistrito,
    }

    return dict_INFO

###################################
# CONTEO Y HORA DE MÁXIMA DEMANDA #
###################################

def counts(path_subarea):
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
                INFO = typicalInterseccion(
                    codinterseccion = excel_dict.codigo,
                    nominterseccion = excel_dict.name,
                    hpinterseccionmt=ph1,
                    hpintersecciontt=ph2,
                    hpinterseccionnt=ph3,
                )
                tipico_info[count_tip] = INFO
                count_tip += 1

            elif key == "Atipico":
                INFO = atypicalInterseccion(
                    codinterseccion = excel_dict.codigo,
                    nominterseccion = excel_dict.name,
                    hpinterseccionma=ph1,
                    hpinterseccionta=ph2,
                    hpinterseccionna=ph3,
                )
                atipico_info[count_ati] = INFO
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

    peakhours_tipico = typicalSystem( #<---
        hpsistemamt=phsystem1,
        hpsistematt=phsystem2,
        hpsistemant=phsystem3,
    )

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

    peakhours_atipico = atypicalSystem( #<---
        hpsistemama=phsystem1,
        hpsistemata=phsystem2,
        hpsistemana=phsystem3,
    )

    day_tip = list(set(day_tip_list))[0]
    day_ati = list(set(day_ati_list))[0]
    dcontet = day_tip.strftime("%d de %B del %Y") #<---
    dcontea = day_ati.strftime("%d de %B del %Y") #<---

    dconteot = day_tip.strftime("%d/%m/%Y") #<---
    dconteoa = day_ati.strftime("%d/%m/%Y") #<---

    HORAS_RESULTADOS = {
        "INFO_TIPICO": tipico_info,
        "INFO_ATIPICO": atipico_info,
        "SISTEMA_TIPICO": peakhours_tipico,
        "SISTEMA_ATIPICO": peakhours_atipico,
        "DIA_CONTEO_TIPICO": dcontet,
        "DIA_CONTEO_ATIPICO": dcontea,
        "FECHA_CONTEO_TIPICO": dconteot,
        "FECHA_CONTEO_ATIPICO": dconteoa
    }

    return HORAS_RESULTADOS

#####################
# LONGITUD DE COLAS #
#####################

def tales(path_subarea):
    path_parts = path_subarea.split("\\")
    subarea_id = path_parts[-1]
    proyect_folder = '\\'.join(path_parts[:-2])

    field_data = Path(proyect_folder) / "7. Informacion de Campo" / subarea_id / "Longitud de Cola"
    excel_tipicidades = {}

    for tipicidad in ["Tipico","Atipico"]:
        tip_data = field_data / tipicidad
        list_excels = os.listdir(tip_data)
        list_excels = [str(tip_data / file) for file in list_excels if file.endswith(".xlsx") and not file.startswith("~")]
        excel_tipicidades[tipicidad] = list_excels

    day_tip_list = []
    day_ati_list = []

    for key, data in excel_tipicidades.items():
        for excel in data:
            dict_info = tale_by_excel(excel)
            if key == "Tipico":
                day_tip_list.append(dict_info)
            elif key == "Atipico":
                day_ati_list.append(dict_info)

    return day_tip_list, day_ati_list

def boarding(path_subarea):
    path_parts = path_subarea.split("\\")
    subarea_id = path_parts[-1]
    proyect_folder = '\\'.join(path_parts[:-2])

    field_data = Path(proyect_folder) / "7. Informacion de Campo" / subarea_id / "Embarque y Desembarque"
    excel_tipicidades = {}

    for tipicidad in ["Tipico","Atipico"]:
        tip_data = field_data / tipicidad
        list_excels = os.listdir(tip_data)
        list_excels = [str(tip_data / file) for file in list_excels if file.endswith(".xlsx") and not file.startswith("~")]
        excel_tipicidades[tipicidad] = list_excels

    day_tip_list = []
    day_ati_list = []
    for key, data in excel_tipicidades.items():
        for excel in data:
            dict_info = board_by_excel(excel)
            if key == "Tipico":
                day_tip_list.append(dict_info)
            elif key == "Atipico":
                day_ati_list.append(dict_info)

    print(day_tip_list)

# if __name__ == '__main__':
#     PATH = r"C:\Users\dacan\OneDrive\Desktop\PRUEBAS\Maxima Entropia\04 Proyecto Universitaria (37 Int. - 19 SA)\6. Sub Area Vissim\Sub Area 016"
#     boarding(PATH)