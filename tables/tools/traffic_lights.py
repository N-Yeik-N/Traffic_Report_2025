from openpyxl import load_workbook
import re
from dataclasses import dataclass

@dataclass
class cycleTime:
    tipicidad: str
    codigo: str
    nombre: str
    cycletime: int
    phases: list

""" 
Solo se extraerá información para los tiempos de ciclo en el turno mañana.
"""

def get_dates(path):
    pass

def get_info(path, tipicidad):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    codigo = ws['D3'].value
    pattern1 = r'([A-Z]+-[0-9]+)'
    pattern2 = r'([A-Z]+[0-9]+)'
    coincidence1 = re.search(pattern1, codigo)
    coincidence2 = re.search(pattern2, codigo)
    if coincidence1:
        codinterseccion = coincidence1.group(1)
    elif coincidence2:
        codinterseccion = coincidence2.group(1)
    else:
        print("ERROR - Excel sin código de tipo AA-99 o AA99:\n",path)
    
    nominterseccion = ws['D4'].value
    parts = nominterseccion.split(":")
    nominterseccion = parts[1].strip()

    if tipicidad == "Tipico":
        tc = int(ws['E10'].value)
        numfase = [[elem.value for elem in row] for row in ws[slice("H10","AK10")]][0].index(None)//2
        phases = [[elem.value for elem in row] for row in ws[slice("H10","AK10")]][0][:numfase*2]
        phases = [phases[i:i+3] for i in range(0, len(phases),3)]
    elif tipicidad == "Atipico":
        tc = int(ws['E18'].value)
        numfase = [[elem.value for elem in row] for row in ws[slice("H18","AK18")]][0].index(None)//2
        phases = [[elem.value for elem in row] for row in ws[slice("H18","AK18")]][0][:numfase*2]
        phases = [phases[i:i+3] for i in range(0, len(phases),3)]

    wb.close()

    data = cycleTime(
        tipicidad = tipicidad,
        codigo = codinterseccion,
        nombre = nominterseccion,
        cycletime = tc,
        phases = phases,
    )

    return data


if __name__ == '__main__':
    PATH = r"data/1. Proyecto Surco (Sub. 16 -59)/7. Informacion de Campo/Sub Area 016/Tiempo de Ciclo Semaforico/SS-77_Av. Rosa Lozano - Jr. Geranios_Tiempo de Ciclo y Fases.xlsx"
    data = get_info(PATH, "Tipico")
    print(data)