import os
from openpyxl import load_workbook


def create_table15(subareaPath):
    balanceadoPath = os.path.join(subareaPath, "Balanceado")
    for tipicidad in ["Tipico", "Atipico"]:
        tipicidadPath = os.path.join(balanceadoPath, tipicidad)
        scenariosList = os.listdir(tipicidadPath)
        scenariosList = [file for file in scenariosList if not file.endswith(".ini")]
        

# if __name__ == '__main__':
#     path = r"C:\Users\dacan\OneDrive\Desktop\PRUEBAS\Maxima Entropia\04 Proyecto Universitaria (37 Int. - 19 SA)\6. Sub Area Vissim\Sub Area 016"
#     create_table15(path)