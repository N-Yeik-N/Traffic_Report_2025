import os
from openpyxl import load_workbook


def create_table15(subareaPath):
    balanceadoPath = os.path.join(subareaPath, "Balanceado")
    for tipicidad in ["Tipico", "Atipico"]:
        tipicidadPath = os.path.join(balanceadoPath, tipicidad)
        scenariosList = os.listdir(tipicidadPath)
        scenariosList = [file for file in scenariosList if not file.endswith(".ini")]