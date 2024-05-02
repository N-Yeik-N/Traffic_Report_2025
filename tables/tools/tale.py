from openpyxl import load_workbook
import pandas as pd

def tale_by_excel(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb['Base Data']
    sentido_slice = slice("C7","C14")
    acceso_slice = slice("D7","D14")

    list_sentido = [row[0].value for row in ws[sentido_slice] if row[0].value != None]
    list_acceso = [row[0].value for row in ws[acceso_slice] if row[0].value != None]

    if len(list_acceso) != len(list_sentido): print(f"Error - No coinciden cantidades de sentido-acceso:\n{path}")

    sentido_acceso_dict = {}
    for sent, access in zip(list_sentido, list_acceso):
        sentido_acceso_dict[sent] = access

    quantity = len(list_acceso)

    dict_turns = {}
    list_slices = [
        slice("D20","K59"),
        slice("O20","V59"),
        slice("AK20","AR59"),
    ]

    for i, turn in enumerate(["Mañana", "Tarde", "Noche"]):
        dict_info = _get_statistics(list_slices[i], ws, quantity)
        dict_turns[turn] = dict_info

    wb.close()

    dict_info = {
        "Datos": sentido_acceso_dict,
        "Numeros": dict_turns
    }
    
    return dict_info

def _get_statistics(slice_turn, ws, quant):
    df = pd.DataFrame([[cell.value for cell in row] for row in ws[slice_turn]]).iloc[:,:quant]
    means = df.mean(axis=0, skipna=True).round(2)
    maxs = df.max(axis=0, skipna=True).round(2)
    stds = df.std(axis=0, skipna=True).round(2)
    dict_info = {
        "max": maxs,
        "mean": means,
        "std": stds,
    }
    return dict_info