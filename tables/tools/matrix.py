from openpyxl import load_workbook
import pandas as pd

def read_matrix(excel_path):
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb.active
    
    sliceOrigin = slice("B1","AZ1") #Maximum 50 origins
    sliceDestiny = slice("A2", "A51") #Maximum 50 detinys

    origins = ['O'+str(elem.value) for row in ws[sliceOrigin] for elem in row if elem.value != None]
    destinys = ['D'+str(row[0].value) for row in ws[sliceDestiny]]

    numCols = len(origins)
    numRows = len(destinys)

    MATRIX = []
    for row in range(2,numRows+2):
        ROW = []
        for col in range(2,numCols+2):
            cell_value = str(ws.cell(row=row, column=col).value) if ws.cell(row=row, column=col).value != None else ""
            ROW.append(cell_value)
        MATRIX.append(ROW)
    wb.close()

    return origins, destinys, MATRIX