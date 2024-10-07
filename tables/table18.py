import os
from pathlib import Path
import xml.etree.ElementTree as ET
from unidecode import unidecode
from openpyxl import load_workbook
import pandas as pd

#docx
from docxtpl import DocxTemplate
from docxcompose.composer import Composer
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell

horarios = {
    "Tipico": [
    "00:00 - 05:00",
    "05:00 - 06:30",
    "06:30 - 10:30",
    "10:30 - 12:30",
    "12:30 - 15:00",
    "15:00 - 17:00",
    "17:00 - 22:00",
    "22:00 - 00:00",
    ],
    "Atipico": [
    "00:00 - 06:00",
    "06:00 - 12:00",
    "12:00 - 17:00",
    "17:00 - 22:00",
    "22:00 - 00:00",
    ],
}

def _translation_greens(sigPath: str) -> None:
    tree = ET.parse(sigPath)
    networkTag = tree.getroot()

    #Movements
    sgList = []
    for sg in networkTag.findall("./sgs/sg"):
        sgList.append(int(sg.get("id")))

    #Stages data
    dfStages = pd.DataFrame(
        columns=[int(stage.attrib["id"]) for stage in networkTag.findall('./stages/stage')],
        index=sgList,
    )

    #Filling data of stages and movements
    for stageTag in networkTag.findall("./stages/stage"):
        idStage = int(stageTag.get("id"))
        for activation in stageTag.findall("./activations/activation"):
            index = int(activation.get("sg_id"))
            value = True if activation.get("activation") == "ON" else False
            dfStages.loc[index, idStage] = value
    
    #Finding movements what starts in the beginning
    check = False
    for sg in sgList:
        if dfStages.loc[sg, 1] and not dfStages.loc[sg, dfStages.columns[-1]]: #Only in the first phase
            selected_sg = sg
            check = True
            break

    if not check:
        print(f"Mensaje para el más feo: No hay movimientos que acaben en la última fase y NO inicien en la primera fase")
        print("ERROR: ", sigPath)

    #Finding interstageprog for the selected signal group
    space = 0
    interstageprogTag = networkTag.findall("./interstageProgs/interstageProg")
    for sg in interstageprogTag[-1].findall("./sgs/sg"):
        idSg = int(sg.get("sg_id"))
        if selected_sg == idSg:
            for cmd in sg.findall("./cmds/cmd"):
                if cmd.get("display") == "3":
                    space += int(cmd.get("begin"))//1000
                    break
    
    for stageProg in networkTag.findall("./stageProgs/stageProg"):
        interstage = stageProg.findall("./interstages/interstage")
        if len(interstage) == len(dfStages.columns):
            cycleTime = int(stageProg.attrib["cycletime"])//1000
            lastTime = int(interstage[-1].attrib["begin"])//1000
            upperLimit = cycleTime - space
            movement = upperLimit - lastTime
            for interstage in stageProg.findall("./interstages/interstage"):
                value = interstage.attrib["begin"]
                #Para Yeik: Sub Area 043
                interstage.attrib["begin"] = str(int((int(value)//1000 + movement)*1000))

    ET.indent(tree)
    tree.write(sigPath, encoding="UTF-8", xml_declaration=True)

def _create_from_excel(sig_path, scenarioValue, tipicidadValue, wb):
    codigo = os.path.split(sig_path)[1][:-4]
    ws = wb[codigo]

    listSlices = [
        (slice("V2", "AJ2"), "HPMAD", "Tipico"),
        (slice("V3", "AJ3"), "HVMAD", "Tipico"),
        (slice("V5", "AJ5"), "HVM", "Tipico"),
        (slice("V7", "AJ7"), "HVT", "Tipico"),
        (slice("V9", "AJ9"), "HVN", "Tipico"),  
        (slice("V10", "AJ10"), "HVMAD", "Atipico"),
        (slice("V14", "AJ14"), "HVN", "Atipico"),
    ]

    for slicev, scenario, tipicidad in listSlices:
        if scenarioValue == scenario and tipicidad == tipicidadValue:
            rowList = [elem.value for row in ws[slicev] for elem in row if elem.value is not None] 
            greens = [elem for i, elem in enumerate(rowList) if i % 3 == 0]
            ambars = [elem for i, elem in enumerate(rowList) if i % 3 == 1]
            reds = [elem for i, elem in enumerate(rowList) if i % 3 == 2]
            sig_info = {
                "sig_name": codigo,
                "turn": scenario,
                "tipicidad": tipicidad,
                "cycle_time": sum(rowList),
                "offset": 0,
                "greens": greens,
                "ambars": ambars,
                "reds": reds
            }
            break
    return sig_info

def _combine_all_docx(filePathMaster, filePathsList, finalPath) -> None:
    number_of_sections = len(filePathsList)
    master = Document(filePathMaster)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document(filePathsList[i])
        composer.append(doc_temp)
    composer.save(finalPath)

def _create_data(sig_path: str, scenario: str, tipicidad: str) -> dict:
    #Traslación de verdes:
    _translation_greens(sig_path)

    tree = ET.parse(sig_path)
    sc_tag = tree.getroot()
    
    #Computing amber and red-red
    greens = []

    #Greens
    maxPhase = (None, 0) #NOTE: Index of stageProg / number of phases
    for i, stageProg in enumerate(sc_tag.findall("./stageProgs/stageProg")):
        interstages = stageProg.findall("./interstages/interstage")
        numberPhases = len(interstages)
        if numberPhases > maxPhase[1]:
            maxPhase = (i, numberPhases)

    stageProg = sc_tag.findall("./stageProgs/stageProg")[maxPhase[0]]
    cycleTime = int(stageProg.get('cycletime'))//1000
    offset = int(stageProg.get('offset'))//1000
    for interstage in stageProg.findall("./interstages/interstage"):
        greens.append(int(interstage.get('begin'))//1000)

    firstValue = greens[0]
    greens = [y-x for x,y in zip(greens[:-1], greens[1:])]
    greens[:0] = [firstValue]

    decreaseGreens = []
    for interstageProg in sc_tag.findall("./interstageProgs/interstageProg"):
        check = False
        for sg in interstageProg.findall("./sgs/sg"):
            if check: break
            if sg.get("signal_sequence") == "1": continue
            for i, cmd in enumerate(sg.findall("./cmds/cmd")):
                if i == 0 and cmd.get("display") == "3": break
                if i == 1 and cmd.get("display") == "3":
                    decreaseGreens.append(int(cmd.get("begin"))//1000)
                    check = True
                    break

    #Ambers
    ambers = []
    for interstageProg in sc_tag.findall("./interstageProgs/interstageProg"):
        checkPhase = False
        for sg in interstageProg.findall("./sgs/sg"):
            fixedState = sg.find("./fixedstates/fixedstate")
            if fixedState is not None and not checkPhase:
                ambers.append(int(fixedState.get("duration"))//1000)
                checkPhase = True
                break
        if not checkPhase: ambers.append(0)

    #Modifying greens values:
    firstValue = greens[0]
    greens = [y-x for x,y in zip(decreaseGreens, greens[1:])]
    greens[:0] = [firstValue]

    #Reds
    reds = [y-x for x,y in zip(ambers, decreaseGreens)]

    sig_info = {
        "sig_name": os.path.split(sig_path)[1][:-4],
        "turn": scenario,
        "tipicidad": tipicidad,
        "cycle_time": cycleTime,
        "offset": offset,
        "greens": greens,
        "ambars": ambers,
        "reds": reds,
    }

    return sig_info

def _create_table(sigs_info, tipicidad, tablasPath) -> None:
    doc = Document()
    maximum = 0
    for sigInfo in sigs_info:
        valueLen = len(sigInfo['greens'])
        if valueLen > maximum: maximum = valueLen

    len_greens = maximum

    # print("TAMAÑO DE VASES DE VERDE:", len_greens)
    table = doc.add_table(rows = 1, cols= 1+4+len_greens*3)
    table.cell(0,0).text = "Int."
    table.cell(0,1).text = "N° Plan"
    table.cell(0,2).text = "Horario"
    table.cell(0,3).text = "Desfase"
    table.cell(0,4).text = "T.C."
    for i in range(len_greens):
        table.cell(0,4+1+3*i).text = f'Fase {i+1}'
        table.cell(0,4+2+3*i).text = f'A'
        table.cell(0,4+3+3*i).text = f'RR'

    for i, sig_info in enumerate(sigs_info):
        new_row = table.add_row()
        if i == 0:
            new_row.cells[0].text = sig_info['sig_name'] #Nombre de la intersección
        new_row.cells[1].text = f"{i+1}" #N° Plan
        new_row.cells[2].text = horarios[tipicidad][i] #00:00 - 05:00
        new_row.cells[3].text = f"{sig_info['offset']}" #Desfase
        new_row.cells[4].text = f"{sig_info['cycle_time']}" #Tiempo de Ciclo
        #Repartos:
        for (j, greens), ambars, reds in zip(enumerate(sig_info['greens']),sig_info['ambars'],sig_info['reds']):
            new_row.cells[4+1+3*j].text = f"{greens}"
            new_row.cells[4+2+3*j].text = f"{ambars}"
            new_row.cells[4+3+3*j].text = f"{reds}"
        if i == 8:
            new_row = table.add_row()
            new_row.cells[0].text = "DÍA ATÍPICO"
            new_row.cells[0].merge(new_row.cells[len(new_row.cells)-1])
    
    table.cell(1,0).merge(table.cell(i+1,0))

    table.style = 'Table Grid'
    table.cell(0,0).width = Cm(1.75)
    table.cell(0,1).width = Cm(1)
    table.cell(0,2).width = Cm(3)

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                try:
                    run = paragraph.runs[0]
                except:
                    continue
                run.font.name = 'Arial Narrow'
                run.font.size = Pt(11)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for i in range(len(table.columns)):
        cell = table.cell(0,i)
        cell_xml_element = cell._tc
        table_cell_properties = cell_xml_element.get_or_add_tcPr()
        shade_obj = OxmlElement('w:shd')
        shade_obj.set(qn('w:fill'),'B4C6E7')
        table_cell_properties.append(shade_obj)

    for i in range(len(table.columns)):
        cell = table.cell(0,i)
        for paragraph in cell.paragraphs:
            run = paragraph.runs[0]
            run.font.bold = True

    finalPath = os.path.join(tablasPath, f"table18_{sig_info['sig_name']}_{unidecode(tipicidad).upper()}.docx")
    doc.save(finalPath)
    return finalPath, sig_info['sig_name'], tipicidad

def create_table18(subarea_path) -> None:
    #Reading each folder
    tablasPath = os.path.join(subarea_path, "Tablas")
    output_folder = Path(subarea_path) / "Output_Base"
    tipicidades = ["Tipico", "Atipico"]

    listData = []
    scenarioByTipicidad = {
        "Tipico": ["HPMAD", "HVMAD", "HPM", "HVM", "HPT", "HVT", "HPN", "HVN"],
        "Atipico": ["HVMAD", "HPM", "HPT", "HPN", "HVN"],
    }

    #Program Results:
    programResultsPath = Path(subarea_path) / "Program_Results.xlsx"
    wb = load_workbook(programResultsPath, read_only=True, data_only=True)

    for tipicidad in tipicidades:
        for i, scenario in enumerate(scenarioByTipicidad[tipicidad]):
            if scenario in ["HPMAD", "HVMAD", "HVM", "HVT", "HVN"]: continue
            scenario_path = output_folder / tipicidad / scenario
            sig_files = os.listdir(scenario_path)
            sig_files = [file for file in sig_files if file.endswith(".sig")]
            break

        for sig_file in sig_files:
            sigs_info = []
            for scenario in scenarioByTipicidad[tipicidad]:
                sig_path = output_folder / tipicidad / scenario / sig_file
                if scenario in ["HPMAD", "HVMAD", "HVM", "HVT", "HVN"]:
                    sig_info = _create_from_excel(sig_path, scenario, tipicidad, wb)
                else:
                    sig_info = _create_data(sig_path, scenario, tipicidad)
                sigs_info.append(sig_info)

            # print("Analizando Sig File:", sig_file)
            # for elem in sigs_info: print(elem)
            #HACK: Existe la posibilidad de que haya problemas en la creación de la tablas por el tamaño de las fases entre horas valles y puntas.
            finalPath, code, tipicidad = _create_table(sigs_info, tipicidad, tablasPath) #TODO: Analizar si funciona con tamaños de fases distintos.
            if tipicidad == "Tipico":
                tipicidadTxt = "típico"
            elif tipicidad == "Atipico":
                tipicidadTxt = "atípico"
            texto = f"Programación semafórica de la intersección {code} día {tipicidadTxt}"
            listData.append((texto, finalPath, code, tipicidad))

    wb.close()

    listWordPaths = []
    for text, pathTable, code, tipicidad in listData:
        doc_template = DocxTemplate("./templates/template_tablas4.docx")
        new_table = doc_template.new_subdoc(pathTable)
        doc_template.render({
            "texto": text,
            "tabla": new_table,
        })
        refPath = os.path.join(
            tablasPath,
            f"table18_{code}_{unidecode(tipicidad).upper()}_REF.docx"
        )
        doc_template.save(refPath)
        listWordPaths.append(refPath)

    programPath = os.path.join(subarea_path, "Tablas", "Programs.docx")
    filePathMaster = listWordPaths[0]
    filePathList = listWordPaths[1:]
    _combine_all_docx(filePathMaster, filePathList, programPath)

    return programPath

# if __name__ == '__main__':
#     sigPath = r"C:\Users\dacan\OneDrive\Desktop\PRUEBAS\Maxima Entropia\04 Proyecto Universitaria (37 Int. - 19 SA)\6. Sub Area Vissim\Sub Area 034\Output_Base\Tipico\HPT\SS-87.sig"
#     _create_data(sigPath, "HPM", "Tipico")