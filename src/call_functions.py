import os
import pandas as pd
from pdfs.tools import *
import re
from docxcompose.composer import Composer

project_names = {
    1: "Mejoramiento y ampliación de la infraestructura semafórica del distrito de Santiago de Surco - Provincia de Lima - Departamento de Lima",
    2: "Mejoramiento y ampliación de la infraestructura semafórica del distrito de Santiago de Surco - Provincia de Lima - Departamento de Lima",
    3: "Mejoramiento del servicio de transitabilidad de la red semafórica de los ejes viales: Av. Piramide del Sol, Av. Chinchaysuyo, Av. Gran Chimu, Av. Riva Aguero, Av. Ancash, Av. Cesar Vallejo, Av. Luringancho, Av. Portada del Sol, en los distritos de El Agustino y San Juan de Lurigancho de la Provincia de Lima - Departamento de Lima",
    4: "Mejoramiento y ampliación del servicio de transitabilidad de la red semafórica de los ejes viale: Av. Defensores del Morro, Av. Miguel Grau, Av. Pedro de Osma, Av. San Martin, Av. El Sol Oeste, Ca. Teodocio Parreño, Av. Lima, Av. Guardia Civil, Av. Chorrillos, Av. Ariosto Matellini, Av. Alameda Sur, Av. El Sol, Av. Alameda San Marcos, Av. Guardia Peruana, Av. Mariscal Castilla, en los distritos de Chorrillos y Barranco de la Provincia de Lima - Departamento de Lima",
    5: "Mejoramiento de la red semafórica en las intersecciones del eje vial de La Av. Universitaria (Tramo: Av. Santa Elvira - Av. La Paz) de los distritos de Lima, San Martin de Porres, Los Olivos y el distrito de Pueblo Libre - Provincia de Lima - Departamento de Lima",
    6: "Mejoramiento de la red semafórica en las intersecciones del eje vial de La Av. Javier Prado - Av. Faustino Sánchez Carrión - Av. La Marina, distrito de La Molina, Santiago de Surco, Jesús María, San Isidro, Magdalena del Mar y distrito de San Miguel - Provincia de Lima - Departamento de Lima",
    7: "Mejoramiento y ampliación de la red semafórica de la ciclovía de la zona centro I y II, de los distritos de Lima, Jesús María, Pueblo Libre, La Victoria y distrito de Breña - Provincia de Lima - Departamento de Lima",
    8: "Mejoramiento y ampliación de la red semafórica de los ejes viale: Av. Salvador Allende, Av. San Juan, Av. Cesar Canevaro, Av. Miguel Iglesias, Av. 26 de Noviembre, Av. Jose Carlos Mariategui, Av. Lima, Av. Pachacutec, en los distritos de San Juan de Miraflores y Villa María del Triunfo de la Provincia de Lima - Departamento de Lima",
    9: "Mejoramiento y ampliación de la red semafórica de los ejes viale: Av. Revolución, Av. Mariano Pastor Sevilla, Av. Micaela Bastidas, Av. Juan Velasco Alvarado, Av. Central, Av. 200 Millas, Av. 1° de Mayo, Av. Separadora Industrial, del distrito de Villa El Salvador - Provincia de Lima - Departamento de Lima",
    10: "Mejoramiento de la red semafórica de los ejes viales: Av. La Molina, Av. La Universidad, Av. Raúl Ferrero, Av. Siete, Av. Manuel Prado Ugarteche, Av. Alam. Del Corregidor, Av. Los Fresnos, Av. Los Constructores, Av. Separadora Industrial en los distritos de La Molina y Santa Anita de la Provincia de Lima - Departamento de Lima",
}

def _combine_all_docx(filePathMaster, filePathsList, finalPath) -> None:
    number_of_sections = len(filePathsList)
    master = Document(filePathMaster)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document(filePathsList[i])
        composer.append(doc_temp)

    composer.save(finalPath)

def location(path_subarea) -> list[dict, list]:
    numsubarea = os.path.split(path_subarea)[1][-3:]
    df_general = pd.read_excel("./data/Datos Generales.xlsx", sheet_name="DATOS", header=0, usecols="A:E")
    nro_entregable = df_general[df_general['Sub_Area'] == int(numsubarea)]["Entregable"].unique()[0]
    nameproject = project_names[nro_entregable]
    
    nomdistrito = df_general[df_general['Sub_Area'] == int(numsubarea)]["Distrito"].unique()[0]
    intersecciones = df_general[df_general['Sub_Area'] == int(numsubarea)]["Interseccion"].unique().tolist()
    codintersecciones = df_general[df_general['Sub_Area'] == int(numsubarea)]["Code"].unique().tolist()
    if len(intersecciones) == 1:
        presinter = "presenta la ubicacion de la intersección"
        presinter2 = "la intersección"
    else:
        presinter = "presentan las ubicaciones de las siguientes intersecciones:"
        presinter2 = "las intersecciones"

    texto = ""
    if len(intersecciones) == 1:
        nominterseccion = intersecciones[0]
        presinter3 = "El nodo de la intersección"
    else:
        presinter3 = "Todos los nodos de las intersecciones"
        for i, nombre_inter in enumerate(intersecciones):
            if i == len(intersecciones)-1:
                texto += ' y ' + nombre_inter
            elif i == len(intersecciones)-2:
                texto += nombre_inter
            else:
                texto += nombre_inter +', '

        nominterseccion = texto

    texto = ""
    if len(codintersecciones) == 1:
        codinterseccion = codintersecciones[0]
    else:
        for i, code_inter in enumerate(codintersecciones):
            if i == len(codintersecciones)-1:
                texto += ' y ' + code_inter
            elif i == len(codintersecciones)-2:
                texto += code_inter
            else:
                texto += code_inter+', '
        codinterseccion = texto

    if len(intersecciones) > 1:
        descsubarea = "las intersecciones pertenecientes"
        prestablas = "las siguientes tablas"
    else:
        descsubarea = "la intersección perteneciente"
        prestablas = "la siguiente tabla"

    VARIABLES = {
        "numsubarea": numsubarea,
        "nameproject": nameproject,
        "nomdistrito": nomdistrito,
        "presinter": presinter,
        "nominterseccion": nominterseccion,
        "codinterseccion": codinterseccion,
        "descsubarea": descsubarea,
        "presinter2": presinter2,
        "prestablas": prestablas,
        "presinter3": presinter3,
    }

    return VARIABLES, codintersecciones

def histogramas(path_subarea) -> str:
    listCodes = get_codes(path_subarea)
    anexos_path = os.path.join(path_subarea, "Anexos")

    folderAnexos = os.listdir(anexos_path)

    assert "Vehicular" in folderAnexos, "ERROR: No se encontro el archivo 'Vehicular' en la carpeta 'Anexos'"
    assert "Peatonal" in folderAnexos, "ERROR: No se encontro el archivo 'Peatonal' en la carpeta 'Anexos'"

    dictContentFolders = {
        'Vehicular': os.path.join(anexos_path, "Vehicular"),
        'Peatonal': os.path.join(anexos_path, "Peatonal"),
    }

    histogramaTotal_path = []
    for agentType, pathFolder in dictContentFolders.items():
        listPDFS = os.listdir(pathFolder)

        pdfs_by_code = {}
        for code in listCodes:
            pdfs_by_code[code] = []

        pattern1 = r"([A-Z]+[0-9]+)"
        pattern2 = r"([A-Z]+-[0-9]+)"
        for pdf in listPDFS:
            match_pdf = re.search(pattern1, pdf) or re.search(pattern2, pdf)
            if match_pdf:
                code_str = match_pdf[1]
                pdfs_by_code[code_str].append(pdf)

        listSelectedPDF = []
        for code, pdfs in pdfs_by_code.items():
            for pdf in pdfs:
                if 'Histograma' in pdf:
                    listSelectedPDF.append((code, os.path.join(pathFolder, pdf)))

        pattern1 = r"(_A)"
        pattern2 = r"(_T)"
        dataPDFs = []
        for code, pdf_path in listSelectedPDF:
            namePDF = os.path.split(pdf_path)[1]
            namePDF = namePDF[:-4]
            match_tipicidad = re.search(pattern1, namePDF) or re.search(pattern2, namePDF)
            if match_tipicidad:
                tipicidad = match_tipicidad[1][1] #[1] "_T" o "_A" / [1][1] "T" o "A"
            if pdf_path.endswith(".png"):
                #print("Ya existe el PDF en .png, si hay correcciones borrar:\n", pdf_path)
                continue
            dataPDFs.append([
                code,
                tipicidad,
                convert_pdf_to_image(pdf_path, pathFolder, namePDF),
            ])

        dataPDFs = sorted(dataPDFs, key = lambda x: (x[0], -ord(x[1])))

        histograma_path = create_histogramas_subdocs(dataPDFs, path_subarea, agentType)
        histogramaTotal_path.append(histograma_path)

    finalPath = os.path.join(path_subarea, "Tablas", "histogramas.docx")
    filePathMaster = histogramaTotal_path[0]
    filePathsList = histogramaTotal_path[1:]

    _combine_all_docx(filePathMaster, filePathsList, finalPath)

    return finalPath

def flujogramas_vehiculares(path_subarea) -> str:
    listCodes = get_codes(path_subarea)
    anexos_path = os.path.join(path_subarea, "Anexos")

    folderAnexos = os.listdir(anexos_path)

    assert "Vehicular" in folderAnexos, "ERROR: No se encontro el archivo 'Vehicular' en la carpeta 'Anexos'"

    folderVehicular = os.path.join(anexos_path, "Vehicular")
    listPDFS = os.listdir(folderVehicular)

    pdfs_by_code = {}
    for code in listCodes:
        pdfs_by_code[code] = []

    pattern1 = r"([A-Z]+[0-9]+)"
    pattern2 = r"([A-Z]+-[0-9]+)"
    for pdf in listPDFS:
        match_pdf = re.search(pattern1, pdf) or re.search(pattern2, pdf)
        if match_pdf:
            code_str = match_pdf[1]
            pdfs_by_code[code_str].append(pdf)

    listSelectedPDF = []
    listCodes = []
    for code, pdfs in pdfs_by_code.items():
        for pdf in pdfs:
            if 'V_Ma_T' in pdf:
                listSelectedPDF.append((code, os.path.join(folderVehicular, pdf)))
                listCodes.append(code)

    dataInfo = []
    for code, pdf_path in listSelectedPDF:
        if pdf_path.endswith('.png'):
            #print("Ya existe el PDF en .png, si hay correcciones, borrar:\n", pdf_path)
            continue
        namePDF = os.path.split(pdf_path)[1]
        namePDF = namePDF[:-4]
        dataInfo.append([
            code,
            convert_pdf_to_image(pdf_path, folderVehicular, namePDF),
        ])
        #print("PDF convertido a imagen:", namePDF)

    flujograma_path = create_flujogramas_vehicular_subdocs(dataInfo, path_subarea)

    return flujograma_path

def flujogramas_peatonales(path_subarea) -> str:
    listCodes = get_codes(path_subarea)
    anexos_path = os.path.join(path_subarea, "Anexos")

    folderAnexos = os.listdir(anexos_path)

    if not "Peatonal" in folderAnexos:
        print("ERROR: No se encontro el archivo 'Vehicular' en la carpeta 'Anexos'")

    folderPeatonal = os.path.join(anexos_path, "Peatonal")
    listPDFS = os.listdir(folderPeatonal)

    pdfs_by_code = {}
    for code in listCodes:
        pdfs_by_code[code] = []

    pattern1 = r"([A-Z]+[0-9]+)"
    pattern2 = r"([A-Z]+-[0-9]+)"
    for pdf in listPDFS:
        match_pdf = re.search(pattern1, pdf) or re.search(pattern2, pdf)
        if match_pdf:
            code_str = match_pdf[1]
            pdfs_by_code[code_str].append(pdf)

    listSelectedPDF = []
    listCodes = []
    for code, pdfs in pdfs_by_code.items():
        for pdf in pdfs:
            if "Turno 01_T" in pdf:
                listSelectedPDF.append((code, os.path.join(folderPeatonal, pdf)))
                listCodes.append(code)
    listCodes = list(set(listCodes))

    listPathImages = {}

    for code in listCodes:
        listPathImages[code] = []

    #Checking if there are images
    for code in listCodes:
        for codePath, pdfPath in listSelectedPDF:
            if code == codePath:
                if pdfPath.endswith('.png'):
                    listPathImages[code].append(pdfPath)
                    break

    #In case there are no .png files
    listPngImages = listPathImages.copy()
    for code, listDocuments in listPngImages.items():
        if len(listDocuments) == 0: #There is no .pngs
            for codePDF, pdfPath in listSelectedPDF:
                if code == codePDF:
                    namePDF = os.path.split(pdfPath)[1]
                    namePDF = namePDF[:-4]
                    listPathImages[code] = convert_pdf_to_image(pdfPath, folderPeatonal, namePDF)
    
    resultList = []
    for code, imagePath in listPathImages.items():
        resultList.append((code, imagePath[0]))

    flujograma_path = create_flujograma_peatonal_subdocs(resultList, path_subarea)

    return flujograma_path