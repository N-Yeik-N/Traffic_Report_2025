import os
import pandas as pd

project_names = {
    1: "MEJORAMIENTO Y AMPLIACION DE LA INFRAESTRUCTURA SEMAFORICA DEL DISTRITO DE SANTIAGO DE SURCO - PROVINCIA DE LIMA - DEPARTAMENTO DE LIMA",
    2: "MEJORAMIENTO Y AMPLIACION DE LA INFRAESTRUCTURA SEMAFORICA DEL DISTRITO DE SANTIAGO DE SURCO - PROVINCIA DE LIMA - DEPARTAMENTO DE LIMA",
    3: "MEJORAMIENTO DEL SERVICIO DE TRANSITABILIDAD DE LA RED SEMAFORICA DE LOS EJES VIALES: AV. PIRAMIDE DEL SOL, AV. CHINCHAYSUYO, AV. GRAN CHIMU, AV. RIVA AGUERO, AV. ANCASH, AV. CESAR VALLEJO, AV. LURIGANCHO, AV. PORTADA DEL SOL, EN LOS DISTRITOS DE EL AGUSTINO Y SAN JUAN DE LURIGANCHO DE LA PROVINCIA DE LIMA - DEPARTAMENTO DE LIMA",
    4: "MEJORAMIENTO Y AMPLIACION DEL SERVICIO DE TRANSITABILIDAD DE LA RED SEMAFORICA DE LOS EJES VIALES: AV. DEFENSORES DEL MORRO, AV. MIGUEL GRAU, AV. PEDRO DE OSMA, AV. SAN MARTIN, AV. EL SOL OESTE, CA. TEODOCIO PARREÑO, AV. LIMA, AV. GUARDIA CIVIL, AV. CHORRILLOS, AV. ARIOSTO MATELLINI, AV. ALAMEDA SUR, AV. EL SOL, AV. ALAMEDA SAN MARCOS, AV. GUARDIA PERUANA, AV. MARISCAL CASTILLA, EN LOS DISTRITOS DE CHORRILLOS Y BARRANCO DE LA PROVINCIA DE LIMA - DEPARTAMENTO DE LIMA",
    5: "MEJORAMIENTO DE LA RED SEMAFORICA EN LAS INTERSECCIONES DEL EJE VIAL DE LA AV. UNIVERSITARIA (TRAMO: AV. SANTA ELVIRA - AV. LA PAZ) DE LOS DISTRITOS DE LIMA, SAN MARTIN DE PORRES, LOS OLIVOS Y EL DISTRITO DE PUEBLO LIBRE - PROVINCIA DE LIMA - DEPARTAMENTO DE LIMA",
    6: "MEJORAMIENTO DE LA RED SEMAFORICA EN LAS INTERSECCIONES DEL EJE VIAL DE LA AV. JAVIER PRADO - AV. FAUSTINO SANCHEZ CARRION - AV. LA MARINA, DISTRITO DE LA MOLINA, SANTIAGO DE SURCO, JESUS MARIA, SAN ISIDRO, MAGDALENA DEL MAR Y DISTRITO DE SAN MIGUEL - PROVINCIA DE LIMA - DEPARTAMENTO DE LIMA",
    7: "MEJORAMIENTO Y AMPLIACION DE LA RED SEMAFORICA DE LA CICLOVIA DE LA ZONA CENTRO I Y II, DE LOS DISTRITOS DE LIMA, JESUS MARIA, PUEBLO LIBRE, LA VICTORIA Y DISTRITO DE BREÑA - PROVINCIA DE LIMA - DEPARTAMENTO DE LIMA",
    8: "MEJORAMIENTO Y AMPLIACION DE LA RED SEMAFORICA DE LOS EJES VIALES: AV. SALVADOR ALLENDE, AV. SAN JUAN, AV. CESAR CANEVARO, AV. MIGUEL IGLESIAS, AV. 26 DE NOVIEMBRE, AV. JOSE CARLOS MARIATEGUI, AV. LIMA, AV. PACHACUTEC, EN LOS DISTRITOS DE SAN JUAN DE MIRAFLORES Y VILLA MARIA DEL TRIUNFO DE LA PROVINCIA DE LIMA - DEPARTAMENTO EN LOS DISTRITOS DE VILLA MARIA DEL TRIUNFO Y SAN JUAN DE MIRAFLORES DE LA PROVINCIA DE LIMA - DEPARTAMENTO DE LIMA",
    9: "MEJORAMIENTO Y AMPLIACION DE LA RED SEMAFORICA DE LOS EJES VIALES: AV. REVOLUCION, AV. MARIANO PASTOR SEVILLA, AV. MICAELA BASTIDAS, AV. JUAN VELASCO ALVARADO, AV. CENTRAL, AV. 200 MILLAS, AV. 1° DE MAYO, AV. SEPARADORA INDUSTRIAL, DEL DISTRITO DE VILLA EL SALVADOR - PROVINCIA DE LIMA - DEPARTAMENTO DE LIMA",
    10: "MEJORAMIENTO DE LA RED SEMAFORICA DE LOS EJES VIALES: AV. LA MOLINA, AV. LA UNIVERSIDAD, AV. RAÚL FERRERO, AV. SIETE, AV. MANUEL PRADO UGARTECHE, AV. ALAM. DEL CORREGIDOR, AV. LOS FRESNOS, AV. LOS CONSTRUCTORES, AV. SEPARADORA INDUSTRIAL EN LOS DISTRITOS DE LA MOLINA Y SANTA ANITA DE LA PROVINCIA DE LIMA - DEPARTAMENTO DE LIMA",
}

def location(path_subarea):
    numsubarea = os.path.split(path_subarea)[1][-3:]
    df_general = pd.read_excel("./data/Datos Generales.xlsx", sheet_name="DATOS", header=0, usecols="A:E")
    nro_entregable = df_general[df_general['Sub_Area'] == int(numsubarea)]["Entregable"].unique()[0]
    nameproject = project_names[nro_entregable]
    
    nomdistrito = df_general[df_general['Sub_Area'] == int(numsubarea)]["Distrito"].unique()[0]
    intersecciones = df_general[df_general['Sub_Area'] == int(numsubarea)]["Interseccion"].unique().tolist()
    codintersecciones = df_general[df_general['Sub_Area'] == int(numsubarea)]["Code"].unique().tolist()
    if len(intersecciones) == 1:
        presinter = "presenta la ubicacion de la intersección"
    else:
        presinter = "presentan las ubicaciones de las siguientes intersecciones:"

    texto = ""
    for i, nombre_inter in enumerate(intersecciones):
        if i == len(intersecciones)-1:
            texto += nombre_inter
        else:
            texto += nombre_inter +', '

    nominterseccion = texto

    texto = ""
    for i, code_inter in enumerate(codintersecciones):
        if i == len(codintersecciones)-1:
            texto += code_inter
        else:
            texto += code_inter+', '
    codinterseccion = texto

    if len(intersecciones) > 1:
        descsubarea = "las intersecciones pertenecientes"
    else:
        descsubarea = "la intersección perteneciente"

    VARIABLES = {
        "numsubarea": numsubarea,
        "nameproject": nameproject,
        "nomdistrito": nomdistrito,
        "presinter": presinter,
        "nominterseccion": nominterseccion,
        "codinterseccion": codinterseccion,
        "descsubarea": descsubarea,
    }

    return VARIABLES