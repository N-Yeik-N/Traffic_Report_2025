from docxtpl import DocxTemplate
from tables.table1 import create_table1
from tables.table2n3 import create_tables2n3
from tables.table4n5 import create_table4n5
from tables.table6 import create_table6
from tables.table7 import create_table7
from tables.table8n9 import create_table8, create_table9
from tables.table12 import create_table12
from tables.table14 import create_table14
from tables.table17 import create_table17
from tables.table18 import create_table18
from src.call_functions import *
from src.histogramas import *
from src.changer_dates import change_peakhours
from sigs.sig_actual import get_sigs_actual
from images.resultados import create_resultados_images
from results.reading_json import generate_results
from conclusions.table23 import create_table23
from pdfs.flujogramas import *

import logging
from pathlib import Path
import win32com.client as com
from tqdm import tqdm

#Interface
from ui.interface import Ui_Form
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
from PyQt5 import QtCore

#Dates
import datetime
import locale

locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
fecha_actual = datetime.datetime.now()
month = fecha_actual.strftime("%B")

LOGGER = logging.getLogger(__name__)
LOGGER.setLevel(logging.DEBUG)
f = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

class MyWindow(QMainWindow, Ui_Form):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)

        self.ui.openPushButton.clicked.connect(self.open_file)
        self.ui.startPushButton.clicked.connect(self.start)
        self.ui.pushButtonChecked.clicked.connect(self.check_all_items)
        self.ui.pushButtonUnchecked.clicked.connect(self.uncheck_all_items)
        self.ui.pushButtonPeakhour.clicked.connect(self.changer_hours)
    
    def open_file(self):
        self.path_subarea = QFileDialog.getExistingDirectory(self, 'Open File')
        if self.path_subarea:
            self.ui.pathLineEdit.setText(self.path_subarea)
            #Logger location
            log_path = os.path.join(self.path_subarea, "logs")

            if not os.path.exists(log_path):
                os.mkdir(log_path)

            fh = logging.FileHandler(os.path.join(log_path, "report.log"), mode='w')
            fh.setFormatter(f)
            LOGGER.addHandler(fh)

    def changer_hours(self):
        pathParts = list(Path(self.path_subarea).parts)
        subarea = pathParts[-1]
        projectPath = Path(*pathParts[:-2])
        vehicularFieldDataPath = projectPath / "7. Informacion de Campo" / subarea / "Vehicular"
        excel = com.Dispatch('Excel.Application')
        excel.Visible = False

        for tipicidad in ["Tipico", "Atipico"]:
            typicallyPath = vehicularFieldDataPath / tipicidad
            listExcels = os.listdir(typicallyPath)
            listExcels = [file for file in listExcels if file.endswith(".xlsm") and not file.startswith("~$")]
            print(f"{f' Abriendo {tipicidad} ':#^{50}}")

            for excelFile in tqdm(listExcels, f"{tipicidad}"):
                excelPath = typicallyPath / excelFile
                try:
                    change_peakhours(excel, excelPath)
                except Exception as inst:
                    print("Error: ", inst)
                    print("Excel: ", excelFile)
                    continue

        excel.Quit()
        print(f"{f' FINALIZADO ':#^{50}}")
                    
    def check_all_items(self):
        row_count = self.ui.tableWidget.rowCount()
        for row in range(row_count):
            item = self.ui.tableWidget.item(row, 0)
            if item:
                item.setCheckState(QtCore.Qt.Checked)

    def uncheck_all_items(self):
        row_count = self.ui.tableWidget.rowCount()
        for row in range(row_count):
            item = self.ui.tableWidget.item(row, 0)
            if item:
                item.setCheckState(QtCore.Qt.Unchecked)

    def start(self): 
        nameSubArea = os.path.split(self.path_subarea)[1]
        print(f"##{'#'*len(nameSubArea)}##")
        print(f"# {nameSubArea} #")
        print(f"##{'#'*len(nameSubArea)}##")

        doc = DocxTemplate("templates/template.docx")

        #Location
        VARIABLES, codintersecciones = location(self.path_subarea)

        #Table paths:
        checkObject = self.ui.tableWidget.item(0,0).checkState()
        if checkObject: #NOTE: Ready tabla1
            try:
                table1_path = create_table1(self.path_subarea)
                table1 = doc.new_subdoc(table1_path)
                VARIABLES.update({"tabla1": table1})
                print("Tabla 1\t\tOK\tDatos generales de intersecciones y códigos")
            except Exception as e:
                print("Tabla 1:\t\tERROR\tDatos generales de intersecciones y códigos")
                LOGGER.warning("Error Tabla 1")
                LOGGER.warning(str(e))
                
        checkObject = self.ui.tableWidget.item(1,0).checkState()
        if checkObject: #NOTE: Ready tabla2n3
            try:
                table2_vehicular, table2_peatonal, table3_path, dcontet, dcontea = create_tables2n3(self.path_subarea)
                table2Vehicular = doc.new_subdoc(table2_vehicular)
                table2Peatonal = doc.new_subdoc(table2_peatonal)
                table3 = doc.new_subdoc(table3_path)
                VARIABLES.update({"tabla2_vehicular": table2Vehicular, "tabla2_peatonal": table2Peatonal, "tabla3": table3, "dcontet": dcontet, "dcontea": dcontea})
                print("Tabla 2\t\tOK\tTabla de las horas puntas")
                print("Tabla 3\t\tOK\tTabla de fechas de conteos")
            except Exception as e:
                print("Tabla 2\t\tERROR\tTabla de las horas puntas")
                print("Tabla 3\t\tERROR\tTabla de fechas de conteos")
                LOGGER.warning("Error Tabla 2 o 3")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(2,0).checkState()
        if checkObject: #NOTE: Ready tabla4
            try:
                table4_path, table5_path = create_table4n5(self.path_subarea)
                table4 = doc.new_subdoc(table4_path)
                VARIABLES.update({"tabla4": table4})
                print("Tabla 4\t\tOK\tFechas de toma de longitud de cola")
            except FileNotFoundError as e:
                print("Tabla 4\t\tERROR\tNo existen archivos de colas")
            except IndexError as e:
                print("Tabla 4\t\tERROR\tNo existen archivos de colas")
            except Exception as e:
                print("Tabla 4\t\tERROR\tFechas de toma de longitud de cola")
                LOGGER.warning("ERRROR Tabla 4")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(3,0).checkState()
        if checkObject: #NOTE: Ready tabla5
            try:
                table5 = doc.new_subdoc(table5_path)
                VARIABLES.update({"tabla5": table5})
                print("Tabla 5\t\tOK\tDatos estadísticas de longitud de cola")
            except FileNotFoundError as e:
                print("Tabla 5\t\tERROR\tNo existen archivos de colas")
            except UnboundLocalError as e:
                print("Tabla 5\t\tERROR\tNo existen archivos de colas")
            except Exception as e:
                print("Tabla 5\t\tERROR\tDatos estadísticas de longitud de cola")
                LOGGER.warning("Error Tabla 5")
                LOGGER.warning(str(e))
        
        checkObject = self.ui.tableWidget.item(4,0).checkState()
        if checkObject: #NOTE: Ready tabla6
            try:
                table6_path = create_table6(self.path_subarea)
                table6 = doc.new_subdoc(table6_path)
                VARIABLES.update({"tabla6": table6})
                print("Tabla 6\t\tOK\tTabla de tiempos de embarque y desembarque")
            except Exception as e:
                print("Tabla 6\t\tERROR\tTabla de tiempos de embarque y desembarque")
                LOGGER.warning("Error Tabla 6")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(5,0).checkState()
        if checkObject: #NOTE: Ready tabla7
            try:
                table7_path = create_table7(self.path_subarea)
                table7 = doc.new_subdoc(table7_path)
                VARIABLES.update({"tabla7": table7})
                print("Tabla 7\t\tOK\tDatos estadísticas de embarque y desembarque")
            except AttributeError as e:
                print("Tabla 7\t\tError\tHay un dato en blanco en algunos de los excels")
            except IndexError as e:
                try:
                    print(f"Tabla 7\t\tError\tDebes pegar la tabla manualmente:\n{table7_path}")
                except:
                    print("Tabla 7\t\tError\tNo existe datos de embarque y desembarque")
            except Exception as e:
                print("Tabla 7\t\tERROR\tDatos estadísticas de embarque y desembarque")
                LOGGER.warning("Error Tabla 7")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(6,0).checkState()
        if checkObject: #NOTE: Ready tabla8
            try:
                table8_path = create_table8(self.path_subarea)
                table8 = doc.new_subdoc(table8_path)
                VARIABLES.update({"tabla8": table8})
                print("Tabla 8\t\tOK\tTabla de fechas de tiempos de ciclo y fases")
            except Exception as e:
                print("Tabla 8:\t\tERROR\tTabla de fechas de tiempos de ciclo y fases")
                LOGGER.warning("Error Tabla 8")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(7,0).checkState()
        if checkObject: #NOTE: Ready table9 and parrafos_programacion
            try:    
                table9_path, parrafos_programacion_path = create_table9(self.path_subarea)
                table9 = doc.new_subdoc(table9_path)
                parrafos_programacion = doc.new_subdoc(parrafos_programacion_path)
                VARIABLES.update({"tabla9": table9, "parrafos_programacion": parrafos_programacion})
                print("Tabla 9\t\tOK\tGráficas de programaciones semafóricas")
            except Exception as e:
                print("Tabla 9\t\tERROR\tGráficas de programaciones semafóricas")
                LOGGER.warning("Error Tabla 9")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(8,0).checkState()
        if checkObject: #NOTE: Ready histogramas and maxTipicidad, maxTurno            
            try:
                histogramas_tipicos, histogramas_atipicos, histograma_path_tipico, histograma_path_atipico, sumvoltip_var, sumvolati_var, maxtipicidad, volturmanana, volturntarde, volturnnoche, maxturno = histogramas_vehiculares(self.path_subarea)
                #histograma_path = histogramas(self.path_subarea)
                histogramas_tip = doc.new_subdoc(histogramas_tipicos)
                histogramas_atip = doc.new_subdoc(histogramas_atipicos)
                histogramas_sist_tip = doc.new_subdoc(histograma_path_tipico)
                histogramas_sist_ati = doc.new_subdoc(histograma_path_atipico)
                VARIABLES.update({ #Check if you need these in strings or not
                    "histogramas_tip": histogramas_tip,
                    "histogramas_atip": histogramas_atip,
                    "histogramas_sist_tip": histogramas_sist_tip,
                    "histogramas_sist_ati": histogramas_sist_ati,
                    "sumvoltip": sumvoltip_var,
                    "sumvolati": sumvolati_var,
                    "maxtipicidad": maxtipicidad, #típico, atípico
                    "volturmanana": volturmanana,
                    "volturntarde": volturntarde,
                    "volturnnoche": volturnnoche,
                    "maxturno": maxturno #Mañana, Tarde, Noche
                    })
                print("Histograma\tOK\tVehiculares")
            except Exception as e:
                print("Histograma\tERROR\tVehiculares")
                LOGGER.warning("Errores de histogramas vehiculares")
                LOGGER.warning(str(e))

            try:
                histogramas_pea_tip, histogramas_pea_atip = histogramas_peatonales(self.path_subarea)
                histogramas_tip_pea = doc.new_subdoc(histogramas_pea_tip)
                histogramas_atip_pea = doc.new_subdoc(histogramas_pea_atip)
                VARIABLES.update({
                    "histogramas_tip_pea": histogramas_tip_pea,
                    "histogramas_atip_pea": histogramas_atip_pea,
                    })
                print("Histograma\tOK\tPeatonales")
            except Exception as e:
                print("Histograma\tERROR\tPeatonales")
                LOGGER.warning("Errores de histogramas peatonales")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(9,0).checkState()
        if checkObject: #NOTE: Ready flujograma_veh_sist and paragraphs
            try:
                flujograma_vehicular_path = flujograma_vehicular(self.path_subarea, maxturno, maxtipicidad)
                flujogvmt_cod_maxtip_maxturno = doc.new_subdoc(flujograma_vehicular_path)
                paragraph_flujograma_veh_path = create_paragraphs(self.path_subarea, maxtipicidad, maxturno)
                paragraphs_flujogramas_vehiculares = doc.new_subdoc(paragraph_flujograma_veh_path)
                VARIABLES.update({
                    "flujogvmt_cod_maxtip_maxturno": flujogvmt_cod_maxtip_maxturno,
                    "paragraphs_flujogramas_vehiculares": paragraphs_flujogramas_vehiculares,
                    }) 
                print("Flujogramas\tOK\tVehiculares")
            except Exception as e:
                print("Flujogramas\tERROR\tVehiculares")
                LOGGER.warning("Flujogramas vehiculares")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(10,0).checkState()
        if checkObject: #NOTE: Ready flujogramas peatonales
            try:
                flujograma_peatonal_path = flujogramas_peatonales(self.path_subarea, maxturno, maxtipicidad)
                flujograma_peatonal = doc.new_subdoc(flujograma_peatonal_path)
                VARIABLES.update({"flujograma_peat_max": flujograma_peatonal})
                print("Flujogramas\tOK\tPeatonales")
            except Exception as e:
                print("Flujogramas\tERROR\tPeatonales")
                LOGGER.warning("Flujogramas peatonales")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(11,0).checkState()
        if checkObject: #NOTE: Ready tabla12
            try:
                table12_path = create_table12(self.path_subarea)
                table12 = doc.new_subdoc(table12_path)
                VARIABLES.update({"tabla12": table12})
                print("Tabla 12\tOK\tTabla para ser llenada de interacciones peatonales")
            except Exception as e:
                print("Tabla 12\tERROR\tTabla para ser llenada de interacciones peatonales")
                LOGGER.warning("Error Tabla 12")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(12,0).checkState()
        if checkObject: #NOTE: Ready Tabla OD only when GEH-R2.xlsm exists
            try:
                table14_path, VARIABLES_OD = create_table14(self.path_subarea, maxturno, maxtipicidad)
                table14 = doc.new_subdoc(table14_path)
                VARIABLES.update({"tabla14": table14})
                VARIABLES.update(VARIABLES_OD)
                print("Tabla 14\tOK\tTabla de OD de situación actual")
            except Exception as e:
                print("Tabla 14\tERROR\tTabla de orígenes y destinos de situación actual")
                LOGGER.warning("Error Tabla 14")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(13,0).checkState()
        if checkObject: #NOTE: Ready tabla16
            try:
                table16_path = create_resultados_images(self.path_subarea)
                table16 = doc.new_subdoc(table16_path)
                VARIABLES.update({"tabla16": table16})
                print("Tabla 16\tOK\tImágenes de GEH y R2")
            except Exception as e:
                print("Tabla 16\tERROR\tImágenes de GEH y R2")
                LOGGER.warning("Error Tabla 16")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(14,0).checkState()
        if checkObject: #NOTE: Ready tabla17
            try:
                table17_path = create_table17(self.path_subarea)
                table17 = doc.new_subdoc(table17_path)
                VARIABLES.update({"tabla17": table17})
                print("Tabla 17\tOK\tResultados del GEH-R2")
            except Exception as e:
                print("Tabla 17\tERROR\tResultados del GEH-R2")
                LOGGER.warning("Error Tabla 17")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(15,0).checkState() #TODO: CHECK
        if checkObject: #TODO: Me parece que cambiará de nombre, es lo mismo que el 19.
            try: #Cambiar a solo horas punta
                table18_path = create_table18(self.path_subarea)
                table18 = doc.new_subdoc(table18_path)
                VARIABLES.update({"tabla18": table18})
                print("Tabla 18\tOK\tProgramación de sigs Output - base")
            except Exception as e:
                print("Tabla 18\tERROR\tProgramación de sigs Output - base")
                LOGGER.warning("Error Tabla 18")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(16,0).checkState()
        if checkObject: #NOTE: Results ready
            try:
                resultsPaths = generate_results(self.path_subarea)
                tabla_tip_veh_nodo_var = resultsPaths["Tipico"]["Vehicular"]["Nodo"]
                tabla_tip_veh_red_var = resultsPaths["Tipico"]["Vehicular"]["Red"]

                tabla_atip_veh_nodo_var = resultsPaths["Atipico"]["Vehicular"]["Nodo"]
                tabla_atip_veh_red_var = resultsPaths["Atipico"]["Vehicular"]["Red"]

                tabla_tip_pea_red_var = resultsPaths["Tipico"]["Peatonal"]["Red"]
                tabla_atip_pea_red_var = resultsPaths["Atipico"]["Peatonal"]["Red"]

                tabla_tip_veh_nodo = doc.new_subdoc(tabla_tip_veh_nodo_var)
                tabla_tip_veh_red = doc.new_subdoc(tabla_tip_veh_red_var)

                tabla_atip_veh_nodo = doc.new_subdoc(tabla_atip_veh_nodo_var)
                tabla_atip_veh_red = doc.new_subdoc(tabla_atip_veh_red_var)

                tabla_tip_pea_red = doc.new_subdoc(tabla_tip_pea_red_var)
                tabla_atip_pea_red = doc.new_subdoc(tabla_atip_pea_red_var)

                VARIABLES.update({
                    "tabla_tip_veh_nodo": tabla_tip_veh_nodo,
                    "tabla_tip_veh_red": tabla_tip_veh_red,
                    "tabla_atip_veh_nodo": tabla_atip_veh_nodo,
                    "tabla_atip_veh_red": tabla_atip_veh_red,
                    "tabla_tip_pea_red": tabla_tip_pea_red,
                    "tabla_atip_pea_red": tabla_atip_pea_red,
                    })
                #SEND_MESSAGE = True
                print("Tabla 20\tOK\tTablas de resultados peatonales, vehiculares y de nodos")
            except Exception as e:
                print("Tabla 20\tERROR\tTablas de resultados peatonales, vehiculares y de nodos")
                LOGGER.warning("Error Tabla 20")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(17,0).checkState()

        if checkObject: #NOTE: Ready tabla23
            try: 
                table23_path = create_table23(self.path_subarea)
                table23 = doc.new_subdoc(table23_path)
                VARIABLES.update({"tabla23": table23})
                print("Tabla 23\tOK\tTabla resumen de resultados")
            except Exception as e:
                print("Tabla 23\tERROR\tTabla resumen de resultados")
                LOGGER.warning("Error Tabla 23")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(18,0).checkState()
        if checkObject: #NOTE: Ready get sigs actual
            try:
                sigActual_path = get_sigs_actual(self.path_subarea, "Actual")
                sigActual = doc.new_subdoc(sigActual_path)
                VARIABLES.update({"sigactual": sigActual})
                print("Sigs actual\tOK")
            except IndexError as e:
                print("Sigs actual\tERROR\tNo hay sigs")
            except Exception as e:
                print("Sigs actual\tERROR")
                LOGGER.warning("Sigs actual")
                LOGGER.warning(str(e))

        checkObject = self.ui.tableWidget.item(19,0).checkState()
        if checkObject: #TODO: ready get sigs propuesto
            try:
                sigPropuesto_path = get_sigs_actual(self.path_subarea, "Output_Base")
                sigPropuesto = doc.new_subdoc(sigPropuesto_path)
                VARIABLES.update({"sigpropuesto": sigPropuesto})
                print("Sigs propuesto\tOK")
            except IndexError as e:
                print("Sigs propuesto\tERROR\tNo hay sigs")
            except Exception as e:
                print("Sigs propuesto\tERROR")
                LOGGER.warning("Sigs propuesto")
                LOGGER.warning(str(e))

        VARIABLES.update({
            "month": month,
        })

        doc.render(VARIABLES)

        #Getting name of subarea
        subareaName = os.path.split(self.path_subarea)[1]

        informePath = Path(self.path_subarea) / f"Informe de transito {subareaName}.docx"
        doc.save(informePath)

        print("\n****STATE: Report created sucessfully****")
        return self.ui.stateLabel.setText("STATE: Report created successfully!")

def main():
    app = QApplication([])
    window = MyWindow()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()