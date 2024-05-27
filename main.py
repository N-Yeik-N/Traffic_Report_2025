from docxtpl import DocxTemplate
from tables.table1 import create_table1
from tables.table2n3 import create_table2n3
from tables.table4n5 import create_table4n5
from tables.table6 import create_table6
from tables.table7 import create_table7
from tables.table8n9 import create_table8, create_table9
from tables.table12 import create_table12
from tables.table10 import create_table10
from tables.table11 import create_table11
from tables.table14 import create_table14
from tables.table17 import create_table17
from tables.table18 import create_table18
from tables.table19 import create_table19
from parrafos.paragraphs import cambios_variable
from src.call_functions import *
from sigs.sig_actual import get_sigs_actual

import logging


#Interface
from ui.interface import Ui_Form
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QErrorMessage, QMessageBox

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
    
    def open_file(self):
        self.path_subarea = QFileDialog.getExistingDirectory(self, 'Open File')
        if self.path_subarea:
            self.ui.pathLineEdit.setText(self.path_subarea)
            #Logger location
            directory, _ = os.path.split(self.path_subarea)
            log_path = os.path.join(directory, "logs_report")

            if not os.path.exists(log_path):
                os.mkdir(log_path)

            fh = logging.FileHandler(os.path.join(log_path, "logger.log"), mode='w')
            fh.setFormatter(f)
            LOGGER.addHandler(fh)

    def start(self):
        doc = DocxTemplate("templates/template.docx")

        #Location
        VARIABLES, codintersecciones = location(self.path_subarea)

        #Table paths:
        try:
            table1_path = create_table1(self.path_subarea)
            table1 = doc.new_subdoc(table1_path)
            VARIABLES.update({"tabla1": table1})
            print("Tabla 1\t\tOK")
        except Exception as e:
            print("Warning: can't create table 1")
            LOGGER.warning(str(e))

        try:
            table2_path, table3_path, dcontet, dcontea = create_table2n3(self.path_subarea)
            table2 = doc.new_subdoc(table2_path)
            table3 = doc.new_subdoc(table3_path)
            VARIABLES.update({"tabla2": table2, "tabla3": table3, "dcontet": dcontet, "dcontea": dcontea})
            print("Tabla 2\t\tOK")
            print("Tabla 3\t\tOK")
        except Exception as e:
            print("Warning: can't create table 2 or table 3")
            LOGGER.warning(str(e))

        try:
            table4_path, table5_path = create_table4n5(self.path_subarea)
            table4 = doc.new_subdoc(table4_path)
            table5 = doc.new_subdoc(table5_path)
            VARIABLES.update({"tabla4": table4, "tabla5": table5})
            print("Tabla 4\t\tOK")
            print("Tabla 5\t\tOK")
        except Exception as e:
            print("Warning: can't create table 4 or table 5")
            LOGGER.warning(str(e))

        try:
            table6_path = create_table6(self.path_subarea)
            table6 = doc.new_subdoc(table6_path)
            VARIABLES.update({"tabla6": table6})
            print("Tabla 6\t\tOK")
        except Exception as e:
            print("Warning: can't create table 6")
            LOGGER.warning(str(e))

        try:
            table7_path = create_table7(self.path_subarea)
            table7 = doc.new_subdoc(table7_path)
            VARIABLES.update({"tabla7": table7})
            print("Tabla 7\t\tOK")
        except Exception as e:
            print("Warning: can't create table 7")
            LOGGER.warning(str(e))

        try:
            table8_path = create_table8(self.path_subarea)
            table8 = doc.new_subdoc(table8_path)
            VARIABLES.update({"tabla8": table8})
            print("Tabla 8\t\tOK")
        except Exception as e:
            print("Warning: can't create table 8")
            LOGGER.warning(str(e))

        try:    
            table9_path = create_table9(self.path_subarea)
            table9 = doc.new_subdoc(table9_path)
            VARIABLES.update({"tabla9": table9})
            print("Tabla 9\t\tOK")
        except Exception as e:
            print("Warning: can't create table 9")
            LOGGER.warning(str(e))

        try:
            table10_path = create_table10(self.path_subarea)
            tabla10 = doc.new_subdoc(table10_path)
            VARIABLES.update({"tabla10": tabla10})
            print("Tabla 10\tOK")
        except Exception as e:
            print("Warning: can't create table 10")
            LOGGER.warning(str(e))

        try:
            table11_path = create_table11(self.path_subarea)
            tabla11 = doc.new_subdoc(table11_path)
            VARIABLES.update({"tabla11": tabla11})
            print("Tabla 11\tOK")
        except Exception as e:
            print("Warning: can't create table 11")
            LOGGER.warning(str(e))

        try:
            table12_path = create_table12(self.path_subarea)
            table12 = doc.new_subdoc(table12_path)
            VARIABLES.update({"tabla12": table12})
            print("Tabla 12\tOK")
        except Exception as e:
            print("Warning: can't create table 12")
            LOGGER.warning(str(e))

        try:
            table14_path, VARIABLES_OD = create_table14(self.path_subarea)
            table14 = doc.new_subdoc(table14_path)
            VARIABLES.update({"tabla14": table14})
            VARIABLES.update(VARIABLES_OD)
            print("Tabla 14\tOK")
        except Exception as e:
            print("Warning: can't create table 14")
            LOGGER.warning(str(e))

        try:
            table17_path = create_table17(self.path_subarea)
            table17 = doc.new_subdoc(table17_path)
            VARIABLES.update({"tabla17": table17})
            print("Tabla 17\tOK")
        except Exception as e:
            print("Warning: can't create table 17")
            LOGGER.warning(str(e))

        try:
            table18_path = create_table18(self.path_subarea)
            table18 = doc.new_subdoc(table18_path)
            VARIABLES.update({"tabla18": table18})
            print("Tabla 18\tOK")
        except Exception as e:
            print("Warning: can't create table 18")
            LOGGER.warning(str(e))

        try:
            table19_path = create_table19(self.path_subarea)
            table19 = doc.new_subdoc(table19_path)
            VARIABLES.update({"tabla19": table19})
            print("Tabla 19\tOK")
        except Exception as e:
            print("Warning: can't create table 19")
            LOGGER.warning(str(e))

        #Paragraphs:
        try:
            cambios_path = cambios_variable(self.path_subarea, codintersecciones)
            cambioParagraph = doc.new_subdoc(cambios_path)
            VARIABLES.update({"cambios": cambioParagraph})
            print("Párrafos\tOK")
        except Exception as e:
            print("Warning: can't create paragraphs (variable)")
            LOGGER.warning(str(e))

        #Image paths:
        try:
            histograma_path = histogramas(self.path_subarea)
            histograma = doc.new_subdoc(histograma_path)
            VARIABLES.update({"histogramas": histograma})
            print("Histogramas\tOK")
        except Exception as e:
            print("Warning: can't create histogramas")
            LOGGER.warning(str(e))

        try:
            flujograma_vehicular_path = flujogramas_vehiculares(self.path_subarea)
            flujograma_vehicular = doc.new_subdoc(flujograma_vehicular_path)
            VARIABLES.update({"flujogvmt": flujograma_vehicular})
            print("Flujogramas vehbiculares\tOK")
        except Exception as e:
            print("Warning: can't create flujogramas vehiculares")
            LOGGER.warning(str(e))

        try:
            flujograma_peatonal_path = flujogramas_peatonales(self.path_subarea)
            flujograma_peatonal = doc.new_subdoc(flujograma_peatonal_path)
            VARIABLES.update({"flujogpmt": flujograma_peatonal})
            print("Flujogramas peatonales\t\tOK")
        except Exception as e:
            print("Warning: can't create flujogramas peatonales")
            LOGGER.warning(str(e))

        try:
            sigActual_path = get_sigs_actual(self.path_subarea)
            sigActual = doc.new_subdoc(sigActual_path)
            VARIABLES.update({"sigactual": sigActual})
            print("Sigs actual\t\tOK")
        except Exception as e:
            print("Warning: can't create sigs actual")
            LOGGER.warning(str(e))

        doc.render(VARIABLES)
        informePath = Path(self.path_subarea) / "INFORME.docx"
        doc.save(informePath)

        print("STATE: Report created sucessfully.")
        return self.ui.stateLabel.setText("STATE: Report created successfully!")

def main():
    app = QApplication([])
    window = MyWindow()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()