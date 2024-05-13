from docxtpl import DocxTemplate
from call_functions import *
from tables.table1 import create_table1
from tables.table2n3 import create_table2n3
from tables.table4n5 import create_table4n5
from tables.table6 import create_table6
from tables.table7 import create_table7
from tables.table8n9 import create_table8, create_table9
from tables.table12 import create_table12

#Interface
from ui.interface import Ui_Form
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QErrorMessage, QMessageBox

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

    def start(self):
        doc = DocxTemplate("templates/template.docx")

        #Location
        VARIABLES = location(self.path_subarea)

        #Table paths:
        table1_path = create_table1(self.path_subarea)
        print("Tabla 1\t\tOK")
        table2_path, table3_path, dcontet, dcontea = create_table2n3(self.path_subarea)
        print("Tabla 2\t\tOK")
        print("Tabla 3\t\tOK")
        table4_path, table5_path = create_table4n5(self.path_subarea)
        print("Tabla 4\t\tOK")
        print("Tabla 5\t\tOK")
        table6_path = create_table6(self.path_subarea)
        print("Tabla 6\t\tOK")
        table7_path = create_table7(self.path_subarea)
        print("Tabla 7\t\tOK")
        table8_path = create_table8(self.path_subarea)
        print("Tabla 8\t\tOK")
        table9_path = create_table9(self.path_subarea)
        print("Tabla 9\t\tOK")
        table12_path = create_table12(self.path_subarea)
        print("Tabla 12\tOK")

        #Image paths:
        histograma_path = histogramas(self.path_subarea)
        print("Histogramas\tOK")
        flujograma_vehicular_path = flujogramas_vehiculares(self.path_subarea)
        print("Flujogramas vehbiculares\tOK")
        flujograma_peatonal_path = flujogramas_peatonales(self.path_subarea)
        print("Flujogramas peatonales\t\tOK")

        #Table new_subdoc
        table1 = doc.new_subdoc(table1_path)
        table2 = doc.new_subdoc(table2_path)
        table3 = doc.new_subdoc(table3_path)
        table4 = doc.new_subdoc(table4_path)
        table5 = doc.new_subdoc(table5_path)
        table6 = doc.new_subdoc(table6_path)
        table7 = doc.new_subdoc(table7_path)
        table8 = doc.new_subdoc(table8_path)
        table9 = doc.new_subdoc(table9_path)
        histograma = doc.new_subdoc(histograma_path)
        flujograma_vehicular = doc.new_subdoc(flujograma_vehicular_path)
        table12 = doc.new_subdoc(table12_path)
        flujograma_peatonal = doc.new_subdoc(flujograma_peatonal_path)

        VARIABLES.update(
            {
                "tabla1": table1,
                "tabla2": table2,
                "tabla3": table3,
                "tabla4": table4,
                "tabla5": table5,
                "tabla6": table6,
                "tabla7": table7,
                "tabla8": table8,
                "tabla9": table9,
                "histogramas": histograma,
                "dcontet": dcontet,
                "dcontea": dcontea,
                "flujogvmt": flujograma_vehicular,
                "tabla12": table12,
                "flujogpmt": flujograma_peatonal,
            }
        )

        doc.render(VARIABLES)
        informePath = Path(self.path_subarea) / "INFORME.docx"
        doc.save(informePath)

        return self.ui.stateLabel.setText("STATE: Report created successfully!")

def main():
    app = QApplication([])
    window = MyWindow()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()