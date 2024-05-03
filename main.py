from docxtpl import DocxTemplate
from call_functions import location

def fill_docx(path):
    doc = DocxTemplate("templates/template.docx")
    
    #Location
    VARIABLES1 = location(path)

    doc.render(VARIABLES1)
    doc.save("TEST.docx")

if __name__ == '__main__':
    PATH = r"/home/chiky/Projects/REPORT/data/1. Proyecto Surco (Sub. 16 -59)/6. Sub Area Vissim/Sub Area 016"
    fill_docx(PATH)