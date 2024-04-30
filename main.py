from docxtpl import DocxTemplate
from recollector import *


def fill_docx(path):
    doc = DocxTemplate("templates/main.docx")
    
    #Location
    dict_info = location(path)

    #counts
    RESULTS = counts(path)
    for topic, dictionario in RESULTS.items():
        print(topic)
        print(dictionario)

    #doc.render(dict_info)
    #doc.save("TEST.docx")

if __name__ == '__main__':
    PATH = r"/home/chiky/Projects/REPORT/data/1. Proyecto Surco (Sub. 16 -59)/6. Sub Area Vissim/Sub Area 016"
    fill_docx(PATH)