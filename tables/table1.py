from docx.shared import Inches
from docxtpl import DocxTemplate

doc = DocxTemplate(r'templates/main.docx')

CONSTANTES = {
    'nameproject': "HOLA",
}

doc.render(CONSTANTES)
doc.save(r"tests/TEST2.docx")