from docx import Document
from docx.shared import Inches

document = Document()

document.add_heading('Document YUDI', 0)

# document.add_section('')

par = document.add_paragraph('A plain text having')

document.add_picture('imgs\IMG_20180720_143234.jpg', width=Inches(3))

document.save('demo.docx')
