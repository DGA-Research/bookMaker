from docx import Document

doc = Document('combined_book (8).docx')
for idx, para in enumerate(doc.paragraphs[:30]):
    print(idx, repr(para.text))
