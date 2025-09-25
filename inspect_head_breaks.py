from docx import Document

doc = Document('combined_book (8).docx')
for idx, para in enumerate(doc.paragraphs[:10]):
    brs = [child for child in para._p.iterchildren() if child.tag.endswith('br')]
    print(idx, repr(para.text), [b.attrib for b in brs])
