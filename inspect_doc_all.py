from docx import Document

doc = Document('combined_book (6).docx')
for idx, para in enumerate(doc.paragraphs):
    brs = [child for child in para._p.iterchildren() if child.tag.endswith('br')]
    print(idx, repr(para.text), [b.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') for b in brs])
