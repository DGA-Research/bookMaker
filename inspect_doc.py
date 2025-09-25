from docx import Document

doc = Document('combined_book (6).docx')
for idx, para in enumerate(doc.paragraphs):
    text = para.text.strip()
    brs = [child for child in para._p.iterchildren() if child.tag.endswith('br')]
    if text or brs:
        attrs = [b.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') for b in brs]
        print(idx, repr(text), attrs)
