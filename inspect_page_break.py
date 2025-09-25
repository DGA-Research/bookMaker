from docx import Document

path = 'combined_book (6).docx'
doc = Document(path)
for idx, para in enumerate(doc.paragraphs):
    ppr = para._p.pPr
    if ppr is not None and ppr.pageBreakBefore is not None:
        print('paragraph', idx, repr(para.text), 'page_break_before:', ppr.pageBreakBefore.val)
