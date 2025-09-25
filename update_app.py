from pathlib import Path

path = Path('app.py')
text = path.read_text(encoding='utf-8')
text = text.replace('    add_table_of_contents(combined, section_style_name)\n    combined.add_page_break()\n\n    for index, (section_name, files) in enumerate(filtered_sections):\n        heading = combined.add_paragraph(section_name, style=section_style_name)\n        heading.alignment = WD_ALIGN_PARAGRAPH.LEFT\n        if index > 0:\n            heading.paragraph_format.page_break_before = True\n\n        for uploaded_file in files:\n', '    add_table_of_contents(combined, section_style_name)\n\n    for index, (section_name, files) in enumerate(filtered_sections):\n        heading = combined.add_paragraph(section_name, style=section_style_name)\n        heading.alignment = WD_ALIGN_PARAGRAPH.LEFT\n        heading.paragraph_format.page_break_before = True\n\n        for uploaded_file in files:\n')
path.write_text(text, encoding='utf-8')
