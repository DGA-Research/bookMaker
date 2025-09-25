import streamlit as st
from io import BytesIO
from copy import deepcopy
from itertools import count
from typing import Dict, List, Tuple

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_TAB_ALIGNMENT,
    WD_TAB_LEADER,
)
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches

SECTION_ORDER = [
    "Top Hits",
    "Methodology",
    "Biographical",
    "Family/Personal Info",
    "Buisness Interests",
    "Race Review",
    "Campaign Finance",
    "Issues",
    "Appendicies",
    "Questionaires",
    "Scorecards",
    "Travel Discosureles",
    "Offical Office Disbursments",
]

BOOKMARK_COUNTER = count(1)


def main() -> None:
    st.set_page_config(page_title="BookMaker", page_icon="📚")
    st.title("BookMaker")
    st.write(
        "Upload DOCX files by section to build a single, organized briefing book. "
        "Each uploader maps to a section in the generated table of contents."
    )
    st.info(
        "Current prototype supports DOCX files. PDFs would need converting to DOCX before combining.",
        icon="ℹ️",
    )

    uploads: Dict[str, List] = {}
    for section in SECTION_ORDER:
        st.subheader(section)
        uploads[section] = st.file_uploader(
            f"Upload DOCX files for {section}",
            type=["docx"],
            accept_multiple_files=True,
            key=section,
            help="Files are appended in the order they appear below.",
        )

    if st.button("Generate combined document"):
        filtered_sections: List[Tuple[str, List]] = [
            (section, files)
            for section, files in uploads.items()
            if files
        ]

        if not filtered_sections:
            st.warning("Upload at least one DOCX file before generating the book.")
            return

        buffer = build_combined_document(filtered_sections)
        st.success(
            "Combined document ready. Open in Word and update fields to refresh the TOC and page numbers."
        )
        st.download_button(
            label="Download combined DOCX",
            data=buffer.getvalue(),
            file_name="combined_book.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


def build_combined_document(filtered_sections: List[Tuple[str, List]]) -> BytesIO:
    combined = Document()
    remove_initial_paragraph_if_empty(combined)
    section_style_name = ensure_section_style(combined)

    sections_with_meta = [
        {
            "name": section_name,
            "files": files,
            "bookmark": bookmark_name_from_section(section_name, position),
        }
        for position, (section_name, files) in enumerate(filtered_sections, start=1)
    ]

    add_table_of_contents(combined, sections_with_meta)
    combined.add_page_break()

    for index, section in enumerate(sections_with_meta):
        if index > 0:
            combined.add_page_break()

        heading_paragraph = combined.add_paragraph(section["name"], style=section_style_name)
        add_bookmark_to_paragraph(heading_paragraph, section["bookmark"])

        for uploaded_file in section["files"]:
            file_bytes = uploaded_file.getvalue()
            source_doc = Document(BytesIO(file_bytes))
            append_document_body(combined, source_doc)

    apply_footer_with_page_numbers(combined)

    output_buffer = BytesIO()
    combined.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer


def ensure_section_style(document: Document) -> str:
    style_name = "BookMaker Section"
    try:
        document.styles[style_name]
        return style_name
    except KeyError:
        pass

    section_style = document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
    section_style.base_style = document.styles["Heading 1"]
    section_style.quick_style = True
    return style_name


def bookmark_name_from_section(section_name: str, position: int) -> str:
    sanitized = "".join(char for char in section_name if char.isalnum())
    if not sanitized:
        sanitized = f"Section{position}"
    return f"bm_{position}_{sanitized[:32]}"


def add_table_of_contents(document: Document, sections_with_meta) -> None:
    document.add_paragraph("Table of Contents", style="Title")

    if not sections_with_meta:
        document.add_paragraph("No sections supplied.")
        return

    for section in sections_with_meta:
        add_toc_entry(document, section["name"], section["bookmark"])

    try:
        document.add_paragraph("Update fields in Word to refresh page links.", style="Caption")
    except KeyError:
        document.add_paragraph("Update fields in Word to refresh page links.")


def add_toc_entry(document: Document, label: str, bookmark_name: str) -> None:
    paragraph = document.add_paragraph()
    try:
        paragraph.style = document.styles["TOC 1"]
    except KeyError:
        pass

    paragraph.paragraph_format.tab_stops.add_tab_stop(
        Inches(6), alignment=WD_TAB_ALIGNMENT.RIGHT, leader=WD_TAB_LEADER.DOTS
    )

    add_internal_hyperlink(paragraph, label, bookmark_name)
    paragraph.add_run("\t")
    create_field_run(paragraph, f"PAGEREF {bookmark_name} \\h", "1")


def add_internal_hyperlink(paragraph, text: str, bookmark_name: str) -> None:
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), bookmark_name)

    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    r_style = OxmlElement("w:rStyle")
    r_style.set(qn("w:val"), "Hyperlink")
    r_pr.append(r_style)
    new_run.append(r_pr)

    text_element = OxmlElement("w:t")
    text_element.text = text
    new_run.append(text_element)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_bookmark_to_paragraph(paragraph, bookmark_name: str) -> None:
    bookmark_id = next(BOOKMARK_COUNTER)
    paragraph_text = paragraph.text

    for run in list(paragraph.runs):
        paragraph._p.remove(run._r)

    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), bookmark_name)
    paragraph._p.append(start)

    paragraph.add_run(paragraph_text)

    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.append(end)




def apply_footer_with_page_numbers(document: Document) -> None:
    for section in document.sections:
        section.footer.is_linked_to_previous = False
        if section.footer.paragraphs:
            footer_paragraph = section.footer.paragraphs[0]
            clear_paragraph(footer_paragraph)
        else:
            footer_paragraph = section.footer.add_paragraph()

        footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer_paragraph.add_run("Page ")
        create_field_run(footer_paragraph, "PAGE", "1")
        footer_paragraph.add_run(" of ")
        create_field_run(footer_paragraph, "NUMPAGES", "1")


def append_document_body(target: Document, source: Document) -> None:
    for element in source.element.body:
        target.element.body.append(deepcopy(element))


    run = paragraph.add_run()

    field_begin = OxmlElement("w:fldChar")
    field_begin.set(qn("w:fldCharType"), "begin")
    run._r.append(field_begin)

    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = field_code
    run._r.append(instruction)

    field_separate = OxmlElement("w:fldChar")
    field_separate.set(qn("w:fldCharType"), "separate")
    run._r.append(field_separate)

    text = OxmlElement("w:t")
    text.text = default_text
    run._r.append(text)

    field_end = OxmlElement("w:fldChar")
    field_end.set(qn("w:fldCharType"), "end")
    run._r.append(field_end)

    return run


def clear_paragraph(paragraph) -> None:
    for child in list(paragraph._p):
        paragraph._p.remove(child)


def remove_initial_paragraph_if_empty(document: Document) -> None:
    if document.paragraphs and not document.paragraphs[0].text:
        p = document.paragraphs[0]._p
        p.getparent().remove(p)


if __name__ == "__main__":
    main()
