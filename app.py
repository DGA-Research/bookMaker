import streamlit as st
from io import BytesIO
from copy import deepcopy
from typing import Dict, List, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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


def main() -> None:
    st.set_page_config(page_title="BookMaker", page_icon="book")
    st.title("BookMaker")
    st.write(
        "Upload DOCX files by section to build a single, organized briefing book. "
        "Each uploader maps to a section in the generated table of contents."
    )
    st.info(
        "Current prototype supports DOCX files. PDFs would need converting to DOCX before combining.",
        icon="i",
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
    add_table_of_contents(combined)

    for index, (section_name, files) in enumerate(filtered_sections):
        if index > 0:
            combined.add_page_break()

        combined.add_paragraph(section_name, style="Heading 1")

        for uploaded_file in files:
            file_bytes = uploaded_file.getvalue()
            source_doc = Document(BytesIO(file_bytes))
            append_document_body(combined, source_doc)

    apply_footer_with_page_numbers(combined)

    output_buffer = BytesIO()
    combined.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer


def add_table_of_contents(document: Document) -> None:
    document.add_paragraph("Table of Contents", style="Title")
    toc_paragraph = document.add_paragraph()
    create_field_run(
        toc_paragraph,
        'TOC \\o "1-3" \\h \\z \\u',
        "Right-click and choose Update Field to populate this table.",
    )
    document.add_page_break()


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


def create_field_run(paragraph, field_code: str, default_text: str = ""):
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
