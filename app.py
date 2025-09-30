
import platform
import streamlit as st
from io import BytesIO
from copy import deepcopy
from typing import Dict, List, Tuple
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
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

WD_PAGE_BREAK = 7
WD_STORY = 6
WD_HEADER_FOOTER_PRIMARY = 1
WD_ALIGN_PARAGRAPH_CENTER = 1


def word_automation_available() -> bool:
    if platform.system().lower() != "windows":
        return False
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return False
    return True


def main() -> None:
    st.set_page_config(page_title="BookMaker", page_icon="BM")
    st.title("BookMaker")
    st.write(
        "Upload DOCX files by section to build a single, organized briefing book. "
        "Each uploader maps to a section in the generated table of contents."
    )
    st.info(
        "Current prototype supports DOCX files. PDFs would need converting to DOCX before combining.",
        icon="i",
    )

    word_available = word_automation_available()
    method_label = "Document assembly method"
    if word_available:
        method = st.radio(
            method_label,
            (
                "Preserve layout (Microsoft Word)",
                "Standard merge (python-docx)",
            ),
            help=(
                "Word automation keeps original pagination, images, and layout intact. "
                "python-docx does a structural merge and is cross-platform."
            ),
        )
    else:
        method = "Standard merge (python-docx)"
        if platform.system().lower() == "windows":
            st.info(
                "Install 'pywin32' locally to enable the Word-based merge.",
                icon="i",
            )
        else:
            st.warning(
                "Microsoft Word automation is only available on Windows. Falling back to python-docx merge.",
                icon="!",
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

        try:
            if method.startswith("Preserve"):
                buffer = build_combined_document_with_word(filtered_sections)
            else:
                buffer = build_combined_document(filtered_sections)
        except RuntimeError as exc:
            st.error(str(exc))
            return

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

    add_table_of_contents(combined, section_style_name)

    for index, (section_name, files) in enumerate(filtered_sections):
        heading = combined.add_paragraph(section_name, style=section_style_name)
        heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        heading.paragraph_format.page_break_before = True

        for uploaded_file in files:
            file_bytes = uploaded_file.getvalue()
            source_doc = Document(BytesIO(file_bytes))
            append_document_body(combined, source_doc)

    apply_footer_with_page_numbers(combined)

    output_buffer = BytesIO()
    combined.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer


def build_combined_document_with_word(filtered_sections: List[Tuple[str, List]]) -> BytesIO:
    if platform.system().lower() != "windows":
        raise RuntimeError("Microsoft Word automation is only supported on Windows.")

    try:
        import win32com.client as win32  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required for Word-based composition. Install it with 'pip install pywin32'."
        ) from exc

    with tempfile.TemporaryDirectory() as tmpdir_str:
        tmpdir = Path(tmpdir_str)
        section_paths: List[Tuple[str, List[Path]]] = []
        for section_name, files in filtered_sections:
            file_paths: List[Path] = []
            for index, uploaded_file in enumerate(files):
                suffix = Path(getattr(uploaded_file, "name", "") or "document.docx").suffix or ".docx"
                safe_name = _safe_stem(section_name)
                file_path = tmpdir / f"{safe_name}_{index}{suffix}"
                file_path.write_bytes(uploaded_file.getvalue())
                file_paths.append(file_path)
            section_paths.append((section_name, file_paths))

        output_path = tmpdir / "combined.docx"
        compose_sections_with_word(section_paths, output_path)
        return BytesIO(output_path.read_bytes())


def compose_sections_with_word(section_paths: List[Tuple[str, List[Path]]], output_path: Path) -> None:
    import win32com.client as win32  # type: ignore

    word = win32.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Add()
        selection = word.Selection
        insert_table_of_contents_word(doc, selection)

        for index, (section_name, files) in enumerate(section_paths):
            if index > 0:
                selection.InsertBreak(WD_PAGE_BREAK)
            selection.Style = "Heading 1"
            selection.TypeText(section_name)
            selection.TypeParagraph()
            for file_path in files:
                selection.InsertFile(str(file_path))
                selection.EndKey(Unit=WD_STORY)

        if doc.TablesOfContents.Count:
            doc.TablesOfContents(1).Update()
        apply_footer_with_page_numbers_word(doc)
        doc.SaveAs(str(output_path))
    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)
        word.Quit()


def insert_table_of_contents_word(doc, selection) -> None:
    selection.Style = "Title"
    selection.TypeText("Table of Contents")
    selection.TypeParagraph()
    doc.TablesOfContents.Add(
        Range=selection.Range,
        UseHeadingStyles=True,
        UpperHeadingLevel=1,
        LowerHeadingLevel=3,
        UseHyperlinks=True,
        IncludePageNumbers=True,
        RightAlignPageNumbers=True,
    )
    selection.EndKey(Unit=WD_STORY)
    selection.TypeParagraph()
    selection.InsertBreak(WD_PAGE_BREAK)
    selection.EndKey(Unit=WD_STORY)


def apply_footer_with_page_numbers_word(doc) -> None:
    for section in doc.Sections:
        footer = section.Footers(WD_HEADER_FOOTER_PRIMARY)
        footer.Range.Text = ""
        footer.PageNumbers.RestartNumberingAtSection = False
        footer.PageNumbers.Add(PageNumberAlignment=WD_ALIGN_PARAGRAPH_CENTER)
        footer.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER


def _safe_stem(value: str) -> str:
    stem = ''.join(ch if ch.isalnum() else '_' for ch in value).strip('_')
    return stem or "section"


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


def add_table_of_contents(document: Document, section_style_name: str) -> None:
    document.add_paragraph("Table of Contents", style="Title")
    toc_paragraph = document.add_paragraph()
    create_field_run(
        toc_paragraph,
        f'TOC \\h \\z \\t "{section_style_name},1"',
        "Update this table in Word to populate the entries.",
    )


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
    for element in list(source.element.body):
        if element.tag == qn("w:sectPr"):
            continue
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
