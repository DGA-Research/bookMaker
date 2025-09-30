"""Compose the briefing book using Microsoft Word automation to preserve layout."""

from pathlib import Path

# Word constant fallbacks (values from the VBA enum)
WD_PAGE_BREAK = 7
WD_STORY = 6
WD_HEADER_FOOTER_PRIMARY = 1
WD_ALIGN_PARAGRAPH_CENTER = 1


def get_section_map():
    # Keep this mapping consistent with app.py's SECTION_ORDER
    return {
        "Top Hits": "TOP HITS.docx",
        "Methodology": "METHODOLOGY.docx",
        "Biographical": "BIOGRAPHICAL.docx",
        "Family/Personal Info": "FAMILY PERSONAL INFO.docx",
        "Buisness Interests": "BUISNESS INTERESTS.docx",
        "Race Review": "RACE REVIEW.docx",
        "Campaign Finance": "CAMPAIGN FINANCE.docx",
        "Issues": "ISSUES.docx",
        "Appendicies": "APPENDICIES.docx",
        "Questionaires": "QUESTIONNAIRES.docx",
        "Scorecards": "SCORECARD.docx",
        "Travel Discosureles": "TRAVEL DISCLOSURES.docx",
        "Offical Office Disbursments": "OFFICIAL OFFICE DISBURSEMENTS.docx",
    }


def section_files(section_order, parts_dir):
    mapping = get_section_map()
    for section in section_order:
        filename = mapping.get(section)
        if not filename:
            continue
        path = parts_dir / filename
        if path.exists():
            yield section, path
        else:
            print(f"Warning: missing DOCX for section '{section}': {path}")


def insert_table_of_contents(doc, selection):
    selection.Style = "Title"
    selection.TypeText("Table of Contents")
    selection.TypeParagraph()
    toc = doc.TablesOfContents.Add(
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
    return toc


def apply_page_numbers(doc):
    for section in doc.Sections:
        footer = section.Footers(WD_HEADER_FOOTER_PRIMARY)
        footer.Range.Text = ""
        footer.PageNumbers.RestartNumberingAtSection = False
        footer.PageNumbers.Add(PageNumberAlignment=WD_ALIGN_PARAGRAPH_CENTER)
        footer.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER


def compose_via_word(input_sections, output_path):
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError("win32com is required for Word-driven composition") from exc

    word = win32.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    toc = None
    try:
        doc = word.Documents.Add()
        selection = word.Selection
        toc = insert_table_of_contents(doc, selection)

        for index, (section_name, file_path) in enumerate(input_sections):
            if index > 0:
                selection.InsertBreak(WD_PAGE_BREAK)
            selection.Style = "Heading 1"
            selection.TypeText(section_name)
            selection.TypeParagraph()
            selection.InsertFile(str(file_path))
            selection.EndKey(Unit=WD_STORY)

        if toc is not None:
            toc.Update()
        apply_page_numbers(doc)
        doc.SaveAs(str(output_path))
    finally:
        if doc is not None:
            doc.Close(SaveChanges=False)
        word.Quit()


def main():
    from app import SECTION_ORDER

    project_root = Path(__file__).resolve().parent
    parts_dir = project_root / "bookParts"
    if not parts_dir.exists():
        raise FileNotFoundError(f"bookParts directory not found: {parts_dir}")

    sections = list(section_files(SECTION_ORDER, parts_dir))
    if not sections:
        raise RuntimeError("No source files found to compose")

    output_path = project_root / "combined_book_word.docx"
    compose_via_word(sections, output_path)
    print(f"Combined document created: {output_path}")


if __name__ == "__main__":
    main()
