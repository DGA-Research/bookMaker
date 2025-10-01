"""Compose the briefing book using Microsoft Word automation to preserve layout."""

from pathlib import Path

from docx import Document
from typing import List, Optional

SECTION_STYLE_NAME = "BookMaker Section"
TOP_HIT_STYLE_NAME = "BookMaker Top Hit"
WD_STYLE_TYPE_PARAGRAPH = 1

TEMPLATE_PATH = Path(__file__).resolve().parent / "testdocument.docx"

# Word constant fallbacks (values from the VBA enum)
WD_PAGE_BREAK = 7
WD_STORY = 6
WD_HEADER_FOOTER_PRIMARY = 1
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_VERTICAL_TOP = 0
WD_ALIGN_VERTICAL_CENTER = 1


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
        folder_path = parts_dir / section
        doc_paths: List[Path] = []
        if folder_path.exists() and folder_path.is_dir():
            doc_paths.extend(
                sorted(
                    (p for p in folder_path.glob("*.docx") if p.is_file()),
                    key=lambda p: p.name.lower(),
                )
            )
        filename = mapping.get(section)
        if not doc_paths and filename:
            candidate = parts_dir / filename
            if candidate.exists():
                doc_paths.append(candidate)
        if doc_paths:
            yield section, doc_paths
        else:
            missing_path = (parts_dir / filename) if filename else folder_path
            print(f"Warning: missing DOCX for section '{section}': {missing_path}")


def ensure_word_paragraph_style(doc, style_name: str, base_style_name: str) -> str:
    try:
        doc.Styles(style_name)
        return style_name
    except Exception:
        pass

    try:
        new_style = doc.Styles.Add(style_name, WD_STYLE_TYPE_PARAGRAPH)
    except Exception:
        return base_style_name

    try:
        new_style.BaseStyle = doc.Styles(base_style_name)
    except Exception:
        pass
    try:
        new_style.QuickStyle = True
    except Exception:
        pass
    return style_name



def insert_table_of_contents(doc, selection, section_style_name: str, top_hit_style_name: str):
    selection.Style = "Title"
    selection.TypeText("Table of Contents")
    selection.TypeParagraph()
    toc = doc.TablesOfContents.Add(
        Range=selection.Range,
        UseHeadingStyles=False,
        UpperHeadingLevel=1,
        LowerHeadingLevel=1,
        UseHyperlinks=True,
        IncludePageNumbers=True,
        RightAlignPageNumbers=True,
    )

    try:
        heading_styles = toc.HeadingStyles
        section_style = doc.Styles(section_style_name)
        heading_styles.Item(1).Style = section_style
        heading_styles.Item(1).Level = 1
    except Exception:
        pass

    try:
        for index in range(2, heading_styles.Count + 1):
            heading_styles.Item(index).Level = 9
    except Exception:
        pass

    selection.EndKey(Unit=WD_STORY)
    selection.TypeParagraph()
    selection.InsertBreak(WD_PAGE_BREAK)
    selection.EndKey(Unit=WD_STORY)
    return toc


def set_narrow_margins(doc, word, inches=0.5):
    margin = inches * 72  # convert inches to points
    page_setup = doc.PageSetup
    page_setup.TopMargin = margin
    page_setup.BottomMargin = margin
    page_setup.LeftMargin = margin
    page_setup.RightMargin = margin



def apply_template_heading_style(doc, start, end, style_name: str, match_text: Optional[str]) -> None:
    try:
        target_style = doc.Styles(style_name)
    except Exception:
        target_style = None

    target_range = doc.Range(Start=start, End=end)
    for paragraph in target_range.Paragraphs:
        text = paragraph.Range.Text.replace("\r", "").replace("\x07", "").strip()
        if not text:
            continue
        if match_text is not None and text.casefold() != match_text.casefold():
            continue
        try:
            if target_style is not None:
                paragraph.Style = target_style
            else:
                paragraph.Style = style_name
        except Exception:
            try:
                paragraph.Style = style_name
            except Exception:
                pass
        break


def remove_duplicate_heading_paragraph(doc, start, end, match_text: str) -> None:
    if not match_text:
        return
    match_key = match_text.casefold()
    target_range = doc.Range(Start=start, End=end)
    for paragraph in target_range.Paragraphs:
        text = paragraph.Range.Text.replace("\r", "").replace("\x07", "").strip()
        if not text:
            continue
        if text.casefold() == match_key:
            paragraph.Range.Delete()
        break


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
        template_to_use = str(TEMPLATE_PATH) if TEMPLATE_PATH.exists() else None
        if template_to_use:
            doc = word.Documents.Add(Template=template_to_use)
            doc.Content.Delete()
        else:
            doc = word.Documents.Add()
        section_style_name = ensure_word_paragraph_style(doc, SECTION_STYLE_NAME, "Heading 1")
        top_hit_style_name = ensure_word_paragraph_style(doc, TOP_HIT_STYLE_NAME, "Heading 2")
        set_narrow_margins(doc, word)
        selection = word.Selection
        toc = insert_table_of_contents(doc, selection, section_style_name, top_hit_style_name)

        for index, (section_name, file_paths) in enumerate(input_sections):
            if index > 0:
                selection.InsertBreak(WD_PAGE_BREAK)

            file_paths = list(file_paths)
            if not file_paths:
                continue

            heading_style_name = section_style_name
            selection.Style = heading_style_name
            selection.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
            try:
                selection.ParagraphFormat.SpaceBefore = word.InchesToPoints(4)
                selection.ParagraphFormat.SpaceAfter = word.InchesToPoints(4)
            except Exception:
                pass
            selection.TypeText(section_name)
            selection.TypeParagraph()
            selection.InsertBreak(WD_PAGE_BREAK)
            selection.EndKey(Unit=WD_STORY)
            try:
                selection.ParagraphFormat.SpaceBefore = 0
                selection.ParagraphFormat.SpaceAfter = 0
            except Exception:
                pass
            selection.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_LEFT

            content_start = selection.Start

            for sub_index, file_path in enumerate(file_paths):
                if section_name == "Top Hits" and sub_index > 0:
                    selection.InsertBreak(WD_PAGE_BREAK)
                    selection.EndKey(Unit=WD_STORY)

                file_insert_start = selection.Start
                selection.InsertFile(str(file_path))
                selection.EndKey(Unit=WD_STORY)
                file_insert_end = selection.Start

                if section_name == "Top Hits":
                    apply_template_heading_style(doc, file_insert_start, file_insert_end, top_hit_style_name, None)

            inserted_end = selection.Start

            if section_name != "Top Hits":
                remove_duplicate_heading_paragraph(doc, content_start, inserted_end, section_name)

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


















