# BookMaker CLI

BookMaker assembles a single briefing book by stitching together per-section DOCX files in a fixed order. The CLI can merge files entirely in Python or, when available, automate Microsoft Word to preserve complex layouts.

## Prerequisites

- Python 3.9 or newer
- pip
- Optional: Microsoft Word (needed for the `word` merge method)
- Optional: Windows `pywin32` dependency for Word automation (already listed in `requirements.txt`)

## Quick Start

1. Clone or download this repository.
2. Open a terminal in the project folder.
3. Create a virtual environment: `python -m venv .venv`
4. Activate it (PowerShell): `.\.venv\Scripts\Activate.ps1`
5. Install dependencies: `pip install -r requirements.txt`

## Prepare Section Files

Place the DOCX parts in the `bookParts/` directory. The tool looks for the following file names in this order (missing files are skipped with a warning):

| Section label               | Expected DOCX name                  |
|-----------------------------|-------------------------------------|
| Top Hits                    | TOP HITS.docx                       |
| Methodology                 | METHODOLOGY.docx                    |
| Biographical                | BIOGRAPHICAL.docx                   |
| Family/Personal Info        | FAMILY PERSONAL INFO.docx           |
| Buisness Interests          | BUISNESS INTERESTS.docx             |
| Race Review                 | RACE REVIEW.docx                    |
| Campaign Finance            | CAMPAIGN FINANCE.docx               |
| Issues                      | ISSUES.docx                         |
| Appendicies                 | APPENDICIES.docx                    |
| Questionaires               | QUESTIONNAIRES.docx                 |
| Scorecards                  | SCORECARD.docx                      |
| Travel Discosureles         | TRAVEL DISCLOSURES.docx             |
| Offical Office Disbursments | OFFICIAL OFFICE DISBURSEMENTS.docx  |

Tip: keep a clean copy of each DOCX part outside `bookParts/` and copy the versions you want to merge into this folder when generating a book.

## Compose a Book

Run the CLI from the project root:

```
python app.py --parts-dir bookParts --output combined_book.docx --method auto
```

Key options:

- `--parts-dir`: Folder containing the DOCX parts (defaults to `bookParts`).
- `--output`: Destination path for the combined DOCX (defaults to `combined_book.docx`).
- `--method`: `auto` (default) tries Word automation first, then falls back to pure Python. Choose `word` to require Word automation or `python-docx` to force the Python-only merge.
- `--quiet`: Suppress progress output (warnings and errors still print).

When the script finishes, open the resulting DOCX in Word and update the Table of Contents (Right-click the TOC > Update Field > Update entire table) so the page numbers reflect the final layout.

## Word Automation Notes

- Word automation requires Windows, Microsoft Word, and the `pywin32` package.
- If Word is not installed or accessible, the CLI automatically uses the Python-based merge unless you explicitly set `--method word`.
- The Word-driven merge preserves headers, footers, and advanced formatting more faithfully than the pure Python approach.

## Troubleshooting

- **Missing section message**: Ensure the DOCX exists and matches the expected file name exactly (including spaces and capitalization).
- **`python-docx` merge issues**: Complex page layouts or macros may not carry over perfectly. Re-run with `--method word` on a machine with Microsoft Word installed for better fidelity.
- **TOC shows placeholder text**: Open the output DOCX in Word and update the Table of Contents to refresh entries and page numbers.
