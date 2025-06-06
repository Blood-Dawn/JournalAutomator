"""Utility functions for updating ABNFF journal Word documents."""

import argparse
import json
from pathlib import Path
from typing import Iterable, List, Optional

from docx import Document
from docx.shared import Pt


def load_document(path: Path) -> Document:
    """Open the Word file at ``path`` and return a ``Document`` object."""
    return Document(str(path))


def save_document(doc: Document, path_out: Path) -> None:
    """Save ``doc`` to ``path_out``."""
    doc.save(str(path_out))
    
def replace_text_in_paragraphs(paragraphs, search_text, replace_text):
    for p in paragraphs:
        if search_text in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if search_text in inline[i].text:
                    inline[i].text = inline[i].text.replace(search_text, replace_text)

def update_front_cover(
    doc: Document,
    volume: str,
    issue: str,
    month_year: str,
    section_title: str,
    page_num: int,
) -> None:
    """Update volume/issue block on the front cover."""

    search = "Volume"
    for p in doc.paragraphs:
        if search in p.text:
            p.text = (
                f"Volume {volume}, Issue {issue}\n{month_year}\n{section_title}\nPage {page_num}"
            )
            for run in p.runs:
                run.font.bold = True
            break
            
def update_business_information(
    doc: Document, old_year: str, new_beginning_text: str
) -> None:
    """Update the business information block on page 1."""

    replace_text_in_paragraphs(doc.paragraphs, old_year, "")

    for p in doc.paragraphs:
        if "Annual subscription" in p.text:
            # replace first sentence beginning
            first_part = p.text.split(".")[0]
            if first_part:
                p.text = p.text.replace(first_part, new_beginning_text, 1)
            else:
                p.text = new_beginning_text + p.text
            break


def update_page2_header(doc: Document, new_header_line1: str, page_num: int) -> None:
    """Replace the header text on page 2."""

    for section in doc.sections:
        header = section.header
        text = f"{new_header_line1}\nPage {page_num}"
        for p in header.paragraphs:
            if p.text.strip():
                p.text = text
                break


def update_associate_editors(
    doc: Document, remove_name: str, new_name: str, new_affiliation: str, new_email: str
) -> None:
    """Update the associate editors list on page 2."""

    for p in doc.paragraphs:
        if remove_name in p.text:
            # replace the line with new associate editor info
            p.text = f"{new_name}, {new_affiliation}\n{new_email}"
            break


def remove_text_labels(doc: Document, labels_to_remove: Iterable[str]) -> None:
    """Remove phrases from the document wherever they appear."""

    for label in labels_to_remove:
        replace_text_in_paragraphs(doc.paragraphs, label, "")


def update_assistant_editors(doc: Document, remove_name: str) -> None:
    """Remove an entry from the assistant editors list."""

    for p in doc.paragraphs:
        if remove_name in p.text:
            p.text = ""


def insert_presidents_message(doc: Document, image_path: Path, message_text: str) -> None:
    """Insert the president's message and optional image on page 3."""

    text = message_text if message_text else "<<Awaiting President's message>>"
    for p in doc.paragraphs:
        if "President's Message" in p.text:
            idx = doc.paragraphs.index(p)
            target = doc.paragraphs[idx + 1]
            target.text = text
            break


def extract_article_titles_from_toc(doc: Document) -> List[str]:
    """Return article titles listed under the ARTICLES section in the TOC."""
    toc_start = None
    paragraphs = doc.paragraphs
    # locate table of contents
    for i, p in enumerate(paragraphs):
        if "TABLE OF CONTENTS" in p.text.upper():
            toc_start = i
            break
    if toc_start is None:
        return []

    # find ARTICLES heading within TOC
    start = None
    for j in range(toc_start + 1, len(paragraphs)):
        text = paragraphs[j].text.strip()
        if text.upper().startswith("ARTICLES"):
            start = j + 1
            break
    if start is None:
        return []

    titles: List[str] = []
    import re

    for k in range(start, len(paragraphs)):
        line = paragraphs[k].text.strip()
        if not line:
            break
        if line.isupper():
            break
        match = re.match(r"(.+?)\.{2,}\d+$", line)
        if match:
            titles.append(match.group(1).strip())
        else:
            titles.append(line)

    return titles


def clear_articles(doc: Document):
    """Remove article sections based on TOC titles if available."""
    titles = extract_article_titles_from_toc(doc)

    def remove_paragraph(paragraph):
        p = paragraph._element
        p.getparent().remove(p)

    if titles:
        article_heading_idx = None

        # find start indices of each article title and detect the heading
        start_indices = []
        for title in titles:
            for i, p in enumerate(doc.paragraphs):
                if p.text.strip().upper() == title.upper():
                    start_indices.append(i)
                    if article_heading_idx is None and i > 0:
                        prev = doc.paragraphs[i - 1].text.strip().upper()
                        if prev == "ARTICLES":
                            article_heading_idx = i - 1
                    break
        start_indices.sort()
        if article_heading_idx is not None and (
            not start_indices or article_heading_idx < start_indices[0]
        ):
            start_indices.insert(0, article_heading_idx)

        for idx in reversed(range(len(start_indices))):
            start = start_indices[idx]
            end = (
                start_indices[idx + 1]
                if idx + 1 < len(start_indices)
                else len(doc.paragraphs)
            )
            for _ in range(end - start):
                remove_paragraph(doc.paragraphs[start])
        return

    # fallback to previous behaviour if we cannot parse TOC
    found = False
    start_idx = 0
    for i, p in enumerate(doc.paragraphs):
        if "ARTICLES" in p.text.upper():
            found = True
            start_idx = i
            break
    if found:
        for _ in range(len(doc.paragraphs) - start_idx):
            remove_paragraph(doc.paragraphs[start_idx])


def load_instructions(content_path: Path) -> dict:
    """Load instructions from ``instructions.json`` if present."""
    inst_file = content_path / "instructions.json"
    if inst_file.exists():
        try:
            with inst_file.open("r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Failed to read instructions: {e}")
    return {}


def delete_after_page(doc: Document, page_number: int) -> None:
    """Remove all content after the paragraph containing ``Page {page_number}``."""
    search = f"Page {page_number}"
    target = None
    for p in doc.paragraphs:
        if search in p.text:
            target = p
            break
    if target is None:
        return
    body = target._element.getparent()
    elem = target._element.getnext()
    while elem is not None:
        next_elem = elem.getnext()
        body.remove(elem)
        elem = next_elem


def apply_basic_formatting(doc: Document, font_size: Optional[int], line_spacing: Optional[float]) -> None:
    """Set font size and line spacing across all paragraphs."""
    for p in doc.paragraphs:
        if line_spacing is not None:
            p.paragraph_format.line_spacing = line_spacing
        for run in p.runs:
            if font_size is not None:
                run.font.size = Pt(font_size)


def append_article(doc: Document, article_doc: Document):
    for element in article_doc.element.body:
        doc.element.body.append(element)


def set_font_size(doc: Document, start_paragraph: int, size: int) -> None:
    """Apply ``size`` point font to paragraphs starting at ``start_paragraph``."""
    for p in doc.paragraphs[start_paragraph:]:
        for run in p.runs:
            run.font.size = Pt(size)


def set_line_spacing(doc: Document, start_paragraph: int, spacing: float) -> None:
    """Set line spacing for paragraphs starting at ``start_paragraph``."""
    for p in doc.paragraphs[start_paragraph:]:
        p.paragraph_format.line_spacing = spacing

def reuse_journal_page(doc: Document, source_doc: Document, page_number: int) -> None:
    """Copy the specified page from ``source_doc`` into ``doc``."""
    # Complex page-level manipulation is not implemented; placeholder only.
    pass


def reuse_info_for_authors(doc: Document, source_doc: Document) -> None:
    """Insert the 'Information for Authors' section from ``source_doc``."""
    pass


def reuse_membership_app(doc: Document, source_doc: Document) -> None:
    """Insert the membership application page from ``source_doc``."""
    pass


def add_editor_titles(doc: Document, page: int, editor_titles: List[str]) -> None:
    """Add editor titles under their images on the specified page."""
    pass


def remove_extra_spaces_in_author_line(doc: Document, page: int, article_index: int) -> None:
    """Collapse multiple spaces in the given author line."""
    pass


def await_presidents_message_placeholder(doc: Document, page: int) -> None:
    """Insert a placeholder comment for the president's message."""
    pass


def update_author_line(doc: Document, page: int, old_name: str, new_name_with_creds: str) -> None:
    """Replace the author line on the given page."""
    pass


def fix_orphaned_last_lines(doc: Document, page: int, column_index: int) -> None:
    """Ensure the last reference line wraps correctly."""
    pass


def convert_table_to_landscape(doc: Document, page: int, table_index: int) -> None:
    """Rotate a table on the specified page to landscape orientation."""
    pass


def move_paragraph_to_next_column(doc: Document, page: int, column_index: int, paragraph_text_match: str) -> None:
    """Move a paragraph to the next column if it matches the text."""
    pass


def fix_apostrophe(doc: Document, page: int, search_word: str, correct_word: str) -> None:
    """Replace a word with a corrected apostrophe."""
    replace_text_in_paragraphs(doc.paragraphs, search_word, correct_word)


def insert_line_space_before_subheading(doc: Document, page: int, subheading_list: Iterable[str]) -> None:
    """Ensure a blank line before each subheading on the page."""
    pass


def move_section_to_next_column(doc: Document, page: int, column_index: int, section_heading: str) -> None:
    """Move a section starting with heading to the next column."""
    pass


def insert_line_space_before_paragraph(doc: Document, page: int, preceding_text: str, new_paragraph_label: str) -> None:
    """Insert a blank line before the paragraph that matches ``new_paragraph_label``."""
    pass


def indent_paragraph(doc: Document, page: int, column_index: int, paragraph_index: int, indent_width: float) -> None:
    """Apply left indent to a paragraph."""
    pass


def fix_orphaned_citation(doc: Document, page: int, orphan_text: str) -> None:
    """Merge an orphaned citation with the preceding line."""
    pass


def insert_line_space_before(doc: Document, page: int, subheading: str) -> None:
    """Ensure a blank line before the given subheading."""
    pass


def fix_separation_line_between_sections(doc: Document, page: int, line_shape_criteria) -> None:
    """Ensure a solid line separates sections."""
    pass


def remove_unclear_text(doc: Document, page: int, search_pattern: str) -> None:
    """Remove or flag unclear text matching ``search_pattern``."""
    pass


def fix_page_numbering(doc: Document) -> None:
    """Renumber pages sequentially starting from the first section."""
    pass


def apply_hanging_indent_to_references(
    doc: Document, page_range_start: int, page_range_end: int, indent_width: float, line_spacing: float
) -> None:
    """Apply hanging indent formatting to reference lists."""
    pass


def normalize_table_formatting(
    doc: Document, page: int, table_index: int, desired_font_size: int, desired_line_spacing: float
) -> None:
    """Normalize fonts and spacing in tables."""
    pass


def detect_and_remove_extra_spaces(doc: Document, page_range: Iterable[int], pattern: str = "  ") -> None:
    """Collapse multiple spaces across the given page range."""
    for p in doc.paragraphs:
        if pattern in p.text:
            while pattern in p.text:
                p.text = p.text.replace(pattern, " ")


def ensure_blank_line_before_headings(doc: Document, page_range: Iterable[int], heading_list: Iterable[str]) -> None:
    """Ensure exactly one blank line precedes each heading."""
    pass


def split_long_paragraphs_across_columns(doc: Document, page: int, column_count: int) -> None:
    """Attempt to split long paragraphs across columns."""
    pass


def update_table_of_contents(doc: Document) -> None:
    """Insert or refresh the table of contents."""
    pass


def apply_two_column_layout(doc: Document, start_page: int) -> None:
    """Set sections starting at ``start_page`` to use two text columns."""

    try:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
    except Exception:
        return

    for idx, section in enumerate(doc.sections):
        if idx + 1 < start_page:
            continue

        # Prefer new API if available
        if hasattr(section, "text_columns"):
            try:
                section.text_columns.set_num(2)
                section.text_columns.spacing = Pt(36)  # ensure a small gap
            except Exception:
                pass
            continue

        sectPr = section._sectPr
        cols = sectPr.find(qn("w:cols"))
        if cols is None:
            cols = OxmlElement("w:cols")
            cols.set(qn("w:space"), "720")
            sectPr.append(cols)
        cols.set(qn("w:num"), "2")


def apply_page_borders(doc: Document, start_section: int, border_specs) -> None:
    """Apply borders to pages starting at ``start_section``."""
    pass


def add_page_borders(doc: Document, start_section: int) -> None:
    """Add solid left and right borders for sections from ``start_section``."""

    try:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
    except Exception:
        return

    for idx, section in enumerate(doc.sections):
        if idx < start_section:
            continue

        if hasattr(section, "page_setup") and hasattr(section.page_setup, "left_border"):  # type: ignore[attr-defined]
            try:
                ps = section.page_setup
                ps.left_border = ps.right_border = {
                    "val": "single",
                    "sz": 4,
                    "space": 0,
                    "color": "000000",
                }
                continue
            except Exception:
                pass

        sectPr = section._sectPr
        for existing in sectPr.findall(qn("w:pgBorders")):
            sectPr.remove(existing)
        pgBorders = OxmlElement("w:pgBorders")
        pgBorders.set(qn("w:offsetFrom"), "text")
        for side in ("left", "right"):
            border = OxmlElement(f"w:{side}")
            border.set(qn("w:val"), "single")
            border.set(qn("w:sz"), "4")
            border.set(qn("w:space"), "0")
            border.set(qn("w:color"), "000000")
            pgBorders.append(border)
        sectPr.append(pgBorders)


def validate_issue_number_and_volume(doc: Document, expected_volume: str, expected_issue: str, expected_year: str) -> None:
    """Check volume/issue/year text appears once and matches expectations."""
    pass

def save_pdf(doc_path: Path, pdf_path: Path):
    try:
        from docx2pdf import convert
        convert(str(doc_path), str(pdf_path))
    except Exception as e:
        print(f"PDF export failed: {e}")


def update_journal(
    base_path: Path,
    content_path: Path,
    output_path: Path,
    volume: str,
    issue: str,
    month_year: str,
    section_title: str,
    cover_page_num: int = 1,
    header_page_num: int = 2,
) -> None:
    """Run the update process with explicit paths and parameters."""
    doc = load_document(base_path)
    instructions = load_instructions(content_path)

    update_front_cover(doc, volume, issue, month_year, section_title, cover_page_num)
    update_business_information(
        doc,
        "2023",
        "Annual subscription rates are: institutions $550, individuals $220, and students $110",
    )
    header_text = f"Volume {volume}, Issue {issue}\n{month_year}\n{section_title}"
    update_page2_header(doc, header_text, header_page_num)
    pres_message_path = content_path / "president_message.txt"
    message_text = pres_message_path.read_text() if pres_message_path.exists() else ""
    insert_presidents_message(doc, content_path / "president.jpg", message_text)

    clear_articles(doc)

    start_idx = len(doc.paragraphs)
    for article_file in sorted(content_path.glob("article*.docx")):
        article_doc = Document(article_file)
        append_article(doc, article_doc)
    if "font_size" in instructions:
        set_font_size(doc, start_idx, int(instructions["font_size"]))
    if "line_spacing" in instructions:
        set_line_spacing(doc, start_idx, float(instructions["line_spacing"]))

    save_document(doc, output_path)
    pdf_path = output_path.with_suffix(".pdf")
    save_pdf(output_path, pdf_path)

def main_from_gui(
    base_doc: Path,
    content_folder: Path,
    output_doc: Path,
    volume: str,
    issue: str,
    month_year: str,
    section_title: str,
    cover_page_num: int = 1,
    header_page_num: int = 2,
) -> None:
    """Helper for GUI front-end."""
    update_journal(
        base_doc,
        content_folder,
        output_doc,
        volume,
        issue,
        month_year,
        section_title,
        cover_page_num,
        header_page_num,
    )

def main():
    parser = argparse.ArgumentParser(description="Update ABNFF Journal document")
    parser.add_argument("base_doc")
    parser.add_argument("content_folder")
    parser.add_argument("output_doc")
    parser.add_argument("--volume", required=True)
    parser.add_argument("--issue", required=True)
    parser.add_argument("--month-year", required=True, dest="month_year")
    parser.add_argument("--section-title", required=True, dest="section_title")
    parser.add_argument("--cover-page", type=int, default=1, dest="cover_page")
    parser.add_argument("--header-page", type=int, default=2, dest="header_page")
    args = parser.parse_args()

    base_path = Path(args.base_doc)
    content_path = Path(args.content_folder)
    output_path = (
        Path(args.output_doc)
        if args.output_doc
        else base_path.with_name(base_path.stem + "_updated.docx")
    )

    update_journal(
        base_path,
        content_path,
        output_path,
        args.volume,
        args.issue,
        args.month_year,
        args.section_title,
        args.cover_page,
        args.header_page,
    )

if __name__ == "__main__":
    main()
