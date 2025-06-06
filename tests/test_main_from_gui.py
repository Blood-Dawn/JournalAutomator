import os
import sys
import json
from pathlib import Path

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import journal_updater.journal_updater as ju


def _build_old_journal(path: Path) -> None:
    doc = ju.Document()
    doc.add_paragraph("Volume 1, Issue 1")
    doc.add_paragraph("Old date line")
    doc.add_paragraph("Table of Contents")
    doc.add_paragraph("ARTICLES")
    doc.add_paragraph("Old Article................1")
    doc.add_paragraph("OTHER")
    doc.add_paragraph("ARTICLES")
    doc.add_paragraph("Old Article")
    doc.add_paragraph("Old article text")
    doc.save(path)


def _build_blank_article(path: Path) -> None:
    doc = ju.Document()
    doc.add_paragraph("New Article Body")
    doc.save(path)


def test_main_from_gui(tmp_path):
    base_path = tmp_path / "base.docx"
    _build_old_journal(base_path)

    content = tmp_path / "content"
    content.mkdir()
    _build_blank_article(content / "article1.docx")

    out_path = tmp_path / "out.docx"

    ju.main_from_gui(
        base_path,
        content,
        out_path,
        volume="2",
        issue="3",
        month_year="July 2025",
        cover_page_num=1,
        start_page=3,
    )

    result = ju.Document(out_path)
    texts = [p.text for p in result.paragraphs]

    assert "Volume 2, Issue 3" in texts[0]
    assert "New Article Body" in texts
    assert "Old Article" not in texts


def test_main_from_gui_font_options(tmp_path):
    base_path = tmp_path / "base.docx"
    _build_old_journal(base_path)

    content = tmp_path / "content"
    content.mkdir()
    _build_blank_article(content / "article1.docx")

    out_path = tmp_path / "out.docx"

    ju.main_from_gui(
        base_path,
        content,
        out_path,
        volume="2",
        issue="3",
        month_year="July 2025",
        cover_page_num=1,
        start_page=3,
        font_size=14,
        line_spacing=1.5,
        font_family="Arial",
    )

    inst = json.load((content / "instructions.json").open())
    assert inst["font_size"] == 14
    assert inst["line_spacing"] == 1.5
    assert inst["font_family"] == "Arial"

    result = ju.Document(out_path)
    para = next(p for p in result.paragraphs if "New Article Body" in p.text)
    assert para.runs[0].font.size.pt == 14
    assert para.paragraph_format.line_spacing == 1.5
    assert para.runs[0].font.name == "Arial"
