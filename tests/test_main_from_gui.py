import os
import sys
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
        section_title="Update Articles",
        cover_page_num=1,
        header_page_num=2,
    )

    result = ju.Document(out_path)
    texts = [p.text for p in result.paragraphs]

    assert "Volume 2, Issue 3" in texts[0]
    assert "New Article Body" in texts
    assert "Old Article" not in texts
