import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import journal_updater.journal_updater as ju


def test_preserve_editorial_sections(tmp_path):
    base = ju.Document()
    base.add_paragraph("President's Message")
    base.add_paragraph("Old message")
    base.add_paragraph("First Editorial")
    base.add_paragraph("Old first text")
    base.add_paragraph("Second Editorial")
    base.add_paragraph("Old second text")
    base.add_paragraph("ARTICLES")
    base.add_paragraph("Old article")
    base_path = tmp_path / "base.docx"
    base.save(base_path)

    content = tmp_path / "content"
    content.mkdir()
    art = ju.Document()
    art.add_paragraph("New article text")
    art.save(content / "article1.docx")

    out_path = tmp_path / "out.docx"
    ju.update_journal(base_path, content, out_path, "1", "1", "June 2025", "Articles")

    result = ju.Document(out_path)
    texts = [p.text for p in result.paragraphs]
    assert "President's Message" in texts
    assert "First Editorial" in texts
    assert "Second Editorial" in texts
    assert "Old article" not in texts
    assert "New article text" in texts
