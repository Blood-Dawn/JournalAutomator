import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import journal_updater.journal_updater as ju


def test_update_journal_appends_articles(tmp_path):
    base = ju.Document()
    base.add_paragraph("ARTICLES")
    base_path = tmp_path / "base.docx"
    base.save(base_path)

    content_dir = tmp_path / "content"
    content_dir.mkdir()

    art1 = ju.Document()
    art1.add_paragraph("First article")
    art1.save(content_dir / "Article1.docx")

    art2 = ju.Document()
    art2.add_paragraph("Second article")
    art2.save(content_dir / "article2.DOCX")

    out_path = tmp_path / "out.docx"
    ju.update_journal(
        base_path,
        content_dir,
        out_path,
        "1",
        "1",
        "June 2025",
    )

    result = ju.Document(out_path)
    texts = [p.text for p in result.paragraphs]
    assert "First article" in texts
    assert "Second article" in texts
