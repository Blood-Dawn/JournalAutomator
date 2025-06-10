import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import journal_updater.journal_updater as ju


def test_update_journal_appends_articles(tmp_path):
    base = ju.Document()
    base.sections[0].header.paragraphs[0].text = "Base header"
    base.add_paragraph("ARTICLES")
    base_path = tmp_path / "base.docx"
    base.save(base_path)

    content_dir = tmp_path / "content"
    content_dir.mkdir()

    art1 = ju.Document()
    art1.sections[0].header.paragraphs[0].text = "H2"
    art1.sections[0].footer.paragraphs[0].text = "F2"
    art1.add_paragraph("Second article")
    art1.save(content_dir / "articleB.docx")

    art2 = ju.Document()
    art2.sections[0].header.paragraphs[0].text = "H1"
    art2.sections[0].footer.paragraphs[0].text = "F1"
    art2.add_paragraph("First article")
    art2.save(content_dir / "articleA.docx")

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
    first_index = texts.index("First article")
    second_index = texts.index("Second article")
    assert first_index < second_index
    headers = [p.text for s in result.sections for p in s.header.paragraphs]
    footers = [p.text for s in result.sections for p in s.footer.paragraphs]
    assert "H1" not in headers and "H2" not in headers
    assert "F1" not in footers and "F2" not in footers


def test_import_articles_sorted_and_clean(tmp_path):
    doc = ju.Document()
    doc.sections[0].header.paragraphs[0].text = "Base"
    doc.add_paragraph("Start")

    art_b = ju.Document()
    art_b.sections[0].header.paragraphs[0].text = "HB"
    art_b.sections[0].footer.paragraphs[0].text = "FB"
    art_b.add_paragraph("Second")
    path_b = tmp_path / "b.docx"
    art_b.save(path_b)

    art_a = ju.Document()
    art_a.sections[0].header.paragraphs[0].text = "HA"
    art_a.sections[0].footer.paragraphs[0].text = "FA"
    art_a.add_paragraph("First")
    path_a = tmp_path / "a.docx"
    art_a.save(path_a)

    ju.import_articles(doc, [path_b, path_a])

    out = tmp_path / "result.docx"
    doc.save(out)
    result = ju.Document(out)
    texts = [p.text for p in result.paragraphs]
    assert texts[1:] == ["First", "Second"]
    headers = [p.text for s in result.sections for p in s.header.paragraphs]
    footers = [p.text for s in result.sections for p in s.footer.paragraphs]
    assert "HA" not in headers and "HB" not in headers
    assert "FA" not in footers and "FB" not in footers
