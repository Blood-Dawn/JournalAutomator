import sys
import os
import json

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import journal_updater.journal_updater as journal_updater


def test_replace_text():
    class FakeRun:
        def __init__(self, text):
            self.text = text

    class FakeParagraph:
        def __init__(self, text):
            self.text = text
            self.runs = [FakeRun(text)]

    paragraphs = [FakeParagraph("2023 subscription info"), FakeParagraph("nothing")]
    journal_updater.replace_text_in_paragraphs(paragraphs, "2023", "")
    assert paragraphs[0].runs[0].text == " subscription info"

def test_load_and_update_front_cover(tmp_path):
    doc_path = tmp_path / "base.docx"
    doc = journal_updater.Document()
    p = doc.add_paragraph("Volume 1, Issue 3")
    doc.save(doc_path)

    loaded = journal_updater.load_document(doc_path)
    journal_updater.update_front_cover(loaded, "1", "1", "June 2025", "Update Articles", 1)
    assert "June 2025" in loaded.paragraphs[0].text


def test_update_page2_header(tmp_path):
    doc = journal_updater.Document()
    section = doc.sections[0]
    header_p = section.header.paragraphs[0]
    header_p.text = "Old header"
    journal_updater.update_page2_header(doc, "Volume 1", 2)
    assert "Page 2" in section.header.paragraphs[0].text


def test_font_utils():
    doc = journal_updater.Document()
    doc.add_paragraph("p0")
    doc.add_paragraph("p1")

    journal_updater.set_font_size(doc, 1, 16)
    journal_updater.set_line_spacing(doc, 1, 1.5)

    assert doc.paragraphs[1].runs[0].font.size.pt == 16
    assert doc.paragraphs[1].paragraph_format.line_spacing == 1.5


def test_update_journal_formatting(tmp_path):
    base = journal_updater.Document()
    base.add_paragraph("ARTICLES")
    base_path = tmp_path / "base.docx"
    base.save(base_path)

    content_dir = tmp_path / "content"
    content_dir.mkdir()
    art = journal_updater.Document()
    art.add_paragraph("Article text")
    art.save(content_dir / "article1.docx")

    import json

    (content_dir / "instructions.json").write_text(
        json.dumps({"font_size": 14, "line_spacing": 2})
    )

    out_path = tmp_path / "out.docx"
    journal_updater.update_journal(base_path, content_dir, out_path)
    result = journal_updater.Document(out_path)

    assert result.paragraphs[1].runs[0].font.size.pt == 14
    assert result.paragraphs[1].paragraph_format.line_spacing == 2