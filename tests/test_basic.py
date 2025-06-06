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


def test_update_journal_with_instructions(tmp_path):
    base_doc = tmp_path / "base.docx"
    doc = journal_updater.Document()
    doc.add_paragraph("Volume 1, Issue 1")
    doc.save(base_doc)
    content = tmp_path / "content"
    content.mkdir()

    instructions = {
        "volume": "2",
        "issue": "3",
        "font_size": 10,
        "line_spacing": 1.0,
    }
    (content / "instructions.json").write_text(json.dumps(instructions))

    output_doc = tmp_path / "out.docx"

    journal_updater.update_journal(base_doc, content, output_doc)

    out = journal_updater.load_document(output_doc)
    assert "Volume 2" in out.paragraphs[0].text
    assert out.paragraphs[0].paragraph_format.line_spacing == 1.0