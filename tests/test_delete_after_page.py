import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import journal_updater.journal_updater as ju
from docx.enum.text import WD_BREAK


def test_delete_after_page():
    doc = ju.Document()
    first = doc.add_paragraph("Keep this")
    first.add_run().add_break(WD_BREAK.PAGE)
    doc.add_paragraph("Remove")

    ju.delete_after_page(doc, 1)

    texts = [p.text for p in doc.paragraphs]
    assert texts == ["Keep this"]
