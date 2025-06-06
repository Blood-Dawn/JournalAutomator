import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import journal_updater.journal_updater as ju


def _paragraph_texts(doc):
    return [p.text for p in doc.paragraphs]


def test_update_assistant_editors_removes_line():
    doc = ju.Document()
    doc.add_paragraph("Assistant Editors")
    doc.add_paragraph("Alice")
    doc.add_paragraph("Bob")

    ju.update_assistant_editors(doc, "Alice")
    texts = _paragraph_texts(doc)
    assert "Alice" not in texts
    assert texts == ["Assistant Editors", "Bob"]
    assert "" not in texts
