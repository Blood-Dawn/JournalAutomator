import sys
import os
import json

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import journal_updater.journal_updater as ju


def test_delete_after_page(tmp_path):
    base_path = tmp_path / "base.docx"
    base = ju.Document()
    base.add_paragraph("Volume 1, Issue 3")
    base.add_paragraph("Old text")
    base.save(base_path)

    content_dir = tmp_path / "content"
    content_dir.mkdir()
    (content_dir / "instructions.json").write_text(json.dumps({"delete_after_page": 1}))

    out_path = tmp_path / "out.docx"
    ju.update_journal(base_path, content_dir, out_path, "1", "1", "June 2025", "Updates", 1, 2)

    result = ju.Document(out_path)
    texts = [p.text for p in result.paragraphs]
    assert len(texts) == 1
    assert "Page 1" in texts[0]
