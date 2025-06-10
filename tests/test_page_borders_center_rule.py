import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from journal_updater import journal_updater as ju


def test_add_page_borders_with_rule(tmp_path):
    doc = ju.Document()
    ju.add_page_borders_with_rule(doc, 0, add_center_line=True)

    from docx.oxml.ns import qn

    shapes = list(doc.sections[0].header._element.findall('.//' + qn('w:shape')))
    assert len(shapes) == 1
