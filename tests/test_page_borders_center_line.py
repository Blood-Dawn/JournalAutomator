import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from journal_updater import journal_updater as ju
from docx.oxml.ns import qn


def test_add_page_borders_with_center_line():
    doc = ju.Document()
    ju.add_page_borders_with_rule(doc, 0, add_center_line=True)

    header_el = doc.sections[0].header._element
    shapes = header_el.findall('.//' + qn('v:shape'))
    assert len(shapes) == 1
