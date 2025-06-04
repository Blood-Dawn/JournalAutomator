import sys, os; sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
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
