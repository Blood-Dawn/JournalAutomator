import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import journal_updater.journal_updater as ju


def _paragraph_texts(doc):
    return [p.text for p in doc.paragraphs]


def test_clear_articles_with_toc():
    doc = ju.Document()
    doc.add_paragraph("Table of Contents")
    doc.add_paragraph("ARTICLES")
    doc.add_paragraph("First Article........................1")
    doc.add_paragraph("Second Article.......................5")
    doc.add_paragraph("")
    doc.add_paragraph("OTHER")

    doc.add_paragraph("ARTICLES")
    doc.add_paragraph("First Article")
    doc.add_paragraph("f content")
    doc.add_paragraph("Second Article")
    doc.add_paragraph("s content")
    doc.add_paragraph("OTHER")

    ju.clear_articles(doc)
    texts = _paragraph_texts(doc)
    assert "First Article" not in texts
    assert "Second Article" not in texts
    assert "f content" not in texts
    assert "s content" not in texts
    assert texts[-1] == "OTHER"


def test_clear_articles_fallback():
    doc = ju.Document()
    doc.add_paragraph("Intro")
    doc.add_paragraph("ARTICLES")
    doc.add_paragraph("Old text")
    doc.add_paragraph("More")

    ju.clear_articles(doc)
    texts = _paragraph_texts(doc)
    assert texts == ["Intro"]

