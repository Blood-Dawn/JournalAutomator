import argparse
from pathlib import Path
from typing import Optional
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt


def replace_text_in_paragraphs(paragraphs, search_text, replace_text):
    for p in paragraphs:
        if search_text in p.text:
            inline = p.runs
            for i in range(len(inline)):
                if search_text in inline[i].text:
                    inline[i].text = inline[i].text.replace(search_text, replace_text)


def update_front_cover(doc: Document, volume: str, issue: str, month_year: str):
    """Update the front cover volume/issue/date text"""
    search = "Volume"
    for p in doc.paragraphs:
        if search in p.text:
            p.text = f"Volume {volume}, Issue {issue}\n{month_year}\nUpdate Articles"
            for run in p.runs:
                run.font.bold = True
            break


def update_business_information(doc: Document):
    """Remove stray 2023 from business information block"""
    replace_text_in_paragraphs(doc.paragraphs, "2023", "")


def update_page2_header(doc: Document, volume: str, issue: str, month_year: str):
    for section in doc.sections:
        header = section.header
        replace_text_in_paragraphs(header.paragraphs, "2024", month_year.split()[1])
        replace_text_in_paragraphs(header.paragraphs, "Number 3", f"Issue {issue}")


def insert_president_message(doc: Document, message_path: Path):
    text = message_path.read_text() if message_path.exists() else "[PRESIDENT'S MESSAGE]"
    for p in doc.paragraphs:
        if "President's Message" in p.text:
            idx = doc.paragraphs.index(p)
            doc.paragraphs[idx + 1].text = text
            break


def clear_articles(doc: Document):
    found = False
    start_idx = 0
    for i, p in enumerate(doc.paragraphs):
        if "ARTICLES" in p.text.upper():
            found = True
            start_idx = i
            break
    if found:
        for _ in range(len(doc.paragraphs) - start_idx):
            doc.paragraphs.pop()


def append_article(doc: Document, article_doc: Document):
    for element in article_doc.element.body:
        doc.element.body.append(element)


def save_pdf(doc_path: Path, pdf_path: Path):
    try:
        from docx2pdf import convert
        convert(str(doc_path), str(pdf_path))
    except Exception as e:
        print(f"PDF export failed: {e}")


def main():
    parser = argparse.ArgumentParser(description="Update ABNFF Journal document")
    parser.add_argument("base_doc")
    parser.add_argument("content_folder")
    parser.add_argument("output_doc")
    args = parser.parse_args()

    base_path = Path(args.base_doc)
    content_path = Path(args.content_folder)
    output_path = Path(args.output_doc)

    doc = Document(base_path)

    update_front_cover(doc, "1", "1", "June 2025")
    update_business_information(doc)
    update_page2_header(doc, "1", "1", "June 2025")

    pres_message = content_path / "president_message.txt"
    insert_president_message(doc, pres_message)

    clear_articles(doc)

    for article_file in sorted(content_path.glob("article*.docx")):
        article_doc = Document(article_file)
        append_article(doc, article_doc)

    doc.save(output_path)

    pdf_path = output_path.with_suffix('.pdf')
    save_pdf(output_path, pdf_path)


if __name__ == "__main__":
    main()
