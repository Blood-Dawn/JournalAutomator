# JournalAutomator

This repository contains a simple Python utility that automates some of the repetitive
steps required to update an ABNFF Journal Word document from one issue to the next.

The `journal_updater.py` script uses `python-docx` to modify a base Word document and
applies new text and article content from a provided folder. Once the document is
updated it can optionally be exported to PDF using `docx2pdf`.

Many helper functions are provided so that future automation steps can call them
individually (e.g. `update_front_cover`, `update_page2_header`, `clear_articles`,
and more). Most advanced operations are currently placeholders but documented
for future work.

For non‑technical users a small Tkinter GUI is provided. Launch it with:

```
python -m journal_updater.gui
```

The window lets you choose the base DOCX, the content folder, and where the
output should be saved. Clicking **Run Update** performs the same steps as the
command line script.
## Usage

```
python journal_updater/journal_updater.py BASE_DOCX CONTENT_FOLDER OUTPUT_DOCX
```

- **BASE_DOCX**: path to the previous issue's Word file (e.g. December 2024).
- **CONTENT_FOLDER**: path to a folder containing new resources. This folder may
  include:
  - `president_message.txt` – new President's Message text.
  - `article*.docx` – Word documents for each new article.
- **OUTPUT_DOCX**: path where the updated June issue should be saved.

The script performs a handful of automated replacements:

1. Updates front cover text with the new volume, issue, and date.
2. Cleans the business information block of outdated years.
3. Updates page 2 header text.
4. Inserts the President's Message from `president_message.txt`.
5. Clears all old articles and appends the new ones found in the
   content folder. The removal step relies on article titles listed
   under the **ARTICLES** section of the Table of Contents.
6. Saves the resulting document and optionally attempts to export a PDF
   alongside it (requires `docx2pdf`).

Ensure your base document includes a Table of Contents with an
**ARTICLES** heading so article titles can be detected and removed.

The implementation is intentionally minimal and serves as a starting
point for further automation as outlined in the program goals.
