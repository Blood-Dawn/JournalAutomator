# JournalAutomator

This repository contains a simple Python utility that automates some of the repetitive
steps required to update an ABNFF Journal Word document from one issue to the next.

The `journal_updater.py` script uses `python-docx` to modify a base Word document and
applies new text and article content from a provided folder. Once the document is
updated it can optionally be exported to PDF using `docx2pdf`.

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
   content folder.
6. Saves the resulting document and optionally attempts to export a PDF
   alongside it (requires `docx2pdf`).

The implementation is intentionally minimal and serves as a starting
point for further automation as outlined in the program goals.
