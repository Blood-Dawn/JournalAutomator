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
Recent additions provide helpers for formatting the front cover and centering footer text across all sections.
An additional helper `add_page_borders_with_rule(doc, start_section, add_center_line=False)`
adds left and right borders and can optionally draw a vertical rule down the center of each page.

For non‑technical users a small Tkinter GUI is provided. Launch it with:

```
python -m journal_updater.gui
```

The window lets you choose the base DOCX, the content folder, and where the
output should be saved. It also collects the volume, issue, date and the page
numbers used for the cover and where new articles should start. Additional
fields let you enter a default font size, line spacing and font family. These
settings are written to an `instructions.json` file in the selected content
folder. Clicking **Run Update** performs the same steps as the command line
script.
## Usage

```
python journal_updater/journal_updater.py BASE_DOCX CONTENT_FOLDER OUTPUT_DOCX \
    --volume 1 --issue 1 --month-year "June 2025" \
    --cover-page 1 --start-page 3
```

- **BASE_DOCX**: path to the previous issue's Word file (e.g. December 2024).
- **CONTENT_FOLDER**: path to a folder containing new resources. This folder may
  include:
  - `president_message.txt` – new President's Message text.
  - `article*.docx` – Word documents for each new article.
- **OUTPUT_DOCX**: path where the updated issue should be saved.
- **--volume / --issue / --month-year**: details for the new
  issue's front matter.
- **--cover-page**: page number used on the front cover.
- **--start-page**: first page where the imported articles should be placed.

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
7. Applies optional front-cover formatting.
8. Centers the footer layout across all pages.
9. Inserts a vertical center rule when using `add_page_borders_with_rule`.

Ensure your base document includes a Table of Contents with an
**ARTICLES** heading so article titles can be detected and removed.

The implementation is intentionally minimal and serves as a starting
point for further automation as outlined in the program goals.

### instructions.json

An optional `instructions.json` file may be placed in the content folder to control certain aspects of the update. The `format_front_and_footer` flag triggers automatic styling of the front page and footer sections. Supported keys are:

- `volume` – volume number for the issue.
- `issue` – issue number.
- `delete_after_page` – remove all content after this page number.
- `delete_after_editorial` – remove all content after the last editorial page.
- `cleanup_black_lines` – remove duplicate separation lines on each page.
- `font_size` – default font size to apply to all text (in points).
- `line_spacing` – line spacing value (e.g. `1.0` or `1.15`).
- `font_family` – default font family name to apply across the document.
- `format_front_and_footer` – optional block with `font_size` and
  `line_spacing` to style the front cover paragraph and all footers.

When present, the `volume` and `issue` values override any command line or GUI
inputs.

Example file:

```json
{
  "volume": "2",
  "issue": "3",
  "delete_after_page": 2,
  "delete_after_editorial": false,
  "cleanup_black_lines": true,
  "font_size": 10,
  "line_spacing": 1.0,
  "font_family": "Times New Roman",
  "format_front_and_footer": {
    "font_size": 14,
    "line_spacing": 1.2
  }
}
```

If no file is present, default values are used.
