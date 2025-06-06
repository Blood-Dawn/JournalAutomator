# JournalAutomator Instructions

This project provides various helper functions to automate updates to the ABNFF Journal. Each helper can be called individually from your own scripts.

## Layout helpers

- `apply_two_column_layout(doc, start_page)` – convert sections starting at `start_page` to a two‑column layout. The function relies on `Section.text_columns` where available and falls back to directly editing the XML.
- `add_page_borders(doc, start_section)` – add simple left and right borders starting with the specified section.

Use these helpers in combination with other utilities defined in `journal_updater.py` to customize your journal layout.

## Front cover and footer

- `format_front_cover(doc)` – apply basic styling to the first paragraph on the front page.
- `layout_footer(doc)` – center footer text across all sections.
- `format_front_and_footer(doc)` – call both helpers at once.
