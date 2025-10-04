# CSV EDITOR PRO

A lightweight, keyboard-friendly CSV editor for power users and data tinkerers.

**Created by:** Sagar Hodar
**Version:** 2.0

---

## 📌 HOW TO USE

Save CSV_EDITOR_PRO.ps1 to local windows system and right click and choose "Run with powershell" DONE!!

---

## 📌 Overview

CSV EDITOR PRO is a compact CSV editing tool focused on quick cell editing, column-level filtering, and fast keyboard-driven workflows. It works with UTF-8 CSV files and provides features to open, edit, filter, and export CSV data while keeping control of rows and columns.

---

## ✨ Key Features

* **Edit cells directly** — Click a cell or press `F2` to edit inline.
* **Global Search** — Search across all columns to quickly locate text or numbers.
* **Column Filters** — Per-column filters supporting numeric comparisons and text matching:

  * `>5`, `<10`, `>=20`, `<=100`, `=exact`, or plain text search.
* **Add / Delete Rows** — Insert a new row next to the selected row, or delete selected rows.
* **Add / Delete Columns** — Insert a new column next to the selected column, or delete columns.
* **Open CSV** — Load existing CSV files from disk.
* **New CSV** — Create a blank CSV file to start from scratch.
* **Save** — Save changes to the current file.
* **Save As** — Save with a new filename. When a filter is active, *Save As* exports only the filtered rows.
* **UTF-8** — All saved CSVs use UTF-8 encoding.

---

## ⌨️ Keyboard Shortcuts

* `F2` or **Click**: Edit the currently focused cell.
* `Ctrl + Click`: Select multiple rows.
* `Delete`: Remove the selected cell content (or delete selected rows depending on UI state).

> Tip: Keyboard navigation (arrow keys / Enter / Tab) should move focus between cells for fast editing.

---

## 🔎 Filtering & Search

* Use the **Global Search** box to find values across the entire sheet.
* Use per-column filter inputs to restrict rows. Supported operators:

  * Numeric: `>`, `<`, `>=`, `<=` followed by a number.
  * Exact match: `=value`.
  * Text contains: type text directly (case-insensitive by default).

When multiple columns have active filters, rows must satisfy **all** column filters (AND logic).

---

## 🧭 Common Workflows

* **Quick edit**: Click a cell (or press `F2`) → type → `Enter` to save the cell.
* **Insert row**: Select a row → use the `Insert Row` action (or context menu) → a new row appears next to the selection.
* **Export filtered data**: Apply filters → `File → Save As` → choose filename → exported CSV will contain only visible rows.

---

## 💾 File Handling

* The app opens and saves standard CSV files.
* `Save` overwrites the current file (if opened from disk) or prompts for a filename if it’s a new file.
* `Save As` always prompts for a destination filename and exports the currently visible (filtered) data.
* Files are saved in **UTF-8** and use a comma (`,`) as the default delimiter. If your CSV uses a different delimiter, convert it beforehand or add support in the import dialog.

---

## 🛠️ Tips & Notes

* Rows and columns are inserted **next to** the currently selected item.
* If many filters are active, remember `Save As` will export only the filtered subset.
* Use `Ctrl + Click` to build non-contiguous row selections for bulk operations.

---

## 🚀 Contributing

If you'd like to contribute improvements or fixes:

1. Fork the repository.
2. Create a feature branch (`feature/my-change`).
3. Open a pull request with a clear description and screenshots if applicable.

Suggested improvements:

* Add CSV dialect detection (delimiter, quoting, newline).
* Undo/Redo history for safer editing.
* Column-type detection (numbers, dates) for smarter filtering.

---

## 📝 Changelog

**v5.0**

* Added per-column filters with numeric comparison operators.
* Save As now exports filtered data.
* Improved inline editing behavior.

---

## 📄 License

This project is released under the **Secure License** — Mail me for project.

---

## Contact

Created by **Sagar Hodar** — thanks for using CSV EDITOR PRO!
COntact: hodarsagar@gmail.com
