# Excel Diff

**Excel Diff** is a lightweight Python tool that uses **pandas** to compare two Excel workbooks sheet-by-sheet and export the differences into a new Excel file.  

It is designed for structured configuration/KPI dictionaries or any tabular Excel data where you need to track changes between versions.

---

## Features

- Sheet-by-sheet comparison: matches sheets across two workbooks  
- Primary key support: match rows by one or multiple index columns  
- Two modes of diff  
  - Same shape → cell-level value changes (`pandas.DataFrame.compare`)  
  - Shape mismatch → row additions/deletions with indicator (`left only` / `right only`)  
- Skip doc sheets: sheets starting with `(DOC)` or other prefixes  
- Export to Excel: results written into a single workbook with per-sheet diff tabs  
- Configurable index map via YAML file  
