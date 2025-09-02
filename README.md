# Excel DB Diff

**Excel DB Diff** is a lightweight Python tool that uses **pandas** to compare two Excel workbooks sheet-by-sheet and export the differences into a new Excel file.

It is designed for structured configuration/KPI dictionaries or any tabular Excel data where you need to track changes between versions.

---

## Features

- Sheet-by-sheet comparison across two workbooks  
- Primary key support: match rows by one or multiple index columns  
- Two modes of diff  
  - **Same shape** → cell-level value differences (`pandas.DataFrame.compare`)  
  - **Shape mismatch** → row additions/deletions with indicator (`left only` / `right only`)  
- Skip doc sheets: sheets starting with `(DOC)` or other prefixes  
- Export to Excel: results written into a single workbook with per-sheet diff tabs  
- Configurable index map via YAML file  

---

## Install

Clone the repo and install dependencies:

```bash
git clone https://github.com/fanglinlo/excel-schema-compare.git
cd excel-schema-compare

python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```
---

## Usage
```
python main.py \
  --file1 ./samples/Book_A.xlsx \
  --file2 ./samples/Book_B.xlsx \
  --name1 BookA \
  --name2 BookB \
  --skiprows 0 \
  --doc-prefix "(DOC)" \
  --index-map ./config/index_map_demo.yaml \
  --outdir ./output
```

| Argument       | Required | Description                                         |
| -------------- | -------- | --------------------------------------------------- |
| `--file1`      | Yes      | Path to the first Excel file      |
| `--file2`      | Yes      | Path to the second Excel file       |
| `--name1`      | Yes      | Short label for file1 (used in headers & filename)  |
| `--name2`      | Yes      | Short label for file2                               |
| `--skiprows`   | No       | Rows to skip before header (default: 3)             |
| `--doc-prefix` | No       | Prefix for sheets to exclude (default: `(DOC)`)     |
| `--index-map`  | Yes      | Path to YAML config mapping sheets → primary key(s) |
| `--outdir`     | No       | Output directory (default: `./`)                    |


---

## Index Map Config

Define primary key columns per sheet in a YAML file.
Supports both single and composite keys:

# Example demo sheets
User_List: user_id
Product_Table: [product_id, version]
Transaction_Log: [txn_id, date]

---
## Notes

- The program will create the output/ directory automatically if it does not exist.
- This tool is not intended to handle unstructured Excel files.
- For demo purposes, include only synthetic/sample Excel files in your repo.
