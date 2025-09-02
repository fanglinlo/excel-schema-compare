#!/usr/bin/env python3
"""
Excel DB Diff - Compare two Excel workbooks sheet-by-sheet and export differences.
- Same shape: cell-level value differences using pandas.DataFrame.compare
- Shape mismatch: row additions/deletions via outer merge with indicator
"""

import argparse
from datetime import date
from pathlib import Path
import pandas as pd
import yaml


def load_index_map(path: Path) -> dict:
    """Load YAML mapping: sheet name -> primary key column(s)."""
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    # normalize keys to original and lower-case for lookup flexibility
    lowered = {str(k).lower(): v for k, v in data.items()}
    return {"_raw": data, "_lower": lowered}


def _is_missing(col) -> bool:
    return col is None or (isinstance(col, str) and col.strip() == "")


def _drop_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    return df.loc[:, ~df.columns.str.contains(r"^Unnamed")]


def compare_excels(
    file1: Path,
    file2: Path,
    name1: str,
    name2: str,
    skiprows: int,
    doc_prefix: str,
    index_map: dict,
    outdir: Path,
) -> Path:
    """Core comparison routine. Returns the output Excel path."""
    xls1, xls2 = pd.ExcelFile(file1), pd.ExcelFile(file2)

    def filter_sheets(xls):
        return {s for s in xls.sheet_names if not (doc_prefix and s.startswith(doc_prefix))}

    s1, s2 = filter_sheets(xls1), filter_sheets(xls2)
    common = sorted(s1 & s2)
    only1, only2 = sorted(s1 - s2), sorted(s2 - s1)

    outdir.mkdir(parents=True, exist_ok=True)  # 自動建立 output 資料夾
    outname = f"{date.today():%y%m%d}_{name1}_{name2}_db_differences.xlsx"
    outpath = outdir / outname

    # Print sheet presence differences
    if only1:
        print("Sheets only in file1:", only1)
    if only2:
        print("Sheets only in file2:", only2)

    # helpers for index lookup
    raw_map = index_map.get("_raw", {})
    lower_map = index_map.get("_lower", {})

    with pd.ExcelWriter(outpath, engine="xlsxwriter") as writer:
        for sheet in common:
            print(f"Comparing: {sheet}")
            try:
                df1 = pd.read_excel(xls1, sheet_name=sheet, skiprows=skiprows)
                df2 = pd.read_excel(xls2, sheet_name=sheet, skiprows=skiprows)
            except Exception as e:
                print(f"  ! read error: {e}")
                continue

            df1, df2 = _drop_unnamed(df1), _drop_unnamed(df2)

            # find index columns by sheet name (original or lower-case key)

            index_col = raw_map.get(sheet)
            if index_col is None:
                index_col = lower_map.get(sheet.lower())

            # set index if provided, else keep default RangeIndex
            if index_col:
                missing_cols = []
                if isinstance(index_col, list):
                    for col in index_col:
                        c = str(col)
                        if c not in df1.columns or c not in df2.columns:
                            missing_cols.append(col)
                else:
                    c = str(index_col)
                    if c not in df1.columns or c not in df2.columns:
                        missing_cols.append(index_col)

                if missing_cols:
                    print(f"  ! skip (missing index cols): {missing_cols}")
                    continue

                df1 = df1.set_index(index_col)
                df2 = df2.set_index(index_col)

            # normalize empties to avoid false diffs
            df1 = df1.fillna("")
            df2 = df2.fillna("")

            # Same shape -> cell-level difference
            if df1.columns.equals(df2.columns) and df1.index.equals(df2.index):
                diff = df1.compare(df2, result_names=(name1, name2), keep_shape=False)
                if diff.empty:
                    print("  = no differences (same shape)")
                    continue
                diff = diff.reset_index()
                if isinstance(diff.columns, pd.MultiIndex):
                    diff.columns = ['_'.join([str(x) for x in c if str(x) != ""]) for c in diff.columns]
                sheet_out = f"{sheet}_same"
                diff.to_excel(writer, sheet_name=sheet_out, index=False)
                print("  -> wrote sheet:", sheet_out)
            else:
                # Shape mismatch -> list added/removed rows
                df1r, df2r = df1.reset_index(), df2.reset_index()

                # for outer merge, need 'on' keys. If no index key provided, merge on intersection of columns
                on_keys = index_col if index_col else sorted(set(df1r.columns) & set(df2r.columns))
                try:
                    merged = pd.merge(df1r, df2r, on=on_keys, how="outer", indicator=True)
                except Exception as e:
                    print(f"  ! merge error: {e}")
                    continue

                merged = merged[merged["_merge"] != "both"]
                if merged.empty:
                    print("  = no differences (shape mismatch)")
                    continue
                merged["Different"] = merged["_merge"].map({
                    "left_only": f"{name1} only",
                    "right_only": f"{name2} only"
                })
                merged.drop(columns=["_merge"], inplace=True)
                sheet_out = f"{sheet}_diff"
                merged.to_excel(writer, sheet_name=sheet_out, index=False)
                print("  -> wrote sheet:", sheet_out)

    print(f"Done. Output: {outpath}")
    return outpath


def main():
    ap = argparse.ArgumentParser(description="Compare two Excel workbooks and export differences.")
    ap.add_argument("--file1", required=True, type=Path, help="First Excel workbook (older)")
    ap.add_argument("--file2", required=True, type=Path, help="Second Excel workbook (newer)")
    ap.add_argument("--name1", required=True, help="Short label for file1 (e.g., v1.0)")
    ap.add_argument("--name2", required=True, help="Short label for file2 (e.g., v1.1)")
    ap.add_argument("--skiprows", type=int, default=3, help="Rows to skip before header (default: 3)")
    ap.add_argument("--doc-prefix", default="(DOC)", help="Prefix for sheets to exclude")
    ap.add_argument("--index-map", type=Path, required=True, help="YAML mapping of sheet -> key columns")
    ap.add_argument("--outdir", type=Path, default=Path("./"), help="Output directory")
    args = ap.parse_args()

    idx_map = load_index_map(args.index_map)
    out = compare_excels(
        file1=args.file1,
        file2=args.file2,
        name1=args.name1,
        name2=args.name2,
        skiprows=args.skiprows,
        doc_prefix=args.doc_prefix,
        index_map=idx_map,
        outdir=args.outdir,
    )
    print(out)


if __name__ == "__main__":
    main()
