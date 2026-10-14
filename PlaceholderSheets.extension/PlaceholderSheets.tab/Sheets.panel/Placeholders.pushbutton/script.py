# -*- coding: utf-8 -*-
"""
Create placeholder sheets in bulk from an Excel (.xlsx) or CSV file.

Runs on pyRevit's default IronPython engine (pyrevit.forms is IronPython-only
on some pyRevit builds). csv, zipfile and xml.etree are pure-Python stdlib.

Header row (case-insensitive):
    Category | Discipline | Order | Suborder | Sheet # | Sheet Name
Only "Sheet #" and "Sheet Name" are required.

Any column header containing "order" / "suborder" (e.g. "Discipline Order")
maps to the Order / Suborder fields. Any OTHER header is used literally as a
Revit parameter name (case-sensitive), so extra columns need no code changes.
Checkbox (Yes/No) parameters accept X / Yes / Y / 1 (checked) or No / 0.

Sheet numbers that already exist in the model are skipped.
"""

import csv
import io
import os
import re
import zipfile
import xml.etree.ElementTree as ET

from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()

# ------------------------------------------------------------------
# PARAMETER NAME MAP - edit the right-hand values if your project's
# real Revit parameter names differ (exact, case-sensitive).
# ------------------------------------------------------------------
PARAM_NAMES = {
    "category": "Category",
    "discipline": "Discipline",
    "order": "Discipline Order",
    "suborder": "Discipline Suborder",
}

HEADER_ALIASES = {
    "category": "category",
    "discipline": "discipline",
    "order": "order",
    "suborder": "suborder",
    "sub order": "suborder",
    "sheet #": "sheet_number",
    "sheet#": "sheet_number",
    "sheet number": "sheet_number",
    "sheet no": "sheet_number",
    "sheet no.": "sheet_number",
    "sheet name": "sheet_name",
}

RESERVED_KEYS = set(["sheet_number", "sheet_name", "category",
                     "discipline", "order", "suborder"])


def header_key(h):
    raw = (h or "").strip()
    if not raw:
        return None
    low = raw.lower()
    if low in HEADER_ALIASES:
        return HEADER_ALIASES[low]
    if "suborder" in low or "sub order" in low:
        return "suborder"
    if "order" in low:
        return "order"
    return raw


def read_csv(path):
    with io.open(path, "r", encoding="utf-8-sig", newline="") as f:
        raw_rows = list(csv.reader(f))
    if not raw_rows:
        return []
    header = [header_key(h) for h in raw_rows[0]]
    rows = []
    for raw in raw_rows[1:]:
        row = {}
        for idx, val in enumerate(raw):
            if idx < len(header) and header[idx]:
                row[header[idx]] = (val or "").strip()
        rows.append(row)
    return rows


def read_xlsx_minimal(path):
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            tree = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in tree.findall("m:si", ns):
                texts = si.findall(".//m:t", ns)
                shared.append("".join(t.text or "" for t in texts))

        sheet_path = None
        for name in z.namelist():
            if name == "xl/worksheets/sheet1.xml":
                sheet_path = name
                break
        if sheet_path is None:
            for name in z.namelist():
                if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                    sheet_path = name
                    break
        if sheet_path is None:
            return []

        tree = ET.fromstring(z.read(sheet_path))
        sheet_data = tree.find("m:sheetData", ns)
        if sheet_data is None:
            return []

        grid = []
        for row in sheet_data.findall("m:row", ns):
            row_vals = {}
            for c in row.findall("m:c", ns):
                ref = c.attrib.get("r", "")
                match = re.match(r"[A-Z]+", ref)
                col_letters = match.group(0) if match else "A"
                col_idx = 0
                for ch in col_letters:
                    col_idx = col_idx * 26 + (ord(ch) - ord("A") + 1)
                col_idx -= 1

                cell_type = c.attrib.get("t")
                v = c.find("m:v", ns)
                text = v.text if v is not None else ""
                if cell_type == "s" and text != "":
                    text = shared[int(text)]
                elif cell_type == "inlineStr":
                    t = c.find(".//m:t", ns)
                    text = t.text if t is not None else ""
                row_vals[col_idx] = text
            max_col = max(row_vals.keys()) if row_vals else -1
            grid.append([row_vals.get(i, "") for i in range(max_col + 1)])
        return grid


def read_xlsx(path):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        grid = [[("" if c is None else c) for c in row]
                for row in ws.iter_rows(values_only=True)]
    except ImportError:
        grid = read_xlsx_minimal(path)

    if not grid:
        return []
    header = [header_key(str(h)) for h in grid[0]]
    rows = []
    for raw in grid[1:]:
        row = {}
        for idx, val in enumerate(raw):
            if idx < len(header) and header[idx]:
                row[header[idx]] = "" if val is None else str(val).strip()
        rows.append(row)
    return rows


param_warnings = set()


def set_param(sheet, param_name, value):
    if not value:
        return
    p = sheet.LookupParameter(param_name)
    if p is None:
        param_warnings.add(param_name)
        return
    if p.IsReadOnly:
        return
    try:
        if p.StorageType == DB.StorageType.String:
            p.Set(str(value))
        elif p.StorageType == DB.StorageType.Integer:
            try:
                p.Set(int(float(value)))
            except ValueError:
                v = value.strip().lower()
                if v in ("x", "yes", "y", "true", "checked", "1"):
                    p.Set(1)
                elif v in ("no", "n", "false", "unchecked", "0"):
                    p.Set(0)
        elif p.StorageType == DB.StorageType.Double:
            p.Set(float(value))
    except Exception:
        pass


def main():
    path = forms.pick_file(
        files_filter="Excel or CSV files (*.xlsx;*.csv)|*.xlsx;*.csv|"
                     "Excel files (*.xlsx)|*.xlsx|"
                     "CSV files (*.csv)|*.csv"
    )
    if not path:
        return

    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        data_rows = read_csv(path)
    elif ext == ".xlsx":
        data_rows = read_xlsx(path)
    else:
        forms.alert("Unsupported file type: {}\nUse .xlsx or .csv".format(ext))
        return

    valid_rows = [r for r in data_rows if r.get("sheet_number") and r.get("sheet_name")]
    if not valid_rows:
        forms.alert("No valid rows found.\n\nThe file needs a header row with at "
                    "least 'Sheet #' and 'Sheet Name' columns, with data below it.")
        return

    existing_numbers = set(
        s.SheetNumber
        for s in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet).ToElements()
    )

    created, skipped_existing = [], []

    with revit.Transaction("Create Placeholder Sheets from File"):
        for row in valid_rows:
            sheet_num = row["sheet_number"].strip()
            sheet_name = row["sheet_name"].strip()

            if sheet_num in existing_numbers:
                skipped_existing.append(sheet_num)
                continue

            new_sheet = DB.ViewSheet.CreatePlaceholder(doc)
            new_sheet.SheetNumber = sheet_num
            new_sheet.Name = sheet_name

            for canon_key, real_param_name in PARAM_NAMES.items():
                set_param(new_sheet, real_param_name, row.get(canon_key))

            for key, value in row.items():
                if key in RESERVED_KEYS:
                    continue
                set_param(new_sheet, key, value)

            created.append(sheet_num)
            existing_numbers.add(sheet_num)

    output.print_md("### Placeholder Sheets")
    output.print_md("**File:** {}".format(os.path.basename(path)))
    output.print_md("**Created ({}):** {}".format(len(created), ", ".join(created) or "-"))
    if skipped_existing:
        output.print_md("**Skipped, already existed ({}):** {}".format(
            len(skipped_existing), ", ".join(skipped_existing)))
    if param_warnings:
        output.print_md("**Note:** these parameters weren't found on the Sheets "
                        "category and were left blank: {}".format(
                            ", ".join(sorted(param_warnings))))


main()
