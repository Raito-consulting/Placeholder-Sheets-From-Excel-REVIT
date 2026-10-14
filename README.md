# Revit Placeholder Sheets

A pyRevit button that creates **placeholder sheets in bulk** from an Excel or CSV list.

## The problem

A US-style sheet index lists every sheet in the set, including ones that don't exist in
the model: structural, MEP, civil and consultant sheets. In Revit those are *placeholder
sheets*, and the built-in way to make them is **Insert Data Row** in the sheet list, one
row at a time, then filling in each parameter by hand. For a set with 30-60 consultant
sheets that is slow and easy to get wrong.

## The solution

Keep the sheet list in a spreadsheet, point the button at the file, and every placeholder
is created with its number, name and parameters in one undoable step.

```
Category,Discipline,Order,Suborder,Sheet #,Sheet Name
Placeholder,ELECTRICAL,10,0,E0.0,ELECTRICAL LEGEND & NOTES
Placeholder,ELECTRICAL,10,2,E2.1,FIRST LEVEL LIGHTING PLAN
```

## Features

- Reads **.xlsx** or **.csv** (the .xlsx reader uses only the Python standard library, no
  Excel or openpyxl needed).
- Only `Sheet #` and `Sheet Name` are required.
- **Forgiving headers**: `Order`, `Discipline Order`, `Sub order` and similar map to the
  right fields; the real Revit parameter names are set in one dictionary (`PARAM_NAMES`).
- **Add columns without changing code**: any other column header is written to the sheet
  parameter of the same name. Text, integer and number parameters are converted for you;
  Yes/No (checkbox) parameters accept `X`, `Yes`, `Y` or `1`.
- **Safe to re-run**: sheet numbers that already exist are skipped, never overwritten.
- Reports what was created, what was skipped and which parameters were not found.

## Install

1. Install [pyRevit](https://github.com/pyrevitlabs/pyRevit).
2. Copy `PlaceholderSheets.extension` somewhere permanent (keep the `.extension` suffix).
3. pyRevit tab -> Settings -> Extensions -> add that folder, then **Reload**.
4. Use **Sheets -> Placeholders** on the new tab.

Tested on Revit with pyRevit's IronPython engine.

## Usage

1. Fill in a list like [`examples/sheet_list.csv`](examples/sheet_list.csv).
2. Click **Placeholders** and choose the file.
3. Check the report in the output window.

## How it works

- `ViewSheet.CreatePlaceholder` creates each sheet inside a single transaction.
- Header text is normalised and matched to the Revit field it means; unknown headers fall
  through to a same-name parameter lookup.
- Parameter values are set according to the parameter's `StorageType` (string, integer,
  double), with special handling for Yes/No integers.
- The `.xlsx` file is read as a zip of XML (shared strings + sheet data).

## Limitations

- Does not edit existing sheets, only creates new ones.
- Parameter names are case-sensitive and must exist on the Sheets category in the project.
- Not tested on workshared models with sheets checked out by other users.

## License

MIT
