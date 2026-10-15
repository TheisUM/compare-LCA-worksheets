# Compare LCA worksheets

Compares two versions of an LCA quantity take-off (*mængdeopgørelse*) in Excel and
highlights what has changed: new lines, changed values in existing lines, and lines
that have been removed.

## Setup

Requires Python 3.9 or newer. Create a virtual environment and install the requirements once:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Usage

1. Put the **old** Excel file in `old_LCA_sheet/` and the **new** one in `new_LCA_sheet/`.
   Each folder must contain exactly one Excel file.
2. Find the sheet names to compare:

   ```powershell
   .\.venv\Scripts\python.exe compare_lca.py --list-sheets
   ```

3. Run the comparison with the sheet to read from each file:

   ```powershell
   .\.venv\Scripts\python.exe compare_lca.py --old-sheet ARK --new-sheet Ark1
   ```

   If a workbook has only one sheet, you can leave out its sheet option.

The result is saved in the project folder as `<new file name> - comparison.xlsx`.

### Options

| Option | Default | Description |
|---|---|---|
| `OLD NEW` (positional) | file in `old_LCA_sheet/` and `new_LCA_sheet/` | Compare two specific files instead |
| `--old-sheet`, `--new-sheet` | the only sheet | Sheet to read in each file |
| `-o`, `--output` | `<new file name> - comparison.xlsx` | Name of the result file |
| `--header-row` | `10` | Row containing the column headers |
| `--start-row` | `11` | First row of data |
| `--tolerance` | `1e-9` | Relative tolerance when comparing numbers |
| `--list-sheets` | | Only list the sheets in the files |

## Expected sheet layout

Both sheets must have the same column layout:

| Column | Content |
|---|---|
| A | TypeCode |
| B | Part name (*Bygningsdelsnavn*) |
| C | Extra description (*Konstruktionsopbygningsbeskrivelse*) |
| D onwards | Quantities, units and notes |

Headers are on row 10 and data starts on row 11. Every column with a header on row 10
is compared. If the headers of the two sheets differ, a warning is printed. This
usually means the wrong sheet was chosen.

## The result file

The result is a copy of the new workbook containing:

- **The compared sheet**:
  - New lines are filled **green**.
  - Changed cells are filled **yellow**, with a cell comment showing the old value.
  - A *Comparison status* column is added to the right of the data, e.g.
    `CHANGED: Antal [stk.], Total areal [m2]`.
- **A `Comparison` sheet** with a summary and one line per new line, changed cell and
  removed line. It includes the old and new value and a link to the row in the new sheet.

## How lines are matched

TypeCodes are not always unique (placeholder codes like `253xxx` are reused for many
parts), so lines are matched in two steps:

1. On **TypeCode + part name**. If the same combination occurs several times, the
   n-th occurrence in the old sheet is paired with the n-th occurrence in the new one.
2. Lines still unmatched are paired on **TypeCode alone** if that code occurs exactly
   once in each sheet. This means the part was renamed, and the name change is shown
   as a changed cell.

Any new line still unmatched is reported as **new**. Any old line still unmatched is
reported as **removed**.

Values are compared so that blank cells and empty text count as equal, leading and
trailing spaces and blank lines are ignored, and a number stored as text (`'216001'`)
equals the same number (`216001`). Formula cells are compared by their last calculated
value.

## Notes

- The result file is written with `openpyxl`. Values, formulas and cell formatting are
  kept, but images, charts and some advanced Excel features in the new workbook may
  not carry over to the result file. The original files are never modified.
- Close the result file in Excel before running the comparison again, otherwise it
  cannot be overwritten.
