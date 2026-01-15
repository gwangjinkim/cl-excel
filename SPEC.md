# Specification: Enhanced write-xlsx

## Overview
The `write-xlsx` function is enhanced to support targeted writing of tabular data into specific sheets and regions within a workbook.

## Function Signature
```lisp
(defun write-xlsx (source path &key sheet start-cell region region-only-p overwrite-p)
  ...)
```

### Arguments

- **`source`**: 
    - A `workbook` object: The workbook is saved to `path`.
    - Tabular data (list of lists): The data is written to a sheet in a workbook and saved to `path`.
- **`path`**: The file path to write to.
- **`sheet`** (Optional):
    - A string: The name of the sheet.
    - An integer: The 1-based index of the sheet.
    - If `source` is tabular data and `path` exists, the data is written into this sheet. If it doesn't exist, a new sheet is created.
- **`start-cell`** (Optional):
    - A string (e.g., "A1"): The cell where the top-left of the tabular data should be placed.
- **`region`** (Optional):
    - A string (e.g., "A1:D10"): A bounding box for writing.
- **`region-only-p`** (Optional, default `nil`):
    - `t`: Only data that fits within the `region` is written. If no `region` is given, this has no effect.
    - `nil`: The `region`'s first cell (top-left) is used as the `start-cell`. Data may be written outside the region if it's larger.
- **`overwrite-p`** (Optional, default `t`):
    - `t`: If `path` exists, it is completely overwritten (unless `:mode :rw` logic is integrated).
    - `nil`: If `path` exists, it is opened in `:rw` (edit) mode, updated, and saved.

## Behavior Matrices

### Region Logic
| Argument | data fits? | `:region-only-p t` | `:region-only-p nil` |
| :--- | :--- | :--- | :--- |
| `region` | Yes | Writes data. | Writes data. |
| `region` | No (too big) | Clips data to fit region. | Writes full data from start of region. |

### Start Cell Logic
If both `region` and `start-cell` are provided, `region` takes precedence for the start position.

### Sheet Logic
- If `sheet` is a number `n`, it refers to the `n`-th sheet (1-based).
- If `sheet` is a string, it refers to the sheet by name.
- If it's a new workbook, the first sheet is always index 1.
