# cl-excel

**cl-excel** is a modern and powerful Common Lisp library for reading and writing Microsoft Excel `.xlsx` and LibreOffice `.ods` files. 

It provides a unified "Sugar" API inspired by Julia's `XLSX.jl` and Python's `openpyxl`, allowing developers to handle complex spreadsheets with minimal code while maintaining memory efficiency for large datasets.

## Why cl-excel?

- **Successor to cl-xlsx**: This library is a significant extension of `gwangjinkim/cl-xlsx` (formerly `a1b10/cl-xlsx`). While the original was primarily a **read-only** tool, **cl-excel** adds full **writing** support from scratch.
- **Unified API**: One function to read them all (`read-file`), one to write them all (`write-file`).
- **Robust Format Detection**: It ignores file extensions and inspects internal ZIP markers to decide if a file is XLSX or ODS.
- **Rich Text Support**: Correctly parses cells with mixed formatting (bold, italic, colors) without data loss.
- **Memory Efficiency**: Supports both eager (DOM) and lazy (Streaming) reading.
- **Table Support**: Native support for Excel Tables (`ListObject`), allowing you to treat ranges as named entities.
- **Metadata Aware**: Detects the creating application and other document metadata.

---

## Mental Model

To work effectively with `cl-excel`, it is helpful to understand the underlying structure:

1.  **Workbook**: The entire file (an archive). It contains global data like shared strings, styles, and a list of sheets.
2.  **Sheet**: A single worksheet within the workbook. It can be accessed by its name (string) or a 1-based index (integer).
3.  **Cell**: The fundamental unit of data. Cells are accessed via coordinates (e.g., "A1") or row/column numbers.
    - **Missing Data**: Empty cells are represented by the special `+missing+` constant. Use `(missing-p val)` to check for it.

---

## Installation

**cl-excel** is best installed via Quicklisp's `local-projects`:

```bash
cd ~/quicklisp/local-projects/
git clone https://github.com/gwangjinkim/cl-excel.git
```

Then in your Common Lisp environment:

```lisp
(ql:quickload :cl-excel)
```

**Dependencies:**
- `cxml` / `cxml-stp` (XML parsing)
- `zip` (Archive handling)
- `local-time` (Date/Time handling)
- `alexandria` / `uiop` (Utilities)

---

## Quickstart

### 0. Setup
```lisp
(use-package :cl-excel)

;; Get an absolute path to an included example
(defparameter *xlsx* (example-path "test_table.xlsx"))

;; List sheet names
(list-sheets *xlsx*) ;; => ("Sheet1")
```

### 1. One-Liner Reading
If you just want the data as a list of lists:

```lisp
;; Read the first sheet
(read-file *xlsx*) 
;; => (("Name" "Age") ("Alice" 30) ("Bob" 25))

;; Read a specific sheet and range
(read-file *xlsx* "Sheet1" "A1:B2")
```

### 2. Simple Writing
```lisp
(let ((data '(("Item" "Count") ("Lisp" 100) ("Excel" 5))))
  (write-file data "output.xlsx")) ;; Auto-detects format from extension
```

---

## Features in Detail

### Robust Format Detection
`cl-excel` does not rely on file extensions. Even if an ODS file is named `data.xlsx`, it will be identified correctly by its internal markers.

```lisp
(detect-file-format "renamed_ods.xlsx") ;; => :ods
```

### Granular Writing with write-xlsx
The `write-xlsx` (and `write-ods`) functions provide fine-grained control:

- **:sheet**: Target a specific sheet name.
- **:start-cell**: Specify where the top-left corner of your data should be placed (e.g., "B2").
- **:region**: Define a bounding box (e.g., "A1:C10").
- **:region-only-p**: If T, data will be clipped to fit exactly within the specified region.

```lisp
(write-xlsx my-data "report.xlsx" :sheet "Sales" :start-cell "B2")
```

### Edit Mode (:rw)
You can open an existing file, modify it, and save it. 
Note: While functional, this is a best-effort regeneration. Complex charts or macros may not be preserved.

```lisp
(with-xlsx (wb "data.xlsx" :mode :rw)
  (let ((sh (sheet wb 1)))
    (setf (c sh "A1") "New Value")
    (save-excel wb "modified.xlsx")))
```

### Excel Tables (ListObjects)
Tables are named ranges in Excel that can grow and shrink. `cl-excel` can read them directly.

```lisp
(with-xlsx (wb *xlsx*)
  (let ((sh (sheet wb 1)))      ;; index is 1-based
    (read-table sh "MyTable"))) ;; Returns list of lists excluding headers
```

### Streaming Iteration (Lazy Loading)
For very large files (e.g., millions of rows), avoid loading everything into memory.

```lisp
(with-xlsx (wb "massive_data.xlsx")
  (with-sheet-iterator (next-row wb "HugeSheet")
    (loop for row = (funcall next-row)
          while row
          do (process-row row))))
```

### Extension Protocol: as-tabular
You can make any custom object compatible with `write-file` by implementing the `as-tabular` method. This is how `cl-tibble` integration works.

```lisp
(defmethod as-tabular ((obj my-custom-type))
  ;; return a list of lists
  ...)
```

---

## API Reference

### File Operations
- `read-file (path &optional sheet range)`: High-level reader.
- `write-file (source path)`: High-level writer.
- `read-xlsx (path)` / `save-excel (wb path)`: Lower-level workbook operations.
- `read-ods (path)` / `write-ods (data path)`: ODS-specific operations.

### Workbook & Sheet Access
- `sheet (workbook designator)`: Get sheet by name or 1-based index.
- `list-sheets (path or workbook)`: Return list of sheet names.
- `sheet-names (workbook)`: Return list of names in an open workbook.

### Cell Access (Sugar)
All of the following can be used with `setf` in `:rw` mode:
- `c (sheet ref)`: General cell value accessor (ref can be "A1").
- `val (sheet ref)`: Alias for `c`.
- `[] (sheet ref)`: Alias for `c`.
- `get-data (sheet ref)`: The underlying protocol function.

### Iterators
- `map-rows (fn sheet)`: Apply `fn` to every row.
- `do-rows ((row-var sheet) &body)`: Logical iteration over rows.
- `each-table-row (fn table)`: Apply `fn` to table data.

---

## Support & Parity

`cl-excel` aims for high reliability, but it is currently in **Beta**. We strive for feature parity with Julia's `XLSX.jl` where it makes sense for the Common Lisp ecosystem.

| Feature | Status |
| :--- | :--- |
| **XLSX Reading** | ✅ Supported |
| **XLSX Writing** | ✅ Supported |
| **ODS Reading** | ✅ Supported |
| **ODS Writing** | ✅ Supported |
| **Tables** | ✅ Supported |
| **Rich Text** | ✅ Supported |
| **Streaming** | ✅ Supported |
| **Styles** | ⚠️ Basic (Formats only) |
| **Formulas** | ❌ Not supported (Read as values) |

---

## Testing

Run the test suite via the provided `Makefile`:

```bash
make test
```

Standard tests include over 150 checks covering roundtrips, edge cases, and cross-format detection.

## License

MIT License.
