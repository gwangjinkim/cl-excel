# cl-excel

**cl-excel** is a Common Lisp library for reading and writing Microsoft Excel 2007+ (`.xlsx`) files. Ideally, it aims for feature parity with the Julia `XLSX.jl` package, providing a convenient and unified API for data analysis and reporting.

| Feature | Status |
| :--- | :--- |
| **Reading** | ✅ Supported (Eager & Lazy) |
| **Writing** | ✅ Supported (New & Edit Mode) |
| **Tables** | ✅ Read & Write |
| **Streaming** | ✅ Memory-efficient Iterators |
| **Styles** | ⚠️ Basic (Dates/Formats only) |

## Installation

**cl-excel** is available via ASDF/Quicklisp (local).

```lisp
(asdf:load-system :cl-excel)
```

## Quickstart (The Natural Way)

The "Sugar" API is designed to be concise and intuitive, similar to Python's pandas or openpyxl.

### 1. One-Liner Read

Just want the data?

```lisp
;; Read the first sheet as a list of lists
(cl-excel:read-file "data.xlsx") 
;; => (("Name" "Age") ("Alice" 30) ("Bob" 25))

;; Read a specific sheet and range
(cl-excel:read-file "data.xlsx" "Sheet1" "A1:B10")
```

### 2. Concise Access

```lisp
(cl-excel:with-xlsx (wb "data.xlsx" :mode :rw)
  (cl-excel:with-sheet (s wb 1)
    
    ;; Get Value using [] or val
    (print (cl-excel:[] s "A1")) 
    
    ;; Set Value
    (setf (cl-excel:[] s "B1") "Updated")
    
    ;; Iterate rows
    (cl-excel:map-rows (lambda (row) (print row)) s)
    
    ;; Save changes
    (cl-excel:save-excel wb "saved.xlsx")))
```

---

## Julia `XLSX.jl` Style

If you are coming from Julia, here is how `cl-excel` compares.

| Operation | Julia (`XLSX.jl`) | Common Lisp (`cl-excel`) |
| :--- | :--- | :--- |
| **Open File** | `xf = XLSX.readxlsx("f.xlsx")` | `(defparameter *xf* (cl-excel:read-excel "f.xlsx"))` |
| **Get Sheet** | `sh = xf["Sheet1"]` | `(defparameter *sh* (cl-excel:sheet-of *xf* "Sheet1"))` |
| **Get Cell** | `val = sh["A1"]` | `(cl-excel:[] *sh* "A1")` |
| **Set Cell** | `sh["A1"] = "foo"` | `(setf (cl-excel:[] *sh* "A1") "foo")` |
| **Write Table** | `XLSX.writetable("f.xlsx", data)` | `(cl-excel:add-table! *sh* data)` (then save) |

**Example Port:**

```lisp
(let ((xf (cl-excel:read-excel "data.xlsx")))
  (let ((sh (cl-excel:sheet-of xf "Sheet1")))
     (format t "Cell A1 is: ~A~%" (cl-excel:[] sh "A1"))))
```

---

## Detailed Documentation

### Core Concepts

- **Workbook**: The main container (`cl-excel:workbook`).
- **Sheet**: A single worksheet (`cl-excel:sheet`).
- **Cell**: The fundamental unit. Empty cells are `+missing+`.

### Advanced Features

#### Edit Mode
Open files with `:mode :rw` to modify them.
> [!WARNING]
> This is a **best-effort** regeneration. Charts and images may be lost.

#### Excel Tables (`ListObject`)
Native support for reading and creating Excel Tables.
```lisp
(cl-excel:add-table! sheet my-data :name "SalesTable")
```

#### Lazy Reading (Streaming)
For large files, use iterators to keep memory usage low.
```lisp
(cl-excel:with-sheet-iterator (iter "large.xlsx" 1)
  (loop
    (let ((row (funcall iter)))
       (unless row (return))
       (print row))))
```

### API Reference

- `read-file (path &optional sheet range)`: High-level reader.
- `read-excel (path)` / `save-excel (wb path)`: File I/O aliases.
- `val (sheet ref)` / `[] (sheet ref)`: Cell accessors.
- `map-rows (fn sheet)`: Functional iteration.
- `wb`, `sheet`, `cell`: Low-level constructors if needed.

---

## Status & Roadmap

This library is currently in **Beta** (M10 complete).

- [x] **M0-M8**: Core Reading/Writing/Tables
- [x] **M9**: Edit Mode (:rw)
- [x] **M10**: Sugar API (User Friendliness)

**Future Plans:**
- Better style preservation in Edit Mode.
- Support for Formula calculation (currently reads cached values).
- Charts and Drawings support.

## License

MIT License.
