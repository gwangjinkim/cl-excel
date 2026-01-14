# cl-excel

**cl-excel** is a Common Lisp library for reading and writing Microsoft Excel 2007+ (`.xlsx`) files. Ideally, it aims for feature parity with the Julia `XLSX.jl` package, providing a convenient and unified API for data analysis and reporting.

| Feature | Status |
| :--- | :--- |
| **Reading** | ✅ Supported (Eager & Lazy) |
| **Writing** | ✅ Supported (New & Edit Mode) |
| **Tables** | ✅ Read & Write |
| **Streaming** | ✅ Memory-efficient Iterators |
| **Styles** | ⚠️ Basic (Dates/Formats only) |

Another library for excel files which I generated in the past is: https://github.com/a1b10/cl-xlsx
But this library now can not only read but also write to excel files.
And I generated this library (cl-excel) using AI.

## Installation

**cl-excel** is available via ASDF/Quicklisp (local).

```lisp
(asdf:load-system :cl-excel)
```

If you are using Quicklisp, got clone into your `local-projects` folder:

```bash
cd ~/quicklisp/local-projects/
# if using roswell: `cd ~/.roswell/local-projects/`
git clone git@github.com:gwangjinkim/cl-excel.git

# start your lisp by:
sbcl
# if using roswell: `ros run`
```

And then in your lisp:

```lisp
(ql:quickload :cl-excel)
```

Alternatively, you can git clone into another folder and add the folder your `ql:*local-project-directories*`:

```bash
cd ~/your/other-folder/
git clone git@github.com:gwangjinkim/cl-excel.git 
```

```lisp
(push #p"~/your/other-folder/" ql:*local-project-directories*)
(ql:quickload :cl-excel)
```

You can make the adding of `other-folder` to your `ql:*local-project-directories*` permanent by adding the line
to your `~/.sbclrc` or if your are using Roswell, to your `~/.roswell/init.lisp`:

```
;; ~/.roswell/init.lisp
(handler-case
    (progn
      (require :asdf)
      (when (find-package :ql)
        (pushnew (merge-pathnames "your/other-folder/" (user-homedir-pathname)) ;; <= modify this folder path!
                 ql:*local-project-directories*
                 :test #'equal)
        (ql:register-local-projects)))
  (error () nil))
```

Or simply symlink your `other-folder/` into your `local-projects` folder.

---

## Quickstart (The Natural Way)

The "Sugar" API is designed to be concise and intuitive, similar to Python's pandas or openpyxl.

### 0. Do This First

```lisp
(ql:quickload :cl-excel)

(use-package :cl-excel)   ;; to get rid of `cl-excel:`


;; List available example Excel files in this package
(list-examples)

;; Get full absolute path from an example file
(example-path "test_table.xlsx")

;; Define an example path
(defparameter *xlsx* (cl-excel:example-path "test_table.xlsx"))

;; List all sheet names of an Excel file
(list-sheets *xlsx*) ;; => ("Sheet1")

```

### 1. One-Liner Read

Just want the data?

```lisp
;; Read the first sheet as a list of lists
(read-file *xlsx*) 
;; => (("Name" "Age") ("Alice" 30) ("Bob" 25))

;; Read a specific sheet and range
(read-file *xlsx* "Sheet1" "A1:B10") ;; Not auto-trimmed! Empty cells are `#<MISSING>`

;; You can use index number (1-based) instead of the sheet name
(read-file *xlsx* 1 "A1:B10")



;; Smart Ranges (M11)
(cl-excel:read-file *xlsx* 1 "B")  ;; Read Column A (auto-trimmed)
(cl-excel:read-file *xlsx* 1 2)    ;; Read Row 2 (auto-trimmed)
(cl-excel:read-file *xlsx* 1 "B2") ;; Read Single Cell "A1"
(cl-excel:read-file *xlsx* 1 "A:C") ;; Read Column A to C (auto-trimmed in width)
(cl-excel:read-file *xlsx* 1 "1:5") ;; Read Rows 1 to 5 (auto-trimmed in width)
```

### 2. Concise Access

```lisp
;; List sheets without opening explicitly
(print (cl-excel:list-sheets *xlsx*))

(with-xlsx (wb *xlsx* :mode :rw)
  (with-sheet (s wb 1)
    
    ;; Get Value using `[]` or `val` or `cell`
    (print (cell s "A1")) 
    
    ;; Set Value `[]` or `val` or `cell`
    (setf (cell s "B1") "Updated")
    
    ;; Iterate rows
    (map-rows (lambda (row) (print row)) s)
    
    ;; Save changes
    (save-excel wb "saved.xlsx")))
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
(let ((wb (cl-excel:read-excel *xlsx*)))
  (let ((sh (cl-excel:sheet-of wb "Sheet1")))
     (format t "Cell A1 is: ~A~%" (cl-excel:cell sh "A1"))))
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
;; To access the included test files reliably:
(cl-excel:with-xlsx (wb (cl-excel:example-path "sugar.xlsx"))
  
  ;; Iterate over "Sheet1" or sheet index 1
  (cl-excel:with-sheet-iterator (next-row wb 1)  ;; 1 is sheet number 1-based
    (loop for row = (funcall next-row)
          while row
          do (format t "Processing Row: ~A~%" row))))
```

### API Reference

- `read-file (path &optional sheet range)`: High-level reader.
- `read-excel (path)` / `save-excel (wb path)`: File I/O aliases.
- `val (sheet ref)` / `[] (sheet ref)` / `cell (sheet ref)`: Cell accessors.
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
- Charts and Drawings support.
- Formulas are not in my interest. The idea is rather Excel as input/output format
  for humans (Scientists etc. LOVE Excel as a document format).

## License

GPL-3.0 License.
