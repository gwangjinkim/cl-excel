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
;; Load system
(asdf:load-system :cl-excel)
```

## Quickstart

### 1. Reading Files

```lisp
(cl-excel:with-xlsx (wb "data.xlsx")
  ;; Access a specific sheet
  (let ((sheet (cl-excel:sheet wb 1))) ;; or by name "Sheet1"
    
    ;; Get cell value
    (format t "Cell A1: ~A~%" (cl-excel:get-data sheet "A1"))
    
    ;; Get range (returns list-of-lists or array depending on impl, currently scalar focused)
    (format t "Cell B2: ~A~%" (cl-excel:get-data sheet "B2"))))
```

### 2. Writing New Files

```lisp
(let ((wb (make-instance 'cl-excel:workbook))
      (sheet (make-instance 'cl-excel:sheet :name "Report" :id 1)))
  
  (setf (cl-excel:workbook-sheets wb) (list sheet))
  
  ;; Set values
  (setf (cl-excel:cell sheet "A1") "Name")
  (setf (cl-excel:cell sheet "B1") "Score")
  (setf (cl-excel:cell sheet "A2") "Alice")
  (setf (cl-excel:cell sheet "B2") 100)
  
  ;; Write to disk
  (cl-excel:write-xlsx wb "report.xlsx"))
```

### 3. Edit Mode (Modify Existing)

You can edit an existing file by opening it with `:mode :rw`.

> [!WARNING]
> This is a **best-effort** feature. We regenerate the XLSX structure. Advanced features (Charts, Macros, Images) from the original file **will be lost**.

```lisp
(cl-excel:with-xlsx (wb "report.xlsx" :mode :rw)
  (let ((sheet (cl-excel:sheet wb 1)))
    ;; Modify a cell
    (setf (cl-excel:cell sheet "B2") 95)
    
    ;; Save (can overwrite or new path)
    (cl-excel:write-xlsx wb "report_updated.xlsx")))
```

### 4. Working with Tables

Excel Tables (`ListObject`) are fully supported.

**Reading:**
```lisp
(cl-excel:with-xlsx (wb "sales.xlsx")
  ;; Get table by name or index
  (let ((data (cl-excel:read-table wb "SalesTable")))
    ;; Returns list of lists
    (print data)))
```

**Writing:**
```lisp
(let ((wb (make-instance 'cl-excel:workbook))
      (sheet (make-instance 'cl-excel:sheet :name "Data" :id 1)))
  (setf (cl-excel:workbook-sheets wb) (list sheet))
  
  (let ((data '(("Product" "Qty") 
                ("Apple" 10) 
                ("Banana" 20))))
    ;; Add data and register as Excel Table
    (cl-excel:add-table! sheet data :name "ProductTable"))
    
  (cl-excel:write-xlsx wb "products.xlsx"))
```

### 5. Large Files (Lazy Reading)

For large datasets, use the streaming iterator to avoid loading the entire file into memory.

```lisp
(cl-excel:with-sheet-iterator (iter "large_file.xlsx" "Sheet1")
  (loop
    (let ((row (funcall iter)))
      (unless row (return))
      (format t "Row ~D: ~A~%" (car row) (cdr row)))))
```
### 6. Sugar API (User-Friendly Interface)

For a more natural (Python-like) experience, use the "Sugar" API aliases.

```lisp
;; Quick Read (gets data as list of lists immediately)
(cl-excel:read-file "data.xlsx") 

;; Concise Workflow
(cl-excel:with-xlsx (wb "data.xlsx" :mode :rw)
  (cl-excel:with-sheet (s wb 1)
    
    ;; Val Get/Set (alias for get-data/get-cell)
    (print (cl-excel:val s "A1")) 
    
    ;; Array-like syntax (alias for val)
    (setf (cl-excel:[] s "B1") "Updated")
    
    ;; Functional Iteration
    (cl-excel:map-rows (lambda (row) (print row)) s)
    
    ;; Save (alias for write-xlsx)
    (cl-excel:save-excel wb "saved.xlsx")))
```

---

## Detailed Documentation

### Core Concepts

- **Workbook**: The main container (`cl-excel:workbook`). Holds sheets, styles, and shared strings.
- **Sheet**: A single worksheet (`cl-excel:sheet`), accessible by index (1-based) or name.
- **Cell**: The fundamental unit. `cl-excel` uses `+missing+` (sentinel) for empty cells to distinguish from `NIL` (boolean false).

### API Reference

#### Opening & Closing
- `(read-xlsx path)`: Eagerly load full workbook.
- `(open-xlsx path &key mode enable-cache)`:
    - `:mode :read` (default)
    - `:mode :write` (creates empty workbook)
    - `:mode :rw` (loads for editing)
- `(close-xlsx wb)`: Frees resources (important for ZIP handles).
- `(with-xlsx (var path ...) body)`: Macro for safe resource handling.

#### Accessing Data
- `(sheet wb ID_OR_NAME)`: Get a sheet object.
- `(get-data sheet REF)`: Get value at "A1".
- `(get-cell sheet REF)`: Get the `cell` struct (contains `.value`, `.type`).
- `(setf (cell sheet REF) VAL)`: Set value at "A1".

#### Tables
- `(get-table wb NAME)`: Get `table` struct.
- `(read-table wb NAME)`: Extract data as list-of-lists.
- `(add-table! sheet data &key name header start-cell)`: Create a table.

#### Iterators
- `(make-sheet-iterator path sheet-name)`: Returns `(values iterator-fn cleanup-fn)`.
- `(with-sheet-iterator (var path sheet-name) body)`: Safe macro wrapper.

### Type Mapping

| Excel Type | Common Lisp Type | Note |
| :--- | :--- | :--- |
| String | `STRING` | |
| Number | `NUMBER` | Integer or Float |
| Boolean | `BOOLEAN` | `T` or `NIL` |
| Empty | `+missing+` | Sentinel value |
| Date | `LOCAL-TIME:TIMESTAMP` | If `local-time` loaded |

---

## Status & Roadmap

This library is currently in **Beta** (M9 complete).
Future plans include:
- Better style preservation in Edit Mode.
- Support for Formulas (reading cached values works, writing formulas needs work).
- Charts and Drawings support.

## License

MIT License.
