(in-package #:cl-excel)

;;; 1. File & Sheet Shortcuts

(defun read-excel (source)
  "Alias for READ-XLSX."
  (read-xlsx source))

(defun save-excel (workbook path &key overwrite)
  "Alias for WRITE-XLSX."
  (declare (ignore overwrite)) ;; write-xlsx generally overwrites in our impl
  (write-xlsx workbook path))

(defun sheet-of (workbook index-or-name)
  "Alias for SHEET."
  (sheet workbook index-or-name))

;;; 2. Cell Access

(defun val (sheet ref)
  "Get value at REF (e.g. 'A1'). Alias for GET-DATA."
  (get-data sheet ref))

(defun (setf val) (new-value sheet ref)
  "Set value at REF."
  (setf (get-cell sheet ref) (make-cell new-value)))

(defun [] (sheet ref)
  "Alias for VAL."
  (val sheet ref))


(defun (setf []) (new-value sheet ref)
  "Alias for (SETF VAL)."
  (setf (val sheet ref) new-value))

(defun cell (sheet ref)
  "Alias for VAL."
  (val sheet ref))

(defun (setf cell) (new-value sheet ref)
  "Alias for (SETF VAL)."
  (setf (val sheet ref) new-value))

;;; 3. Iteration Sugar

(defun map-rows (function sheet)
  "Apply FUNCTION to each row in SHEET.
   Returns a list of results."
  (let ((results '()))
    (do-rows (row sheet)
      (push (funcall function row) results))
    (nreverse results)))

;;; 4. Context Macros

(defmacro with-sheet ((var workbook name-or-index) &body body)
  "Execute BODY with VAR bound to the specified sheet."
  `(let ((,var (sheet ,workbook ,name-or-index)))
     ,@body))

;;; 5. Helper Functions (M11)

(defun list-sheets (source)
  "Return a list of sheet names from SOURCE (path or workbook)."
  (if (typep source 'workbook)
      (sheet-names source)
      (with-xlsx (wb source)
        (sheet-names wb))))

(defun used-range (sheet)
  "Calculate the used range of SHEET (bounding box of all non-missing cells).
   Returns a RANGE-REF structure or NIL if empty."
  (let ((min-r most-positive-fixnum)
        (max-r 0)
        (min-c most-positive-fixnum)
        (max-c 0)
        (count 0))
    (maphash (lambda (coord cell)
               (declare (ignore cell))
               (let ((r (car coord))
                     (c (cdr coord)))
                 (setf min-r (min min-r r))
                 (setf max-r (max max-r r))
                 (setf min-c (min min-c c))
                 (setf max-c (max max-c c))
                 (incf count)))
             (sheet-cells sheet))
    (if (zerop count)
        nil
        (make-range-ref (make-cell-ref min-r min-c)
                        (make-cell-ref max-r max-c)
                        :sheet (sheet-name sheet)))))

;;; 6. Smart Readers

(defun resolve-smart-range (sheet desig)
  "Resolve a range locator DESIG into a RANGE-REF with smart expansion.
   - :ALL -> returns :ALL
   - 'A1:B2' -> parsed range
   - 'A1' -> parsed single cell range A1:A1
   - 'A' -> Column A, trimmed to used-range height
   - 1 -> Column 1 (A), trimmed to used-range height"
  (cond
    ((eq desig :all) :all)
    ((typep desig 'range-ref) desig)
    ((integerp desig) 
     ;; Column index. Trim to used height.
     (let ((ur (used-range sheet)))
       (unless ur (return-from resolve-smart-range nil))
       (make-range-ref (make-cell-ref (range-start-row ur) desig)
                       (make-cell-ref (range-end-row ur) desig))))
    ((stringp desig)
     ;; Check if it's a column name only ("A", "AB")
     (if (every #'alpha-char-p desig)
         ;; Treated as column name
         (let ((idx (col-index desig))
               (ur (used-range sheet)))
           (unless ur (return-from resolve-smart-range nil))
           (make-range-ref (make-cell-ref (range-start-row ur) idx)
                           (make-cell-ref (range-end-row ur) idx)))
         ;; Else parse as ref
         (if (find #\: desig)
             (parse-cell-ref desig :range t)
             ;; Single cell "A1" -> range "A1:A1"
             (let ((cr (parse-cell-ref desig)))
               (make-range-ref cr cr)))))
    (t (error "Unknown range designator: ~A" desig))))

(defun read-file (path &optional (sheet-id 1) (range :all))
  "Quickly read data from a file.
   Returns list of lists (rows).
   PATH: Path to .xlsx file.
   SHEET-ID: Sheet name or index (default 1).
   RANGE: Smart range selector.
     - :all (default) - read all used rows
     - 'A1:B2' - specific range
     - 'A' - column A (trimmed)
     - 1 - column 1 (trimmed)
     - 'A1' - single cell"
  (with-xlsx (wb path)
    (let ((sh (sheet wb sheet-id)))
      (let ((resolved (resolve-smart-range sh range)))
        (cond
          ((eq resolved :all)
           (map-rows #'identity sh))
          ((null resolved) 
           ;; Empty sheet or range
           nil)
          ((typep resolved 'range-ref)
           ;; Iterate range
           (let ((r1 (range-start-row resolved))
                 (c1 (range-start-col resolved))
                 (r2 (range-end-row resolved))
                 (c2 (range-end-col resolved))
                 (res '()))
             (loop for r from r1 to r2 do
               (let ((row '()))
                 (loop for c from c1 to c2 do
                   (push (get-data sh (cons r c)) row))
                 (push (nreverse row) res)))
             (nreverse res)))
          (t (error "Unable to resolve range"))))))) 

;;; 7. Demo Utilities (sugar for examples)

(defun example-path (filename)
  "Return the full path to a file in tests/fixtures/.
   Example: (example-path \"basic_types.xlsx\")"
  (asdf:system-relative-pathname :cl-excel 
                                 (merge-pathnames filename "tests/fixtures/")))

(defun list-examples ()
  "List available example files in tests/fixtures/."
  (let ((path (asdf:system-relative-pathname :cl-excel "tests/fixtures/")))
    (when (probe-file path)
      (directory (merge-pathnames "*.xlsx" path)))))
