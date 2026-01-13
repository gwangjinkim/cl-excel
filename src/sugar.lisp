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

;;; 5. Quick Read

(defun read-file (path &optional (sheet-id 1) (range :all))
  "Quickly read data from a file.
   Returns list of lists (rows).
   PATH: Path to .xlsx file.
   SHEET-ID: Sheet name or index (default 1).
   RANGE: Cell range to read/scan (default :all)."
  (with-xlsx (wb path)
    (let ((sh (sheet wb sheet-id)))
      (if (eq range :all)
          (map-rows #'identity sh)
          ;; If explicit range "A1:B2"
          (let ((range-struct (if (stringp range) (parse-cell-ref range :range t) range)))
            ;; Simple iteration over the range
            (let ((r1 (range-start-row range-struct))
                  (c1 (range-start-col range-struct))
                  (r2 (range-end-row range-struct))
                  (c2 (range-end-col range-struct))
                  (res '()))
              (loop for r from r1 to r2 do
                (let ((row '()))
                  (loop for c from c1 to c2 do
                    (push (get-data sh (cons r c)) row))
                  (push (nreverse row) res)))
              (nreverse res))))))) 
