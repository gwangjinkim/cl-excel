;;;; tests/run.lisp — FiveAM harness (M0)

(defpackage #:cl-excel.tests
  (:use #:cl #:fiveam)
  (:export #:run-tests
           #:fixture-path))

(in-package #:cl-excel.tests)

(def-suite :cl-excel/test-suite
  :description "Milestone M0 — skeleton & build tests.")
(in-suite :cl-excel/test-suite)

(test package-exists
  "The CL-EXCEL package is present."
  (is (find-package :cl-excel)))

(test exports-present
  "Key public API symbols are exported by the package."
  (flet ((exported-p (name)
           (multiple-value-bind (sym status)
               (find-symbol name :cl-excel)
             (declare (ignore sym))
             (eq status :external))))
    (dolist (nm '("READ-XLSX" "OPEN-XLSX" "CLOSE-XLSX" "WITH-XLSX" "WRITE-XLSX"
                  "SHEET-NAMES" "SHEET-COUNT" "HAS-SHEET-P" "SHEET" "ADD-SHEET!"
                  "RENAME-SHEET!" "GET-DATA" "READ-DATA" "GET-CELL" "GET-CELL-RANGE"
                  "CELL" "RANGE" "SHEET-REF" "NAMED" "EACH-ROW" "ROW-NUMBER"
                  "DO-ROWS" "EACH-TABLE-ROW" "DO-TABLE-ROWS"
                  "READ-TABLE" "GET-TABLE"
                  "TABLE" "TABLE-NAME" "TABLE-REF" "TABLE-COLUMNS" "TABLE-DISPLAY-NAME"
                  "TABLE-COLUMN" "TABLE-COLUMN-NAME" "TABLE-COLUMN-ID"
                  "XLSX-ERROR" "XLSX-PARSE-ERROR" "SHEET-MISSING-ERROR"
                  "INVALID-RANGE-ERROR" "READ-ONLY-ERROR"))
      (is (exported-p nm) "~A should be exported" nm))))

(defun run-tests ()
  "Entry point for ASDF test-op. Returns T when the suite is invoked."
  (let ((results (and (run! :cl-excel/test-suite)
                      (run! :cl-excel/refs)
                      (run! :cl-excel/workbook)
                      (run! :cl-excel/tables))))
    (declare (ignore results))
    t))

(defun fixture-path (filename)
  "Return the absolute path to a file in tests/fixtures/"
  (asdf:system-relative-pathname :cl-excel 
                                 (merge-pathnames filename "tests/fixtures/")))

