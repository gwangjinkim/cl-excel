;;;; tests/run.lisp — FiveAM harness (M0)

(defpackage #:cl-excel.tests
  (:use #:cl #:fiveam)
  (:export #:run-tests))

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
    (dolist (nm '("READXLSX" "OPENXLSX" "CLOSEXLSX" "WITH-XLSX" "WRITEXLSX"
                  "SHEETNAMES" "SHEETCOUNT" "HASSHEET" "SHEET" "ADDSHEET!"
                  "RENAME!" "GETDATA" "READDATA" "GETCELL" "GETCELLRANGE"
                  "CELL" "RANGE" "SHEETREF" "NAMED" "EACHROW" "ROW-NUMBER"
                  "EACHTABLEROW" "READTABLE" "GETTABLE" "DATATABLE-DATA"
                  "DATATABLE-COLUMN-LABELS" "DATATABLE-COLUMN-INDEX"
                  "WRITETABLE" "WRITETABLE!"))
      (is (exported-p nm) "~A should be exported" nm))))

(defun run-tests ()
  "Entry point for ASDF test-op. Returns T when the suite is invoked."
  (let ((results (run! :cl-excel/test-suite)))
    (declare (ignore results))
    t))

