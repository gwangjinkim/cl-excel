;;;; tests/workbook.lisp

(in-package #:cl-excel.tests)

(def-suite :cl-excel/workbook
  :description "Milestone M2 — Workbook loading.")
(in-suite :cl-excel/workbook)

(defparameter *fixture-path* 
  (fixture-path "basic_types.xlsx"))

(test read-basic-fixture
  "Open basic_types.xlsx and check sheets."
  (is (probe-file *fixture-path*) "Fixture file must exist")
  
  (let ((wb (cl-excel:read-xlsx *fixture-path*)))
    (is wb)
    (is (= 1 (cl-excel:sheet-count wb)))
    (is (string= "Sheet1" (first (cl-excel:sheet-names wb))))
    (is (cl-excel:has-sheet-p wb "Sheet1"))
    (is (not (cl-excel:has-sheet-p wb "NonExistent")))

    ;; M3 Verification: Check values (Row 2 data)
    (let ((sh (cl-excel:sheet wb 1)))
      (is sh)
      ;; E1: Header "Date"
      (is (equal "Date" (cl-excel:get-data sh "E1")))
      ;; F1: Header "DateTime"
      (is (equal "DateTime" (cl-excel:get-data sh "F1")))
      
      ;; Data Row (Row 2)
      ;; A2: "Hello"
      (is (equal "Hello" (cl-excel:get-data sh "A2")))
      ;; B2: 42
      (is (= 42 (cl-excel:get-data sh "B2")))
      ;; C2: Float 3.14
      (is (= 3.14 (cl-excel:get-data sh "C2")))
      ;; D2: Boolean True
      (is (eq t (cl-excel:get-data sh "D2")))
      ;; E2: Date (approx 2023-11-04)
      (let ((val (cl-excel:get-data sh "E2")))
        (is (typep val 'local-time:timestamp))
        (is (= 2023 (local-time:timestamp-year val))))
      ;; F2: DateTime
      (let ((val (cl-excel:get-data sh "F2")))
        (is (typep val 'local-time:timestamp)))
      ;; G2: Missing
      (is (cl-excel:missing-p (cl-excel:get-data sh "G2"))))
    
    (cl-excel:close-xlsx wb)))
