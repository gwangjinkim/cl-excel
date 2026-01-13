;;;; tests/refs.lisp

(in-package #:cl-excel.tests)

(def-suite :cl-excel/refs
  :description "Milestone M1 — Ref parsing and types.")
(in-suite :cl-excel/refs)

(test missing-values
  "Test +missing+ semantics."
  (is (eq cl-excel:+missing+ cl-excel:+missing+))
  (is (cl-excel:missing-p cl-excel:+missing+))
  (is (not (cl-excel:missing-p nil)))
  (is (not (cl-excel:missing-p "some value"))))

(test column-conversion
  "Test col-name and col-index."
  (is (string= "A" (cl-excel:col-name 1)))
  (is (string= "Z" (cl-excel:col-name 26)))
  (is (string= "AA" (cl-excel:col-name 27)))
  (is (string= "AB" (cl-excel:col-name 28)))
  (is (string= "XFD" (cl-excel:col-name 16384)))
  
  (is (= 1 (cl-excel:col-index "A")))
  (is (= 26 (cl-excel:col-index "Z")))
  (is (= 27 (cl-excel:col-index "AA")))
  (is (= 28 (cl-excel:col-index "AB")))
  (is (= 16384 (cl-excel:col-index "XFD")))
  (is (= 1 (cl-excel:col-index "a")) "Should be case-insensitive"))

(test parse-cell-refs
  "Test parsing of various cell reference formats."
  ;; Simple A1
  (let ((ref (cl-excel:parse-cell-ref "A1")))
    (is (= 1 (cl-excel::cell-ref-col ref)))
    (is (= 1 (cl-excel::cell-ref-row ref)))
    (is (not (cl-excel::cell-ref-sheet ref)))
    (is (not (cl-excel::cell-ref-abs-col ref)))
    (is (not (cl-excel::cell-ref-abs-row ref))))
  
  ;; Absolute $B$2
  (let ((ref (cl-excel:parse-cell-ref "$B$2")))
    (is (= 2 (cl-excel::cell-ref-col ref)))
    (is (= 2 (cl-excel::cell-ref-row ref)))
    (is (cl-excel::cell-ref-abs-col ref))
    (is (cl-excel::cell-ref-abs-row ref)))
  
  ;; Mixed B$5
  (let ((ref (cl-excel:parse-cell-ref "B$5")))
    (is (= 2 (cl-excel::cell-ref-col ref)))
    (is (= 5 (cl-excel::cell-ref-row ref)))
    (is (not (cl-excel::cell-ref-abs-col ref)))
    (is (cl-excel::cell-ref-abs-row ref)))

  ;; With Sheet
  (let ((ref (cl-excel:parse-cell-ref "Sheet1!C3")))
    (is (string= "Sheet1" (cl-excel::cell-ref-sheet ref)))
    (is (= 3 (cl-excel::cell-ref-col ref)))
    (is (= 3 (cl-excel::cell-ref-row ref)))))

(test ref-to-string
  "Test formatting of cell references."
  (is (string= "A1" (cl-excel:cell-ref-to-string 
                     (cl-excel::make-cell-ref 1 1))))
  (is (string= "$B$2" (cl-excel:cell-ref-to-string 
                       (cl-excel::make-cell-ref 2 2 :abs-row t :abs-col t))))
  (is (string= "Sheet1!C3" (cl-excel:cell-ref-to-string 
                            (cl-excel::make-cell-ref 3 3 :sheet "Sheet1")))))
