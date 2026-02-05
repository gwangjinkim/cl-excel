;;;; tests/compat.lisp

(in-package #:cl-excel.tests)

(def-suite :cl-excel/compat
  :description "Compatibility and bug fix tests (M12+)"
  :in :cl-excel/test-suite)

(in-suite :cl-excel/compat)

(test sparse-columns-preserved
  "Ensure sparse columns return +missing+ instead of being collapsed."
  (let ((wb (make-instance 'cl-excel::workbook :zip nil :sheets nil))
        (sh (make-instance 'cl-excel::sheet :name "Sparse" :id 1 :rel-id "rId1")))
    
    ;; Setup sparse data manually in Hash Table: A1="A", C1="C"
    (setf (cl-excel:val sh "A1") "A")
    (setf (cl-excel:val sh "C1") "C")
    
    (let ((rows (cl-excel:map-rows #'identity sh)))
      (is (= 1 (length rows)))
      (let ((row (first rows)))
        (is (= 3 (length row)) "Row should have length 3 (A, B, C)")
        (is (string= "A" (first row)))
        (is (cl-excel:missing-p (second row)) "B1 should be +missing+")
        (is (string= "C" (third row)))))))

(test date-timezone-handling
  "Ensure dates are converted with timezone respect."
  ;; 44960 = 2023-02-03 (Verified 1899-12-30 base)
  (let ((val (cl-excel::excel-date-to-timestamp 44960 local-time:+utc-zone+)))
    (is (= 2023 (local-time:timestamp-year val :timezone local-time:+utc-zone+)))
    (is (= 2 (local-time:timestamp-month val :timezone local-time:+utc-zone+)))
    (is (= 3 (local-time:timestamp-day val :timezone local-time:+utc-zone+)))
    (is (= 0 (local-time:timestamp-hour val :timezone local-time:+utc-zone+)))))

(test workbook-metadata-slot
  "Ensure workbook has app-name slot."
  (let ((wb (make-instance 'cl-excel::workbook :zip nil :sheets nil :app-name "Apache POI")))
    (is (string= "Apache POI" (cl-excel:workbook-app-name wb)))))

;; Mocks for empty rows integration test would require full XML parsing.
;; We rely on manual verification or integration tests for that.
