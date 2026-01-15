;;;; tests/write-enhancements.lisp

(in-package #:cl-excel.tests)

(def-suite :cl-excel/write-enhancements
  :description "M12 — Enhanced write-xlsx logic.")
(in-suite :cl-excel/write-enhancements)

(test write-start-cell
  "Test writing data starting at a specific cell."
  (let ((data '(("val1" "val2") (10 20)))
        (path (uiop:with-temporary-file (:pathname p :suffix "xlsx") p)))
    ;; Write to B2
    (cl-excel:write-xlsx data path :start-cell "B2")
    
    ;; Verify
    (let ((wb (cl-excel:read-xlsx path)))
      (let ((sh (cl-excel:sheet wb 1)))
        (is (equal "val1" (cl-excel:get-data sh "B2")))
        (is (equal "val2" (cl-excel:get-data sh "C2")))
        (is (equal 10 (cl-excel:get-data sh "B3")))
        (is (equal 20 (cl-excel:get-data sh "C3")))
        ;; Check some boundaries
        (is (cl-excel:missing-p (cl-excel:get-data sh "A1")))
        (is (cl-excel:missing-p (cl-excel:get-data sh "B1")))
        (is (cl-excel:missing-p (cl-excel:get-data sh "A2")))))))

(test write-sheet-name-and-index
  "Test writing to a specific sheet by name or index."
  (let ((data '(("SheetData")))
        (path (uiop:with-temporary-file (:pathname p :suffix "xlsx") p)))
    
    ;; 1. Write to named sheet
    (cl-excel:write-xlsx data path :sheet "MySheet")
    (let ((wb (cl-excel:read-xlsx path)))
      (is (cl-excel:has-sheet-p wb "MySheet"))
      (is (equal "SheetData" (cl-excel:get-data (cl-excel:sheet wb "MySheet") "A1"))))
    
    ;; 2. Write to sheet index 2 (should create it if missing in a new wb? 
    ;; Actually, usually index refers to existing. Let's see how we implement it.)
    ))

(test write-region-clipping
  "Test region clipping logic."
  (let ((data '(("A1" "B1" "C1")
                ("A2" "B2" "C2")
                ("A3" "B3" "C3")))
        (path (uiop:with-temporary-file (:pathname p :suffix "xlsx") p)))
    
    ;; Write into B2:C3 region with clipping
    (cl-excel:write-xlsx data path :region "B2:C3" :region-only-p t)
    
    (let ((wb (cl-excel:read-xlsx path)))
      (let ((sh (cl-excel:sheet wb 1)))
        ;; Top-left of data "A1" goes into B2
        (is (equal "A1" (cl-excel:get-data sh "B2")))
        (is (equal "B1" (cl-excel:get-data sh "C2")))
        ;; "C1" should be clipped
        (is (cl-excel:missing-p (cl-excel:get-data sh "D2")))
        
        (is (equal "A2" (cl-excel:get-data sh "B3")))
        (is (equal "B2" (cl-excel:get-data sh "C3")))
        
        ;; Row 3 (data "A3"...) should be clipped
        (is (cl-excel:missing-p (cl-excel:get-data sh "B4")))))))

(test write-region-no-clipping
  "Test region without clipping (uses only start-cell behavior)."
  (let ((data '(("A1" "B1" "C1")
                ("A2" "B2" "C2")))
        (path (uiop:with-temporary-file (:pathname p :suffix "xlsx") p)))
    
    ;; Write into B2:B2 region without clipping -> should write full data from B2
    (cl-excel:write-xlsx data path :region "B2:B2" :region-only-p nil)
    
    (let ((wb (cl-excel:read-xlsx path)))
      (let ((sh (cl-excel:sheet wb 1)))
        (is (equal "A1" (cl-excel:get-data sh "B2")))
        (is (equal "B1" (cl-excel:get-data sh "C2")))
        (is (equal "C1" (cl-excel:get-data sh "D2")))
        (is (equal "A2" (cl-excel:get-data sh "B3")))))))
