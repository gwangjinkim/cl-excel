(in-package #:cl-excel.tests)

(def-suite :cl-excel/ods
  :description "Test ODS read/write support."
  :in :cl-excel/test-suite)
(in-suite :cl-excel/ods)

(test ods-roundtrip
  "Test that we can write data to ODS and read it back."
  (let ((path (fixture-path "test_roundtrip.ods"))
        (data '(("Sheet1" . (("Name" "Age") ("Alice" 30) ("Bob" 25))))))
    
    (when (probe-file path) (delete-file path))
    
    ;; 1. Write ODS
    (cl-excel:write-ods data path)
    (is (probe-file path))
    
    ;; 2. Read ODS generic
    (let ((wb (cl-excel:read-xlsx path))) ;; Auto-detects
      (is (typep wb 'cl-excel::workbook))
      (is (= 1 (cl-excel:sheet-count wb)))
      (let ((sh (cl-excel:sheet wb 1)))
        (is (string= "Sheet1" (cl-excel:sheet-name sh)))
        (is (string= "Name" (cl-excel:get-data sh "A1")))
        (is (= 30 (cl-excel:get-data sh "B2")))))
        
    (delete-file path)))

(test write-file-dispatch
  "Test write-file logic."
  (let ((path (fixture-path "test_dispatch.ods"))
        (data '(("S1" . ((1 2) (3 4))))))
    (when (probe-file path) (delete-file path))
    
    (cl-excel:write-file data path)
    (is (probe-file path))
    
    (let ((wb (cl-excel:read-file path)))
      (is (= 2 (length wb))) ;; Data has 2 rows: (1 2) and (3 4)
      ;; read-file returns just rows of the specified sheet.
      ;; Let's check using read-xlsx to be sure of structure
    )
    (let ((wb (cl-excel:read-xlsx path)))
      (is (string= "S1" (cl-excel:sheet-name (cl-excel:sheet wb 1)))))
      
    (delete-file path)))

(test ods-with-xlsx-extension
  "Test that ODS files are detected correctly even with .xlsx extension."
  (let* ((orig-path (fixture-path "test_ext.ods"))
         (new-path (fixture-path "test_ext.xlsx"))
         (data '(("Sheet1" . (("A" "B") (1 2))))))
    
    (when (probe-file orig-path) (delete-file orig-path))
    (when (probe-file new-path) (delete-file new-path))
    
    ;; 1. Write as ODS
    (cl-excel:write-ods data orig-path)
    
    ;; 2. Rename to .xlsx
    (rename-file orig-path new-path)
    (is (probe-file new-path))
    (is (not (probe-file orig-path)))
    
    ;; 3. Try to read with read-xlsx (which uses detect-file-format)
    (let ((wb (cl-excel:read-xlsx new-path)))
      (is (typep wb 'cl-excel::workbook))
      ;; Check internal flag or just verify it was read via ODS reader
      ;; Since we don't have a direct flag, we check if it was parsed correctly
      (is (= 1 (cl-excel:sheet-count wb)))
      (is (string= "Sheet1" (cl-excel:sheet-name (cl-excel:sheet wb 1))))
      (is (= 1 (cl-excel:get-data (cl-excel:sheet wb 1) "A2"))))
    
    ;; 4. Try read-file dispatch
    (let ((rows (cl-excel:read-file new-path)))
      (is (= 2 (length rows)))
      (is (equal '(1.0 2.0) (second rows))))
    
    (when (probe-file new-path) (delete-file new-path))))
