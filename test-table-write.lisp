(require :asdf)
(asdf:load-system :cl-excel)
(asdf:load-system :zip)

(defun test-write-table ()
  (format t "Testing Table Write...~%")
  (let ((wb (make-instance 'cl-excel::workbook))
        (sheet (make-instance 'cl-excel::sheet :name "DataSheet" :id 1)))
    
    (setf (cl-excel::workbook-sheets wb) (list sheet))
    
    ;; Add data and table
    (let ((data '(("Name" "Age" "Score")
                  ("Alice" 30 100)
                  ("Bob" 25 90))))
      (cl-excel:add-table! sheet data :name "StatsTable" :start-cell "B2"))
    
    ;; Write
    (let ((path "test_table.xlsx"))
      (when (probe-file path) (delete-file path))
      (cl-excel:write-xlsx wb path)
      (format t "Written to ~A~%" path)
      
      ;; Verify ZIP contents
      (zip:with-zipfile (zip path)
        (format t "Entries: ~S~%" (loop for k being the hash-keys of (zip:zipfile-entries zip) collect k))
        (unless (zip:get-zipfile-entry "xl/tables/table1.xml" zip)
          (error "Missing table1.xml"))
        (unless (zip:get-zipfile-entry "xl/worksheets/_rels/sheet1.xml.rels" zip)
          (error "Missing sheet1.xml.rels"))
        
        ;; Read back check
        (format t "Reading back...~%")
        (finish-output)
        ;; Note: We need to close zip before reading back with cl-excel if cl-excel opens it again? 
        ;; actually cl-excel uses zip:open-zipfile too.
        )
      
      ;; Read back using library
      (let* ((wb2 (cl-excel:read-xlsx path))
             (sh2 (cl-excel:sheet wb2 1))
             (tbls (cl-excel::sheet-tables sh2)))
        (format t "Read sheet: ~A~%" (cl-excel:sheet-name sh2))
        (format t "Tables found: ~D~%" (length tbls))
        (when tbls
          (format t "Table 1: ~A (Ref: ~A)~%" (cl-excel:table-name (first tbls)) (cl-excel:table-ref (first tbls)))
          (unless (string= (cl-excel:table-name (first tbls)) "StatsTable")
            (format t "FAILURE: Table name mismatch~%")))
        (cl-excel:close-xlsx wb2))
      
      (format t "SUCCESS.~%"))))

(test-write-table)
(sb-ext:exit)
