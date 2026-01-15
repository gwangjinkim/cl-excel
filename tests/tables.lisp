;;;; tests/tables.lisp

(in-package #:cl-excel.tests)

(def-suite :cl-excel/tables
  :description "Milestone M4 — Table parsing.")
(in-suite :cl-excel/tables)

(defparameter *sample-table-xml*
  "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"Table1\" displayName=\"Table1\" ref=\"A1:C5\" totalsRowShown=\"0\">
    <tableColumns count=\"3\">
        <tableColumn id=\"1\" name=\"Column1\"/>
        <tableColumn id=\"2\" name=\"Column2\"/>
        <tableColumn id=\"3\" name=\"Column3\"/>
    </tableColumns>
    <tableStyleInfo name=\"TableStyleMedium2\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/>
</table>")

(test parse-table-xml
  "Test parsing of a table definition from XML string."
  (let* ((stream (flexi-streams:make-in-memory-input-stream 
                  (flexi-streams:string-to-octets *sample-table-xml*)))
         (tbl (cl-excel::read-table-xml stream)))
    (is (not (null tbl)))
    (is (string= "Table1" (cl-excel:table-name tbl)))
    (is (string= "A1:C5" (cl-excel:table-ref tbl)))
    (is (= 3 (length (cl-excel:table-columns tbl))))
    
    (let ((c1 (first (cl-excel:table-columns tbl))))
      (is (string= "Column1" (cl-excel:table-column-name c1)))
      (is (= 1 (cl-excel:table-column-id c1))))))


(test read-table-data
  "Test reading data from a manually constructed table."
  (let* ((wb (make-instance 'cl-excel::workbook))
         (sh (make-instance 'cl-excel::sheet :name "Sheet1" :id 1))
         (tbl (make-instance 'cl-excel:table 
                             :name "Table1" 
                             :ref "A1:B2" 
                             :sheet sh
                             :columns (list (cl-excel::make-table-column :name "C1" :id 1)
                                            (cl-excel::make-table-column :name "C2" :id 2)))))
    ;; Setup Workbook/Sheet
    (setf (cl-excel::workbook-sheets wb) (list sh))
    (setf (cl-excel::workbook-tables wb) (list tbl))
    
    ;; Populate Cells
    (setf (cl-excel:get-cell sh "A1") (cl-excel::make-cell "C1"))
    (setf (cl-excel:get-cell sh "B1") (cl-excel::make-cell "C2"))
    (setf (cl-excel:get-cell sh "A2") (cl-excel::make-cell 10))
    (setf (cl-excel:get-cell sh "B2") (cl-excel::make-cell 20))
    
    ;; Read Table
    (let ((data (cl-excel:read-table wb "Table1")))
      (is (= 2 (length data)))
      (is (equal '("C1" "C2") (first data)))
      (is (equal '(10 20) (second data))))))
