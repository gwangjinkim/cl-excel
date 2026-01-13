(require :asdf)
(asdf:load-system :cl-excel)

(format t "Testing Legacy DOM Read...~%")
(let* ((wb (cl-excel:read-xlsx "basic_types.xlsx"))
       (sheets (cl-excel:workbook-sheets wb))
       (sheet (first sheets)))
  (format t "Sheets: ~S~%" (length sheets))
  (format t "First Sheet Name: ~A~%" (cl-excel:sheet-name sheet))
  (format t "Cells count: ~A~%" (hash-table-count (cl-excel:sheet-cells sheet)))
  (let ((cell (gethash (cons 2 1) (cl-excel:sheet-cells sheet)))) ;; A2 "Hello"
    (format t "Cell A2: ~S~%" (if cell (cl-excel:cell-value cell) "NIL"))))

(format t "Legacy test done.~%")
(quit)
