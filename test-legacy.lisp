(require :asdf)
(asdf:load-system :cl-excel)

(format t "Testing Legacy DOM Read...~%")
(let* ((wb (cl-excel:read-xlsx "tests/fixtures/basic_types.xlsx"))
       (sheet-count (cl-excel:sheet-count wb))
       (sheet (cl-excel:sheet wb 1)))
  (format t "Sheets: ~S~%" sheet-count)
  (format t "First Sheet Name: ~A~%" (cl-excel:sheet-name sheet))
  (format t "Cells count: ~A~%" (hash-table-count (cl-excel:sheet-cells sheet)))
  (let ((cell (gethash (cons 2 1) (cl-excel:sheet-cells sheet)))) ;; A2 "Hello"
    (format t "Cell A2: ~S~%" (if cell (cl-excel:cell-value cell) "NIL"))))

(format t "Legacy test done.~%")
(quit)
