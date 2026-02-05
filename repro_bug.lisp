(require :asdf)
(push (uiop:getcwd) asdf:*central-registry*)
(ql:quickload :cl-excel)

(defparameter *file* "repro_sparse.xlsx")

(format t "Writing to ~A...~%" *file*)
;; Write "C1" to cell C1. A1 and B1 should be empty.
;; Note: write-xlsx might not support sparse writing easily if we pass a list of lists.
;; But the README says: (write-xlsx data "output.xlsx" :start-cell "B2")

;; Let's try writing a single cell to C1.
(cl-excel:write-xlsx '(("C1 Value")) *file* :start-cell "C1" :sheet "SparseSheet")

(format t "Reading from ~A...~%" *file*)
(let ((rows (cl-excel:read-file *file* "SparseSheet")))
  (format t "Rows: ~S~%" rows)
  (if (and (= (length (first rows)) 3)
           (cl-excel::missing-p (first (first rows)))
           (cl-excel::missing-p (second (first rows))))
      (format t "SUCCESS: Sparse columns preserved.~%")
      (format t "FAILURE: Sparse columns NOT preserved. (Expected 3 cols, got ~A)~%" (length (first rows)))))

(uiop:quit)
