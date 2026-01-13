
(require :asdf)
(asdf:load-system :cl-excel)

(defun test-lazy ()
  (format t "Testing Lazy Iterator...~%")
  (let ((wb (cl-excel:read-xlsx "tests/fixtures/basic_types.xlsx")))
    (unwind-protect
         (cl-excel:with-sheet-iterator (next-row wb "Sheet1")
           (loop for row = (funcall next-row)
                 while row
                 do (format t "Row: ~S~%" row)))
      (cl-excel:close-xlsx wb))))

(test-lazy)
(sb-ext:exit)
